#!/usr/bin/env python3
"""
Capture Optibus OTP JSON from Chrome and generate Planning/Scheduling reports.

This tool is intentionally schema-led instead of endpoint-led: it watches JSON
network responses and saves only payloads that look like the Planning OTP export
or Scheduling vehicle OTP payloads used by the existing report scripts.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import re
import secrets
import socket
import struct
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd


CHROME_APP = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def slug(value: str, limit: int = 80) -> str:
    out = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("_")
    return (out or "payload")[:limit]


def read_json_path(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as f:
        text = f.read()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return json.loads("{" + text + "}")


def iter_json_inputs(paths: Sequence[Path]) -> Iterable[Tuple[Path, Any]]:
    for path in paths:
        if path.is_dir():
            for child in sorted(path.glob("*.json")):
                yield child, read_json_path(child)
        elif path.suffix.lower() == ".zip":
            with zipfile.ZipFile(path, "r") as zf:
                for name in sorted(n for n in zf.namelist() if n.lower().endswith(".json")):
                    with zf.open(name) as f:
                        yield Path(f"{path}!{name}"), json.loads(f.read().decode("utf-8"))
        else:
            yield path, read_json_path(path)


def find_payload(root: Any, kind: str, depth: int = 0) -> Optional[Dict[str, Any]]:
    if depth > 5:
        return None
    if isinstance(root, dict):
        if kind == "planning" and looks_like_planning_payload(root):
            return root
        if kind == "schedule" and looks_like_schedule_payload(root):
            return root
        for key in ("data", "result", "payload", "response", "body"):
            found = find_payload(root.get(key), kind, depth + 1)
            if found:
                return found
        for value in root.values():
            if isinstance(value, (dict, list)):
                found = find_payload(value, kind, depth + 1)
                if found:
                    return found
    elif isinstance(root, list):
        for item in root[:25]:
            found = find_payload(item, kind, depth + 1)
            if found:
                return found
    return None


def looks_like_planning_payload(data: Dict[str, Any]) -> bool:
    timetables = data.get("timetables")
    if not isinstance(timetables, list) or "routes" not in data:
        return False
    if not timetables:
        return True
    for timetable in timetables[:10]:
        for direction in (timetable or {}).get("directions") or []:
            for trip in (direction or {}).get("trips") or []:
                if isinstance(trip, dict) and "otp" in trip:
                    return True
    return True


def looks_like_schedule_payload(data: Dict[str, Any]) -> bool:
    vehicles = data.get("vehicles")
    if not isinstance(vehicles, list):
        return False
    for vehicle in vehicles[:20]:
        if isinstance(vehicle, dict) and isinstance(vehicle.get("on_time_performance"), dict):
            return True
    return False


def payload_score(data: Dict[str, Any], kind: str) -> int:
    if kind == "planning":
        score = len(data.get("timetables") or []) * 10 + len(data.get("routes") or [])
        for timetable in data.get("timetables") or []:
            for direction in (timetable or {}).get("directions") or []:
                for trip in (direction or {}).get("trips") or []:
                    if isinstance(trip, dict) and trip.get("otp"):
                        score += 1
        return score
    if kind == "schedule":
        return sum(len((v or {}).get("on_time_performance") or {}) for v in data.get("vehicles") or [])
    return 0


def save_payload(data: Dict[str, Any], out_dir: Path, kind: str, url: str) -> Optional[Path]:
    ensure_dir(out_dir)
    canonical = json.dumps(data, sort_keys=True, separators=(",", ":"), default=str)
    digest = hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:12]
    base = slug(urllib.parse.urlparse(url).path.rsplit("/", 1)[-1] or kind)
    path = out_dir / f"{kind}_{base}_{digest}.json"
    if path.exists():
        return None
    with path.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return path


class CDPWebSocket:
    def __init__(self, ws_url: str):
        parsed = urllib.parse.urlparse(ws_url)
        self.host = parsed.hostname or "127.0.0.1"
        self.port = parsed.port or 80
        self.path = parsed.path + (("?" + parsed.query) if parsed.query else "")
        self.sock = socket.create_connection((self.host, self.port), timeout=10)
        self.next_id = 1
        self._handshake()

    def _handshake(self) -> None:
        key = base64.b64encode(secrets.token_bytes(16)).decode("ascii")
        request = (
            f"GET {self.path} HTTP/1.1\r\n"
            f"Host: {self.host}:{self.port}\r\n"
            "Upgrade: websocket\r\n"
            "Connection: Upgrade\r\n"
            f"Sec-WebSocket-Key: {key}\r\n"
            "Sec-WebSocket-Version: 13\r\n\r\n"
        )
        self.sock.sendall(request.encode("ascii"))
        response = self.sock.recv(4096)
        if b" 101 " not in response.split(b"\r\n", 1)[0]:
            raise RuntimeError(f"WebSocket handshake failed: {response[:200]!r}")

    def close(self) -> None:
        try:
            self.sock.close()
        except OSError:
            pass

    def send_json(self, message: Dict[str, Any]) -> None:
        raw = json.dumps(message, separators=(",", ":")).encode("utf-8")
        header = bytearray([0x81])
        if len(raw) < 126:
            header.append(0x80 | len(raw))
        elif len(raw) <= 0xFFFF:
            header.append(0x80 | 126)
            header.extend(struct.pack("!H", len(raw)))
        else:
            header.append(0x80 | 127)
            header.extend(struct.pack("!Q", len(raw)))
        mask = secrets.token_bytes(4)
        header.extend(mask)
        masked = bytes(b ^ mask[i % 4] for i, b in enumerate(raw))
        self.sock.sendall(bytes(header) + masked)

    def recv_json(self, timeout: Optional[float] = None) -> Dict[str, Any]:
        self.sock.settimeout(timeout)
        chunks: List[bytes] = []
        while True:
            b1, b2 = self._recv_exact(2)
            opcode = b1 & 0x0F
            masked = bool(b2 & 0x80)
            length = b2 & 0x7F
            if length == 126:
                length = struct.unpack("!H", self._recv_exact(2))[0]
            elif length == 127:
                length = struct.unpack("!Q", self._recv_exact(8))[0]
            mask = self._recv_exact(4) if masked else b""
            payload = self._recv_exact(length) if length else b""
            if masked:
                payload = bytes(b ^ mask[i % 4] for i, b in enumerate(payload))
            if opcode == 0x8:
                raise RuntimeError("Chrome closed the debugging socket")
            if opcode in (0x1, 0x0):
                chunks.append(payload)
                if b1 & 0x80:
                    return json.loads(b"".join(chunks).decode("utf-8"))

    def _recv_exact(self, n: int) -> bytes:
        out = bytearray()
        while len(out) < n:
            chunk = self.sock.recv(n - len(out))
            if not chunk:
                raise RuntimeError("Socket closed")
            out.extend(chunk)
        return bytes(out)

    def command(self, method: str, params: Optional[Dict[str, Any]] = None) -> int:
        msg_id = self.next_id
        self.next_id += 1
        message = {"id": msg_id, "method": method}
        if params is not None:
            message["params"] = params
        self.send_json(message)
        return msg_id


def http_json(url: str, timeout: float = 2.0) -> Any:
    with urllib.request.urlopen(url, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def launch_chrome(url: str, port: int, profile_dir: Path) -> subprocess.Popen:
    if not Path(CHROME_APP).exists():
        raise FileNotFoundError(f"Google Chrome was not found at {CHROME_APP}")
    ensure_dir(profile_dir)
    args = [
        CHROME_APP,
        f"--remote-debugging-port={port}",
        f"--user-data-dir={profile_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        url,
    ]
    return subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def wait_for_debugger(port: int, timeout: float = 20.0) -> None:
    deadline = time.time() + timeout
    last_error: Optional[Exception] = None
    while time.time() < deadline:
        try:
            http_json(f"http://127.0.0.1:{port}/json/version")
            return
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            time.sleep(0.25)
    raise RuntimeError(f"Chrome debugging endpoint did not start: {last_error}")


def select_page_ws(port: int, url_hint: str) -> str:
    pages = http_json(f"http://127.0.0.1:{port}/json")
    candidates = [p for p in pages if p.get("type") == "page" and p.get("webSocketDebuggerUrl")]
    if not candidates:
        raise RuntimeError("No debuggable Chrome page found.")
    host = urllib.parse.urlparse(url_hint).hostname or ""
    for page in candidates:
        if host and host in (page.get("url") or ""):
            return page["webSocketDebuggerUrl"]
    return candidates[0]["webSocketDebuggerUrl"]


def wait_for_command(cdp: CDPWebSocket, msg_id: int, timeout: float = 10.0) -> Dict[str, Any]:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            message = cdp.recv_json(timeout=max(0.1, deadline - time.time()))
        except socket.timeout:
            continue
        if message.get("id") == msg_id:
            return message
    raise TimeoutError(f"Timed out waiting for CDP command {msg_id}")


def capture_browser(args: argparse.Namespace) -> None:
    out_dir = Path(args.out_dir).expanduser().resolve()
    profile_dir = Path(args.profile_dir).expanduser().resolve()
    route_urls = read_lines(Path(args.route_urls).expanduser()) if args.route_urls else []
    process = launch_chrome(args.url, args.port, profile_dir)
    cdp: Optional[CDPWebSocket] = None
    saved = 0
    pending: Dict[str, str] = {}
    try:
        wait_for_debugger(args.port)
        cdp = CDPWebSocket(select_page_ws(args.port, args.url))
        wait_for_command(cdp, cdp.command("Network.enable", {"maxResourceBufferSize": 1024 * 1024 * 100}))
        wait_for_command(cdp, cdp.command("Page.enable"))
        if route_urls:
            print(f"Visiting {len(route_urls)} route URLs and recording matching JSON responses.")
            for index, route_url in enumerate(route_urls, start=1):
                print(f"[{index}/{len(route_urls)}] {route_url}")
                wait_for_command(cdp, cdp.command("Page.navigate", {"url": route_url}))
                saved += drain_network(cdp, pending, args.kind, out_dir, args.settle_seconds, args.url_regex)
        else:
            print("Chrome is open. Log in and navigate/click routes as usual.")
            print("Matching OTP JSON responses will be saved automatically. Press Ctrl+C here when done.")
            while True:
                saved += drain_network(cdp, pending, args.kind, out_dir, 1.0, args.url_regex)
    except KeyboardInterrupt:
        print("\nCapture stopped.")
    finally:
        if cdp:
            cdp.close()
        if not args.keep_chrome_open:
            process.terminate()
    print(f"Saved {saved} new {args.kind} payload(s) under {out_dir}")


def drain_network(
    cdp: CDPWebSocket,
    pending: Dict[str, str],
    kind: str,
    out_dir: Path,
    seconds: float,
    url_regex: Optional[str],
) -> int:
    saved = 0
    deadline = time.time() + seconds
    regex = re.compile(url_regex) if url_regex else None
    while time.time() < deadline:
        try:
            message = cdp.recv_json(timeout=max(0.1, deadline - time.time()))
        except socket.timeout:
            continue
        method = message.get("method")
        params = message.get("params") or {}
        if method == "Network.responseReceived":
            response = params.get("response") or {}
            url = response.get("url") or ""
            mime = (response.get("mimeType") or "").lower()
            if (regex and not regex.search(url)) or ("json" not in mime and not regex):
                continue
            pending[params["requestId"]] = url
        elif method == "Network.loadingFinished":
            request_id = params.get("requestId")
            url = pending.pop(request_id, None)
            if url:
                saved += fetch_and_maybe_save_body(cdp, request_id, url, kind, out_dir)
        elif method == "Network.loadingFailed":
            pending.pop(params.get("requestId"), None)
    return saved


def fetch_and_maybe_save_body(cdp: CDPWebSocket, request_id: str, url: str, kind: str, out_dir: Path) -> int:
    msg_id = cdp.command("Network.getResponseBody", {"requestId": request_id})
    try:
        response = wait_for_command(cdp, msg_id, timeout=8.0)
    except Exception:
        return 0
    result = response.get("result") or {}
    body = result.get("body")
    if not body:
        return 0
    if result.get("base64Encoded"):
        try:
            body = base64.b64decode(body).decode("utf-8")
        except Exception:
            return 0
    try:
        data = json.loads(body)
    except Exception:
        return 0
    payload = find_payload(data, kind)
    if not payload:
        return 0
    path = save_payload(payload, out_dir, kind, url)
    if path:
        print(f"Saved {kind} payload: {path.name} (score {payload_score(payload, kind)})")
        return 1
    return 0


def read_lines(path: Path) -> List[str]:
    with path.open("r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip() and not line.lstrip().startswith("#")]


def build_route_map(data: Dict[str, Any]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for route in data.get("routes", []) or []:
        route_id = (route or {}).get("_id")
        short = (route or {}).get("sign") or (route or {}).get("code") or (route or {}).get("name")
        if route_id and short is not None:
            out[str(route_id)] = str(short)
    return out


def build_service_map(data: Dict[str, Any]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for service in data.get("services", []) or []:
        service_id = (service or {}).get("id")
        name = (service or {}).get("name")
        if service_id and name is not None:
            out[str(service_id)] = str(name)
    return out


def merge_planning_payloads(payloads: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    if not payloads:
        raise ValueError("No Planning payloads found.")
    merged = dict(payloads[0])
    for key in ("routes", "services", "timetables"):
        seen = set()
        values = []
        for payload in payloads:
            for item in payload.get(key, []) or []:
                item_id = (item or {}).get("_id") or (item or {}).get("id")
                if key == "timetables":
                    item_id = ((item or {}).get("route_id"), (item or {}).get("service_id"))
                marker = json.dumps(item_id if item_id is not None else item, sort_keys=True, default=str)
                if marker not in seen:
                    seen.add(marker)
                    values.append(item)
        merged[key] = values
    return merged


def parse_force_list(value: str) -> set:
    forced = set()
    for chunk in (value or "").split(";"):
        if ":" in chunk:
            route, service = chunk.split(":", 1)
            forced.add((route.strip(), service.strip()))
    return forced


def extract_planning_records(
    data: Dict[str, Any],
    force_dirpos: bool = False,
    force_pairs: Optional[set] = None,
    dirpos_labels: bool = True,
) -> pd.DataFrame:
    route_map = build_route_map(data)
    service_map = build_service_map(data)
    force_pairs = force_pairs or set()
    rows = []
    for timetable in data.get("timetables", []) or []:
        route_id = str((timetable or {}).get("route_id"))
        route = route_map.get(route_id, route_id if route_id != "None" else "UNKNOWN")
        service_id = str((timetable or {}).get("service_id"))
        service = service_map.get(service_id, service_id if service_id != "None" else "UNKNOWN")
        use_dirpos = force_dirpos or (route, service) in force_pairs
        for dir_pos, direction in enumerate((timetable or {}).get("directions") or []):
            direction_label = str(dir_pos)
            if use_dirpos and dirpos_labels:
                direction_label = "Outbound" if dir_pos == 0 else ("Inbound" if dir_pos == 1 else str(dir_pos))
            for trip in (direction or {}).get("trips") or []:
                otp = (trip or {}).get("otp") or {}
                rows.append(
                    {
                        "route_id": route_id,
                        "route": route,
                        "service_id": service_id,
                        "service": service,
                        "direction": direction_label,
                        "trip_id": (trip or {}).get("trip_id"),
                        "trip_user_id": (trip or {}).get("trip_user_id"),
                        "pattern": (trip or {}).get("pattern"),
                        "departure_time": (trip or {}).get("departure_time"),
                        "otpScore": to_float(otp.get("otpScore")),
                        "otpLast": to_float(otp.get("otpLast")),
                        "confidenceLevel": to_float(otp.get("confidenceLevel")),
                        "timepointOtp": normalize_timepoint_list(otp.get("timepointOtp")),
                        "isNA": 1 if (otp.get("otpScore") is None or otp.get("otpScore") == "NA") else 0,
                    }
                )
    return pd.DataFrame(rows)


def to_float(value: Any) -> Optional[float]:
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def normalize_timepoint_list(value: Any) -> Optional[List[float]]:
    if not isinstance(value, list):
        return None
    out = []
    for item in value:
        number = to_float(item)
        if number is None:
            return None
        out.append(number)
    return out


def dedupe_planning(df: pd.DataFrame, include_direction: bool = False) -> pd.DataFrame:
    if df.empty:
        return df

    def key(row: pd.Series) -> Tuple[Any, ...]:
        if pd.notna(row.get("trip_id")) and str(row.get("trip_id")):
            base = ("trip", row["route_id"], row["service_id"], str(row["trip_id"]))
        else:
            base = ("patdep", row["route_id"], row["service_id"], str(row.get("pattern")), str(row.get("departure_time")))
        return base if not include_direction else base + (row["direction"],)

    out = df.copy()
    out["__key__"] = out.apply(key, axis=1)
    return out.loc[~out.duplicated("__key__", keep="first")].drop(columns="__key__")


def avg_timepoint(series: pd.Series) -> str:
    lists = [value for value in series if isinstance(value, list)]
    if not lists:
        return ""
    avgs = []
    for index in range(max(len(value) for value in lists)):
        nums = [value[index] for value in lists if len(value) > index and value[index] is not None]
        if nums:
            avgs.append(f"{sum(nums) / len(nums):.2f}")
    return ",".join(avgs)


def summarize_planning_df(df: pd.DataFrame, complete_grid: bool = False) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=[
                "route",
                "direction",
                "service",
                "avg_otpScore",
                "avg_otpLast",
                "avg_confidenceLevel",
                "avg_timepointOtp",
                "trips_with_NA",
                "total_trips",
                "valid_trips",
            ]
        )
    base = (
        df.groupby(["route", "direction", "service"], dropna=False)
        .agg(
            avg_otpScore=("otpScore", "mean"),
            avg_otpLast=("otpLast", "mean"),
            avg_confidenceLevel=("confidenceLevel", "mean"),
            trips_with_NA=("isNA", "sum"),
            total_trips=("isNA", "count"),
            valid_trips=("otpScore", "count"),
        )
        .round(2)
    )
    summary = base.join(df.groupby(["route", "direction", "service"], dropna=False).agg(avg_timepointOtp=("timepointOtp", avg_timepoint)))
    if complete_grid:
        idx = pd.MultiIndex.from_product(
            [
                sorted(df["route"].dropna().unique().tolist()),
                sorted(df["direction"].dropna().unique().tolist()),
                sorted(df["service"].dropna().unique().tolist()),
            ],
            names=["route", "direction", "service"],
        )
        summary = summary.reindex(idx)
        for column in ("trips_with_NA", "total_trips", "valid_trips"):
            summary[column] = summary[column].fillna(0).astype(int)
        summary["avg_timepointOtp"] = summary["avg_timepointOtp"].fillna("")
    return summary.reset_index().sort_values(["route", "direction", "service"], kind="mergesort")


def write_table(df: pd.DataFrame, csv_path: Optional[str], xlsx_path: Optional[str], sheet_name: str) -> None:
    if csv_path:
        path = Path(csv_path).expanduser()
        ensure_dir(path.parent)
        df.to_csv(path, index=False)
        print(f"Saved CSV: {path}")
    if xlsx_path:
        path = Path(xlsx_path).expanduser()
        ensure_dir(path.parent)
        try:
            with pd.ExcelWriter(path) as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name)
            print(f"Saved Excel: {path}")
        except ImportError:
            print("No Excel writer engine is installed. Install openpyxl or xlsxwriter, or use --out-csv only.", file=sys.stderr)


def summarize_planning(args: argparse.Namespace) -> None:
    payloads = []
    for _, data in iter_json_inputs([Path(p).expanduser() for p in args.input]):
        payload = find_payload(data, "planning")
        if payload:
            payloads.append(payload)
    merged = merge_planning_payloads(payloads)
    if args.merged_json:
        path = Path(args.merged_json).expanduser()
        ensure_dir(path.parent)
        with path.open("w", encoding="utf-8") as f:
            json.dump(merged, f, ensure_ascii=False, indent=2)
        print(f"Saved merged Planning JSON: {path}")
    detail = extract_planning_records(
        merged,
        force_dirpos=args.force_dirpos or bool(args.force_dirpos_list),
        force_pairs=parse_force_list(args.force_dirpos_list),
        dirpos_labels=not args.no_dirpos_labels,
    )
    detail = dedupe_planning(detail, include_direction=args.dedupe_include_direction)
    summary = summarize_planning_df(detail, complete_grid=args.complete_grid)
    write_table(summary, args.out_csv, args.out_xlsx, "OTP Summary")
    if args.detail_csv:
        out = detail.copy()
        out["timepointOtp"] = out["timepointOtp"].apply(lambda v: "[" + ",".join(f"{x:.2f}" for x in v) + "]" if isinstance(v, list) else "")
        write_table(out, args.detail_csv, None, "Trip Detail")
    print(summary.head(20).to_string(index=False))


def summarize_schedule(args: argparse.Namespace) -> None:
    props_path = Path(args.trip_properties).expanduser()
    if not props_path.exists():
        raise FileNotFoundError(f"Trip properties file not found: {props_path}")
    props = pd.read_excel(props_path)
    if "System Id" not in props.columns:
        raise ValueError("'System Id' column not found in the trip properties file.")
    props["System Id"] = props["System Id"].astype(str).str.strip()
    trip_properties = props.set_index("System Id").to_dict(orient="index")
    rows = []
    for path, data in iter_json_inputs([Path(p).expanduser() for p in args.input]):
        payload = find_payload(data, "schedule")
        if not payload:
            print(f"Skipped non-Scheduling payload: {path}")
            continue
        for vehicle in payload.get("vehicles", []) or []:
            for trip_id, stats in ((vehicle or {}).get("on_time_performance") or {}).items():
                clean_trip_id = str(trip_id).strip()
                row = {
                    "id": clean_trip_id,
                    "start_otp": (stats or {}).get("start_otp", ""),
                    "end_otp": (stats or {}).get("end_otp", ""),
                    "predicted_otp": (stats or {}).get("predicted_otp", ""),
                }
                if clean_trip_id in trip_properties:
                    row.update(trip_properties[clean_trip_id])
                rows.append(row)
    if not rows:
        raise ValueError("No trips with Scheduling vehicle OTP data were found.")
    headers = ["id", "start_otp", "end_otp", "predicted_otp", "Service", "Route", "Route Code", "Direction", "Pattern", "Start Time"]
    output = pd.DataFrame(rows).reindex(columns=headers)
    write_table(output, args.out_csv, args.out_xlsx, "Schedule OTP")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Capture Optibus OTP JSON and generate OTP reports.")
    sub = parser.add_subparsers(dest="command", required=True)

    capture = sub.add_parser("capture", help="Open Chrome, capture matching OTP JSON responses, and save them.")
    capture.add_argument("--url", required=True, help="Optibus page URL to open.")
    capture.add_argument("--kind", choices=["planning", "schedule"], required=True)
    capture.add_argument("--out-dir", default="otp_captures")
    capture.add_argument("--route-urls", help="Optional text file of route URLs to visit, one per line.")
    capture.add_argument("--url-regex", help="Optional regex to restrict which response URLs are inspected.")
    capture.add_argument("--settle-seconds", type=float, default=8.0, help="Seconds to wait after each route URL.")
    capture.add_argument("--port", type=int, default=9222)
    capture.add_argument("--profile-dir", default=".chrome-otp-profile")
    capture.add_argument("--keep-chrome-open", action="store_true")
    capture.set_defaults(func=capture_browser)

    planning = sub.add_parser("planning", help="Generate Planning OTP summary from captured/exported JSON or ZIP.")
    planning.add_argument("--input", nargs="+", required=True, help="JSON/ZIP files or folders containing captured JSON.")
    planning.add_argument("--out-csv", default="otp_planning_summary.csv")
    planning.add_argument("--out-xlsx")
    planning.add_argument("--detail-csv")
    planning.add_argument("--merged-json", help="Optional path to save merged Planning JSON.")
    planning.add_argument("--complete-grid", action="store_true")
    planning.add_argument("--force-dirpos", action="store_true")
    planning.add_argument("--force-dirpos-list", default="")
    planning.add_argument("--no-dirpos-labels", action="store_true")
    planning.add_argument("--dedupe-include-direction", action="store_true")
    planning.set_defaults(func=summarize_planning)

    schedule = sub.add_parser("schedule", help="Generate Scheduling OTP report from captured JSON.")
    schedule.add_argument("--input", nargs="+", required=True, help="JSON files or folders containing captured JSON.")
    schedule.add_argument("--trip-properties", required=True, help="Trip Properties Excel file with System Id.")
    schedule.add_argument("--out-csv")
    schedule.add_argument("--out-xlsx", default="otp_schedule_report.xlsx")
    schedule.set_defaults(func=summarize_schedule)

    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        args.func(args)
        return 0
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
