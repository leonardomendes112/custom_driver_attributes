from __future__ import annotations

import json
import re
from datetime import date
from io import BytesIO
from typing import Any, Dict, Iterable, List, Sequence, Tuple

import pandas as pd
import requests
import streamlit as st


TEMPLATE_VERSION = "driver_custom_attributes_v1"
TEMPLATE_COLUMNS = [
    "driver_id",
    "attribute_id",
    "value",
    "value_type",
    "start_date",
    "end_date",
    "entry_id",
    "notes",
]
MAX_API_BATCH_SIZE = 1000
ENDPOINT_PATH = "/v2/drivers/custom-attributes"
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
DURATION_RE = re.compile(r"^\d{1,3}:[0-5]\d$")
ENTRY_ERROR_RE = re.compile(r"entries\[(\d+)\]")


class TemplateError(ValueError):
    pass


def template_dataframe() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "driver_id": "driver-123",
                "attribute_id": "color",
                "value": "red",
                "value_type": "string",
                "start_date": "2026-01-01",
                "end_date": "2026-01-31",
                "entry_id": "",
                "notes": "Add a new value period.",
            },
            {
                "driver_id": "driver-123",
                "attribute_id": "skills",
                "value": "route_a;route_b",
                "value_type": "multi_select",
                "start_date": "2026-02-01",
                "end_date": "",
                "entry_id": "",
                "notes": "Use semicolons or a JSON array for multi-select values.",
            },
            {
                "driver_id": "driver-456",
                "attribute_id": "available_for_overtime",
                "value": "true",
                "value_type": "boolean",
                "start_date": "2026-01-01",
                "end_date": "",
                "entry_id": "",
                "notes": "Boolean values accept true/false, yes/no, or 1/0.",
            },
            {
                "driver_id": "driver-456",
                "attribute_id": "temporary_badge",
                "value": "",
                "value_type": "unset",
                "start_date": "2026-03-01",
                "end_date": "2026-03-31",
                "entry_id": "existing-entry-id",
                "notes": "In edit mode, unset a non-mandatory value by omitting value.",
            },
        ],
        columns=TEMPLATE_COLUMNS,
    )


def dataframe_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def dataframe_excel_bytes(df: pd.DataFrame, sheet_name: str = "Template") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        worksheet = writer.sheets[sheet_name[:31]]
        for index, column in enumerate(df.columns):
            width = min(max(len(str(column)), 14), 36)
            worksheet.set_column(index, index, width)
    return output.getvalue()


def normalize_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {column: str(column).strip().lower().replace(" ", "_") for column in df.columns}
    normalized = df.rename(columns=renamed).copy()
    for column in TEMPLATE_COLUMNS:
        if column not in normalized.columns:
            normalized[column] = ""
    return normalized[TEMPLATE_COLUMNS]


def read_template(uploaded_file: Any) -> pd.DataFrame:
    suffix = uploaded_file.name.rsplit(".", 1)[-1].lower()
    raw = uploaded_file.getvalue()
    if suffix == "csv":
        df = pd.read_csv(BytesIO(raw), dtype=str, keep_default_na=False)
    elif suffix in {"xlsx", "xls"}:
        df = pd.read_excel(BytesIO(raw), dtype=str).fillna("")
    else:
        raise TemplateError("Upload a CSV or Excel file.")
    return normalize_dataframe_columns(df)


def require_date(value: str, field_name: str, row_number: int, *, allow_blank: bool = False) -> str | None:
    cleaned = str(value or "").strip()
    if not cleaned:
        if allow_blank:
            return None
        raise TemplateError(f"Row {row_number}: {field_name} is required.")
    if not DATE_RE.match(cleaned):
        raise TemplateError(f"Row {row_number}: {field_name} must be formatted as YYYY-MM-DD.")
    try:
        date.fromisoformat(cleaned)
    except ValueError as exc:
        raise TemplateError(f"Row {row_number}: {field_name} is not a valid calendar date.") from exc
    return cleaned


def parse_bool(value: str, row_number: int) -> bool:
    normalized = str(value).strip().lower()
    if normalized in {"true", "t", "yes", "y", "1"}:
        return True
    if normalized in {"false", "f", "no", "n", "0"}:
        return False
    raise TemplateError(f"Row {row_number}: boolean value must be true/false, yes/no, or 1/0.")


def parse_number(value: str, row_number: int) -> int | float:
    cleaned = str(value).strip()
    if not cleaned:
        raise TemplateError(f"Row {row_number}: number value cannot be blank.")
    try:
        number = float(cleaned)
    except ValueError as exc:
        raise TemplateError(f"Row {row_number}: number value is invalid.") from exc
    return int(number) if number.is_integer() else number


def parse_multi_select(value: str, row_number: int) -> List[str | int | float]:
    cleaned = str(value or "").strip()
    if not cleaned:
        return []
    if cleaned.startswith("["):
        try:
            parsed = json.loads(cleaned)
        except json.JSONDecodeError as exc:
            raise TemplateError(f"Row {row_number}: multi_select JSON array is invalid.") from exc
        if not isinstance(parsed, list) or not all(isinstance(item, (str, int, float)) for item in parsed):
            raise TemplateError(f"Row {row_number}: multi_select JSON must be an array of strings or numbers.")
        return parsed
    return [item.strip() for item in cleaned.split(";") if item.strip()]


def parse_attribute_value(value: str, value_type: str, row_number: int) -> Tuple[bool, Any]:
    normalized_type = str(value_type or "string").strip().lower()
    cleaned = str(value or "").strip()
    if normalized_type in {"unset", "blank", "none", "null"}:
        return False, None
    if normalized_type == "auto":
        if not cleaned:
            return False, None
        try:
            return True, json.loads(cleaned)
        except json.JSONDecodeError:
            return True, cleaned
    if normalized_type in {"string", "single_select"}:
        if not cleaned:
            raise TemplateError(f"Row {row_number}: {normalized_type} value cannot be blank. Use value_type=unset to clear.")
        return True, cleaned
    if normalized_type == "number":
        return True, parse_number(cleaned, row_number)
    if normalized_type == "boolean":
        return True, parse_bool(cleaned, row_number)
    if normalized_type == "date":
        return True, require_date(cleaned, "value", row_number)
    if normalized_type == "duration":
        if not DURATION_RE.match(cleaned):
            raise TemplateError(f"Row {row_number}: duration must be formatted as HH:mm, with HH from 0 to 999.")
        return True, cleaned
    if normalized_type in {"multi_select", "array"}:
        return True, parse_multi_select(cleaned, row_number)
    if normalized_type == "json":
        if not cleaned:
            raise TemplateError(f"Row {row_number}: JSON value cannot be blank.")
        try:
            return True, json.loads(cleaned)
        except json.JSONDecodeError as exc:
            raise TemplateError(f"Row {row_number}: JSON value is invalid.") from exc
    raise TemplateError(
        f"Row {row_number}: unknown value_type '{value_type}'. "
        "Use string, number, boolean, date, duration, single_select, multi_select, json, auto, or unset."
    )


def records_to_entries(df: pd.DataFrame, *, strip_entry_ids: bool) -> List[Dict[str, Any]]:
    entries: List[Dict[str, Any]] = []
    for index, row in df.iterrows():
        row_number = index + 2
        driver_id = str(row["driver_id"] or "").strip()
        attribute_id = str(row["attribute_id"] or "").strip()
        if not driver_id:
            raise TemplateError(f"Row {row_number}: driver_id is required.")
        if not attribute_id:
            raise TemplateError(f"Row {row_number}: attribute_id is required.")

        start_date = require_date(str(row["start_date"]), "start_date", row_number)
        end_date = require_date(str(row["end_date"]), "end_date", row_number, allow_blank=True)
        if end_date and start_date and date.fromisoformat(end_date) < date.fromisoformat(start_date):
            raise TemplateError(f"Row {row_number}: end_date cannot be before start_date.")

        include_value, parsed_value = parse_attribute_value(str(row["value"]), str(row["value_type"]), row_number)
        entry: Dict[str, Any] = {
            "driverId": driver_id,
            "attributeId": attribute_id,
            "startDate": start_date,
        }
        if end_date:
            entry["endDate"] = end_date
        entry_id = str(row["entry_id"] or "").strip()
        if entry_id and not strip_entry_ids:
            if len(entry_id) > 36:
                raise TemplateError(f"Row {row_number}: entry_id cannot exceed 36 characters.")
            entry["entryId"] = entry_id
        if include_value:
            entry["value"] = parsed_value
        entries.append(entry)
    return entries


def entries_to_payload(entries: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    return {"entries": list(entries)}


def chunk_entries(entries: Sequence[Dict[str, Any]], batch_size: int) -> Iterable[List[Dict[str, Any]]]:
    for start in range(0, len(entries), batch_size):
        yield list(entries[start : start + batch_size])


def api_headers(api_key: str, account_name: str) -> Dict[str, str]:
    return {
        "Authorization": api_key.strip(),
        "X-Optibus-Api-Client": account_name.strip(),
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def endpoint_url(base_url: str) -> str:
    base = base_url.strip().rstrip("/")
    if base.endswith(ENDPOINT_PATH):
        return base
    return f"{base}{ENDPOINT_PATH}"


def parse_csv_filter(value: str) -> List[str]:
    normalized = str(value or "").replace("\n", ",").replace(";", ",")
    return [item.strip() for item in normalized.split(",") if item.strip()]


def fetch_custom_attributes(
    *,
    base_url: str,
    api_key: str,
    account_name: str,
    driver_ids: Sequence[str],
    attribute_ids: Sequence[str],
    from_date: str | None,
    to_date: str | None,
    timeout: int,
) -> List[Dict[str, Any]]:
    url = endpoint_url(base_url)
    headers = api_headers(api_key, account_name)
    page = 1
    entries: List[Dict[str, Any]] = []

    while True:
        params: Dict[str, Any] = {"page": page}
        if driver_ids:
            params["driverIds"] = ",".join(driver_ids)
        if attribute_ids:
            params["attributeIds"] = ",".join(attribute_ids)
        if from_date:
            params["fromDate"] = from_date
        if to_date:
            params["toDate"] = to_date

        response = requests.get(url, headers=headers, params=params, timeout=timeout)
        if not response.ok:
            raise RuntimeError(format_api_error(response))

        payload = response.json()
        entries.extend(payload.get("entries", []) or [])
        pagination = payload.get("pagination", {}) or {}
        next_page = pagination.get("nextPage")
        total_pages = pagination.get("totalPages")
        if next_page:
            page = int(next_page)
        elif total_pages and page < int(total_pages):
            page += 1
        else:
            break

    return entries


def put_custom_attribute_batches(
    *,
    base_url: str,
    api_key: str,
    account_name: str,
    entries: Sequence[Dict[str, Any]],
    batch_size: int,
    timeout: int,
) -> List[Dict[str, Any]]:
    url = endpoint_url(base_url)
    headers = api_headers(api_key, account_name)
    responses: List[Dict[str, Any]] = []
    for index, batch in enumerate(chunk_entries(entries, batch_size), start=1):
        response = requests.put(url, headers=headers, json=entries_to_payload(batch), timeout=timeout)
        if not response.ok:
            raise RuntimeError(f"Batch {index} failed: {format_api_error(response, batch)}")
        responses.append(response.json() if response.content else {})
    return responses


def format_api_error(response: requests.Response, batch: Sequence[Dict[str, Any]] | None = None) -> str:
    try:
        payload = response.json()
    except ValueError:
        payload = response.text
    entry_context = format_entry_error_context(payload, batch or [])
    return f"HTTP {response.status_code}: {payload}{entry_context}"


def format_entry_error_context(payload: Any, batch: Sequence[Dict[str, Any]]) -> str:
    if not batch:
        return ""
    matches = ENTRY_ERROR_RE.findall(str(payload))
    details = []
    for match in matches:
        index = int(match)
        if 0 <= index < len(batch):
            entry = batch[index]
            details.append(
                f"entries[{index}] is driverId={entry.get('driverId')}, "
                f"attributeId={entry.get('attributeId')}, startDate={entry.get('startDate')}"
            )
    return " | " + "; ".join(details) if details else ""


def fetched_entries_to_template(entries: Sequence[Dict[str, Any]]) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    for entry in entries:
        value = entry.get("value", "")
        value_type = infer_value_type(value, "unset" if "value" not in entry else "string")
        if isinstance(value, list):
            rendered_value = json.dumps(value, ensure_ascii=False)
        elif isinstance(value, bool):
            rendered_value = "true" if value else "false"
        else:
            rendered_value = "" if value is None else str(value)
        rows.append(
            {
                "driver_id": entry.get("driverId", ""),
                "attribute_id": entry.get("attributeId", ""),
                "value": rendered_value,
                "value_type": value_type,
                "start_date": entry.get("startDate", ""),
                "end_date": entry.get("endDate", ""),
                "entry_id": entry.get("entryId", ""),
                "notes": "",
            }
        )
    return pd.DataFrame(rows, columns=TEMPLATE_COLUMNS)


def infer_value_type(value: Any, default: str) -> str:
    if isinstance(value, bool):
        return "boolean"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, list):
        return "multi_select"
    return default


def clean_entries_from_fetched(
    entries: Sequence[Dict[str, Any]],
    *,
    skip_attribute_ids: Sequence[str] = (),
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    clean_entries: List[Dict[str, Any]] = []
    skipped_entries: List[Dict[str, Any]] = []
    skip_set = set(skip_attribute_ids)
    for entry in entries:
        if entry.get("attributeId") in skip_set:
            skipped_entries.append(entry)
            continue
        clean_entry = {
            "entryId": entry.get("entryId"),
            "driverId": entry.get("driverId"),
            "attributeId": entry.get("attributeId"),
            "startDate": entry.get("startDate"),
        }
        if entry.get("endDate"):
            clean_entry["endDate"] = entry.get("endDate")
        clean_entries.append({key: value for key, value in clean_entry.items() if value})
    return clean_entries, skipped_entries


def summarize_entries_by_attribute(entries: Sequence[Dict[str, Any]]) -> pd.DataFrame:
    rows = [{"attribute_id": entry.get("attributeId", ""), "entries": 1} for entry in entries]
    if not rows:
        return pd.DataFrame(columns=["attribute_id", "entries"])
    return (
        pd.DataFrame(rows)
        .groupby("attribute_id", as_index=False)["entries"]
        .sum()
        .sort_values(["entries", "attribute_id"], ascending=[False, True])
    )


def clean_payload_dataframe(entries: Sequence[Dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for index, entry in enumerate(entries):
        rows.append(
            {
                "payload_index": index,
                "entry_id": entry.get("entryId", ""),
                "driver_id": entry.get("driverId", ""),
                "attribute_id": entry.get("attributeId", ""),
                "start_date": entry.get("startDate", ""),
                "end_date": entry.get("endDate", ""),
            }
        )
    return pd.DataFrame(rows)


def api_credentials_ready(base_url: str, api_key: str, account_name: str) -> bool:
    return bool(base_url.strip() and api_key.strip() and account_name.strip())


def render_connection_controls() -> Tuple[str, str, str, int, int]:
    st.subheader("API Connection")
    col1, col2 = st.columns(2)
    with col1:
        base_url = st.text_input(
            "Base URL",
            placeholder="https://YOUR-OPTIBUS-ACCOUNT.api.ops.optibus.co",
            help="Paste either the account base URL or the full /v2/drivers/custom-attributes endpoint.",
        )
        account_name = st.text_input("X-Optibus-Api-Client account name")
    with col2:
        api_key = st.text_input("Authorization API key", type="password")
        timeout = st.number_input("Request timeout seconds", min_value=5, max_value=300, value=60, step=5)
    batch_size = st.slider("PUT batch size", min_value=1, max_value=MAX_API_BATCH_SIZE, value=MAX_API_BATCH_SIZE)
    return base_url, api_key, account_name, int(timeout), int(batch_size)


def render_template_download() -> None:
    st.subheader("Template")
    st.caption(
        f"Template version: {TEMPLATE_VERSION}. Required columns are driver_id, attribute_id, value_type, "
        "start_date, and value unless value_type is unset. end_date and entry_id are optional."
    )
    template = template_dataframe()
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "Download CSV template",
            dataframe_csv_bytes(template),
            file_name=f"{TEMPLATE_VERSION}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col2:
        st.download_button(
            "Download Excel template",
            dataframe_excel_bytes(template),
            file_name=f"{TEMPLATE_VERSION}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def render_fetch_export(base_url: str, api_key: str, account_name: str, timeout: int) -> None:
    with st.expander("Fetch and export current attributes", expanded=False):
        st.caption(
            "Use this first when changing existing driver attributes. Export the current entries, keep the periods "
            "that should remain unchanged, then edit only the values/dates that should change."
        )
        try:
            filters = render_filter_controls("fetch")
        except Exception as exc:  # noqa: BLE001
            st.error(str(exc))
            return
        if st.button(
            "Fetch Current Attributes",
            disabled=not api_credentials_ready(base_url, api_key, account_name),
            key="fetch_current_attributes",
        ):
            try:
                entries = fetch_custom_attributes(
                    base_url=base_url,
                    api_key=api_key,
                    account_name=account_name,
                    driver_ids=filters["driver_ids"],
                    attribute_ids=filters["attribute_ids"],
                    from_date=filters["from_date"],
                    to_date=filters["to_date"],
                    timeout=timeout,
                )
            except Exception as exc:  # noqa: BLE001
                st.error(str(exc))
                return
            st.success(f"Fetched {len(entries):,} entries.")
            export_df = fetched_entries_to_template(entries)
            st.dataframe(export_df, use_container_width=True, hide_index=True)
            st.download_button(
                "Download fetched attributes as CSV",
                dataframe_csv_bytes(export_df),
                file_name="driver_custom_attributes_current.csv",
                mime="text/csv",
                use_container_width=True,
            )


def render_filter_controls(key_prefix: str) -> Dict[str, Any]:
    today = date.today()
    col1, col2 = st.columns(2)
    with col1:
        driver_ids = st.text_area(
            "Driver IDs filter",
            placeholder="driver-123, driver-456",
            key=f"{key_prefix}_driver_ids",
            help="Comma-separated. Leave blank only when you intentionally want a broad query.",
        )
        from_date = st.date_input(
            "From date",
            value=today,
            key=f"{key_prefix}_from_date",
            help="Only entries that end on or after this date are returned. Defaults to today.",
        )
    with col2:
        attribute_ids = st.text_area(
            "Attribute IDs filter",
            placeholder="color, skills",
            key=f"{key_prefix}_attribute_ids",
        )
        to_date = st.date_input(
            "To date",
            value=today,
            key=f"{key_prefix}_to_date",
            help="Only entries that start on or before this date are returned. Defaults to today.",
        )

    if to_date < from_date:
        raise TemplateError("to_date cannot be before from_date.")
    return {
        "driver_ids": parse_csv_filter(driver_ids),
        "attribute_ids": parse_csv_filter(attribute_ids),
        "from_date": from_date.isoformat(),
        "to_date": to_date.isoformat(),
    }


def render_import_mode(
    *,
    mode: str,
    base_url: str,
    api_key: str,
    account_name: str,
    timeout: int,
    batch_size: int,
) -> None:
    is_add_mode = mode.startswith("Append")
    st.subheader(mode)
    if is_add_mode:
        st.caption(
            "Creates new entries by stripping entry_id values. Existing entries are not sent in the payload, but the "
            "resulting timeline for each driver + attribute must still be valid. Fetch existing entries first if you "
            "need to preserve or adjust surrounding periods."
        )
    else:
        st.caption(
            "Sends entry_id values when present. The API updates matching entries and creates missing entry IDs. "
            "For safe edits, start from a fetched export so unchanged entries for the affected driver + attribute "
            "remain in the template."
        )
    st.info(
        "Important: /v2/drivers/custom-attributes is entry-based, not a full driver replacement. However, Optibus "
        "validates the full timeline for each affected driver + attribute: no overlaps, no gaps, and exactly one "
        "open-ended entry. Preserve unchanged timeline entries when editing existing attributes."
    )

    uploaded = st.file_uploader(
        "Upload completed template",
        type=["csv", "xlsx", "xls"],
        key=f"{mode}_upload",
    )
    dry_run = st.checkbox("Dry run only", value=True, key=f"{mode}_dry_run")

    if uploaded is None:
        return

    try:
        df = read_template(uploaded)
        entries = records_to_entries(df, strip_entry_ids=is_add_mode)
    except Exception as exc:  # noqa: BLE001
        st.error(str(exc))
        return

    st.success(f"Parsed {len(entries):,} entries.")
    st.dataframe(df, use_container_width=True, hide_index=True)

    payload = entries_to_payload(entries)
    with st.expander("Payload preview", expanded=False):
        st.json(payload)
        st.download_button(
            "Download JSON payload",
            json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name="driver_custom_attributes_payload.json",
            mime="application/json",
        )

    disabled = dry_run or not entries or not api_credentials_ready(base_url, api_key, account_name)
    action_label = "Dry run complete" if dry_run else "Run PUT Import"
    if st.button(action_label, type="primary", disabled=disabled, key=f"{mode}_run"):
        try:
            responses = put_custom_attribute_batches(
                base_url=base_url,
                api_key=api_key,
                account_name=account_name,
                entries=entries,
                batch_size=batch_size,
                timeout=timeout,
            )
        except Exception as exc:  # noqa: BLE001
            st.error(str(exc))
            return
        st.success(f"Imported {len(entries):,} entries across {len(responses):,} API batch(es).")
        st.json(responses)


def render_clean_mode(
    *,
    base_url: str,
    api_key: str,
    account_name: str,
    timeout: int,
    batch_size: int,
) -> None:
    st.subheader("Clean All Attributes")
    st.caption(
        "This fetches matching entries and sends updates with value omitted. Optibus rejects this for mandatory "
        "attributes, so clean mode should target only optional attributes or skip known mandatory attribute IDs."
    )
    st.warning(
        "If Optibus returns 'mandatory attribute cannot have empty value', add that attribute ID to the skip list "
        "or narrow the Attribute IDs filter to optional attributes only."
    )
    try:
        filters = render_filter_controls("clean")
    except Exception as exc:  # noqa: BLE001
        st.error(str(exc))
        return

    skip_attribute_ids = parse_csv_filter(
        st.text_area(
            "Mandatory attribute IDs to skip",
            placeholder="licenseNumber, homeDepot",
            key="clean_skip_attribute_ids",
            help="Comma, semicolon, or newline separated. These fetched attributes are shown in preview but excluded from the clean payload.",
        )
    )
    require_attribute_filter = st.checkbox(
        "Require Attribute IDs filter before cleaning",
        value=True,
        key="clean_require_attribute_filter",
        help="Recommended. Cleaning all attributes usually fails when mandatory attributes are present.",
    )
    dry_run = st.checkbox("Dry run only", value=True, key="clean_dry_run")
    allow_unfiltered = st.checkbox("Allow unfiltered clean", key="clean_allow_unfiltered")
    confirm_clean = st.checkbox(
        "I understand this will clean matching optional attributes",
        key="clean_confirmation_checkbox",
        help="Required before Run Clean is enabled. Keep Dry run only checked until the preview looks correct.",
    )
    has_filter = bool(filters["driver_ids"] or filters["attribute_ids"] or filters["from_date"] or filters["to_date"])
    has_attribute_filter = bool(filters["attribute_ids"])
    if require_attribute_filter and not has_attribute_filter:
        st.info("Enter one or more optional Attribute IDs to clean, or disable the attribute-filter requirement.")
    ready = (
        api_credentials_ready(base_url, api_key, account_name)
        and (has_filter or allow_unfiltered)
        and (has_attribute_filter or not require_attribute_filter)
    )

    if st.button("Preview Clean Payload", disabled=not ready, key="preview_clean"):
        render_clean_preview(
            base_url=base_url,
            api_key=api_key,
            account_name=account_name,
            timeout=timeout,
            batch_size=batch_size,
            filters=filters,
            skip_attribute_ids=skip_attribute_ids,
            execute=False,
        )

    execute_ready = ready and not dry_run and confirm_clean
    if st.button("Run Clean", type="primary", disabled=not execute_ready, key="run_clean"):
        render_clean_preview(
            base_url=base_url,
            api_key=api_key,
            account_name=account_name,
            timeout=timeout,
            batch_size=batch_size,
            filters=filters,
            skip_attribute_ids=skip_attribute_ids,
            execute=True,
        )


def render_clean_preview(
    *,
    base_url: str,
    api_key: str,
    account_name: str,
    timeout: int,
    batch_size: int,
    filters: Dict[str, Any],
    skip_attribute_ids: Sequence[str],
    execute: bool,
) -> None:
    try:
        fetched = fetch_custom_attributes(
            base_url=base_url,
            api_key=api_key,
            account_name=account_name,
            driver_ids=filters["driver_ids"],
            attribute_ids=filters["attribute_ids"],
            from_date=filters["from_date"],
            to_date=filters["to_date"],
            timeout=timeout,
        )
        entries, skipped_entries = clean_entries_from_fetched(fetched, skip_attribute_ids=skip_attribute_ids)
    except Exception as exc:  # noqa: BLE001
        st.error(str(exc))
        return

    metric_cols = st.columns(3)
    metric_cols[0].metric("Matched entries", f"{len(fetched):,}")
    metric_cols[1].metric("Clean payload updates", f"{len(entries):,}")
    metric_cols[2].metric("Skipped entries", f"{len(skipped_entries):,}")

    st.write("Matched attributes")
    st.dataframe(summarize_entries_by_attribute(fetched), use_container_width=True, hide_index=True)
    st.write("Fetched entries")
    st.dataframe(fetched_entries_to_template(fetched), use_container_width=True, hide_index=True)
    if skipped_entries:
        st.info("Skipped entries were excluded because their attribute IDs are in the mandatory skip list.")
        st.dataframe(fetched_entries_to_template(skipped_entries), use_container_width=True, hide_index=True)
    st.write("Clean payload index map")
    st.dataframe(clean_payload_dataframe(entries), use_container_width=True, hide_index=True)
    with st.expander("Clean payload preview", expanded=True):
        st.json(entries_to_payload(entries))

    if execute:
        if not entries:
            st.error("No entries remain in the clean payload after applying filters and skip list.")
            return
        try:
            responses = put_custom_attribute_batches(
                base_url=base_url,
                api_key=api_key,
                account_name=account_name,
                entries=entries,
                batch_size=batch_size,
                timeout=timeout,
            )
        except Exception as exc:  # noqa: BLE001
            st.error(str(exc))
            return
        st.success(f"Cleaned {len(entries):,} entries across {len(responses):,} API batch(es).")
        st.json(responses)


def render_driver_custom_attributes_tab() -> None:
    st.subheader("Driver Custom Attributes")
    st.caption(
        "Upload templates for /v2/drivers/custom-attributes, preview the JSON payload, and run controlled PUT imports."
    )
    st.warning(
        "Before editing existing attributes, fetch the current entries and build your upload from that export. "
        "This avoids accidentally changing the affected driver + attribute timeline."
    )

    render_template_download()
    st.divider()
    base_url, api_key, account_name, timeout, batch_size = render_connection_controls()
    render_fetch_export(base_url, api_key, account_name, timeout)
    st.divider()

    mode = st.radio(
        "Mode",
        [
            "Append new entries",
            "Edit existing timeline",
            "Clean all attributes",
        ],
        horizontal=True,
    )

    if mode == "Clean all attributes":
        render_clean_mode(
            base_url=base_url,
            api_key=api_key,
            account_name=account_name,
            timeout=timeout,
            batch_size=batch_size,
        )
    else:
        render_import_mode(
            mode=mode,
            base_url=base_url,
            api_key=api_key,
            account_name=account_name,
            timeout=timeout,
            batch_size=batch_size,
        )
