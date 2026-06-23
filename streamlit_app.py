from __future__ import annotations

import json
import importlib.util
import zipfile
from io import BytesIO
from typing import Any, Dict, Iterable, List, Sequence, Tuple

import pandas as pd
import streamlit as st

from driver_custom_attributes_app import render_driver_custom_attributes_tab
from otp_report_automation import (
    dedupe_planning,
    extract_planning_records,
    find_payload,
    merge_planning_payloads,
    parse_force_list,
    summarize_planning_df,
)
from optibus_schedule_fixer import fix_schedule


SCHEDULE_COLUMNS = [
    "id",
    "start_otp",
    "end_otp",
    "predicted_otp",
    "Service",
    "Route",
    "Route Code",
    "Direction",
    "Pattern",
    "Start Time",
]


def parse_json_bytes(raw: bytes, name: str) -> Any:
    text = raw.decode("utf-8-sig")
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        try:
            return json.loads("{" + text + "}")
        except json.JSONDecodeError as exc:
            raise ValueError(f"{name} is not valid JSON: {exc}") from exc


def iter_uploaded_json(uploaded_files: Sequence[Any]) -> Iterable[Tuple[str, Any]]:
    for uploaded in uploaded_files:
        raw = uploaded.getvalue()
        filename = uploaded.name
        if filename.lower().endswith(".zip"):
            with zipfile.ZipFile(BytesIO(raw), "r") as zf:
                json_names = sorted(name for name in zf.namelist() if name.lower().endswith(".json"))
                if not json_names:
                    raise ValueError(f"{filename} does not contain any JSON files.")
                for json_name in json_names:
                    yield f"{filename}!{json_name}", parse_json_bytes(zf.read(json_name), json_name)
        else:
            yield filename, parse_json_bytes(raw, filename)


def format_timepoint_column(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "timepointOtp" in out.columns:
        out["timepointOtp"] = out["timepointOtp"].apply(
            lambda value: "[" + ",".join(f"{item:.2f}" for item in value) + "]" if isinstance(value, list) else ""
        )
    return out


def csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    if importlib.util.find_spec("xlsxwriter"):
        engine = "xlsxwriter"
    elif importlib.util.find_spec("openpyxl"):
        engine = "openpyxl"
    else:
        raise RuntimeError("Install xlsxwriter or openpyxl to enable Excel downloads.")

    output = BytesIO()
    with pd.ExcelWriter(output, engine=engine) as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
            if engine == "xlsxwriter":
                worksheet = writer.sheets[sheet_name[:31]]
                for index, column in enumerate(df.columns):
                    width = min(max(len(str(column)), 12), 40)
                    worksheet.set_column(index, index, width)
    return output.getvalue()


def build_schedule_report(schedule_files: Sequence[Any], properties_file: Any) -> pd.DataFrame:
    props = pd.read_excel(BytesIO(properties_file.getvalue()))
    if "System Id" not in props.columns:
        raise ValueError("'System Id' column not found in the trip properties Excel file.")

    props["System Id"] = props["System Id"].astype(str).str.strip()
    trip_properties = props.set_index("System Id").to_dict(orient="index")
    rows: List[Dict[str, Any]] = []

    for source_name, data in iter_uploaded_json(schedule_files):
        payload = find_payload(data, "schedule")
        if not payload:
            st.warning(f"Skipped non-Scheduling payload: {source_name}")
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
    return pd.DataFrame(rows).reindex(columns=SCHEDULE_COLUMNS)


def build_planning_reports(
    planning_files: Sequence[Any],
    *,
    complete_grid: bool,
    force_dirpos: bool,
    force_dirpos_list: str,
    use_dirpos_labels: bool,
    dedupe_include_direction: bool,
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any], int]:
    payloads = []
    for source_name, data in iter_uploaded_json(planning_files):
        payload = find_payload(data, "planning")
        if not payload:
            st.warning(f"Skipped non-Planning payload: {source_name}")
            continue
        payloads.append(payload)

    if not payloads:
        raise ValueError("No Planning OTP payloads were found.")

    merged = merge_planning_payloads(payloads)
    detail = extract_planning_records(
        merged,
        force_dirpos=force_dirpos or bool(force_dirpos_list.strip()),
        force_pairs=parse_force_list(force_dirpos_list),
        dirpos_labels=use_dirpos_labels,
    )
    detail = dedupe_planning(detail, include_direction=dedupe_include_direction)
    summary = summarize_planning_df(detail, complete_grid=complete_grid)
    return summary, detail, merged, len(payloads)


def render_downloads(base_name: str, summary: pd.DataFrame, detail: pd.DataFrame | None = None) -> None:
    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button(
            "Download CSV",
            csv_bytes(summary),
            file_name=f"{base_name}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with col2:
        sheets = {"Summary": summary}
        if detail is not None:
            sheets["Trip Detail"] = format_timepoint_column(detail)
        try:
            excel_data = excel_bytes(sheets)
        except RuntimeError as exc:
            st.button("Download Excel", disabled=True, use_container_width=True, help=str(exc))
        else:
            st.download_button(
                "Download Excel",
                excel_data,
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with col3:
        if detail is not None:
            st.download_button(
                "Download Detail CSV",
                csv_bytes(format_timepoint_column(detail)),
                file_name=f"{base_name}_detail.csv",
                mime="text/csv",
                use_container_width=True,
            )


def render_schedule_tab() -> None:
    st.subheader("Scheduling OTP Extract")
    st.caption("Upload Scheduling JSON files and the Trip Properties Excel file. The result matches the Excel output from the Scheduling script.")

    schedule_files = st.file_uploader(
        "Scheduling JSON, TXT, or ZIP files",
        type=["json", "txt", "zip"],
        accept_multiple_files=True,
        key="schedule_files",
    )
    properties_file = st.file_uploader(
        "Trip Properties Excel file",
        type=["xlsx", "xls"],
        key="trip_properties",
    )

    if st.button("Build Scheduling Report", type="primary", disabled=not schedule_files or properties_file is None):
        try:
            with st.spinner("Processing Scheduling OTP files..."):
                report = build_schedule_report(schedule_files, properties_file)
            st.success(f"Built report with {len(report):,} trips.")
            st.dataframe(report, use_container_width=True, hide_index=True)
            render_downloads("otp_schedule_report", report)
        except Exception as exc:  # noqa: BLE001
            st.error(str(exc))


def render_planning_tab() -> None:
    st.subheader("Planning OTP Summary")
    st.caption("Upload a Planning OTP JSON or ZIP export. The app dedupes by route, service, and trip ID by default.")

    planning_files = st.file_uploader(
        "Planning JSON or ZIP files",
        type=["json", "zip"],
        accept_multiple_files=True,
        key="planning_files",
    )

    with st.expander("Options", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            complete_grid = st.checkbox("Complete route/direction/service grid")
            force_dirpos = st.checkbox("Force direction from array position", value=True)
            use_dirpos_labels = st.checkbox("Label direction positions as Outbound/Inbound", value=True)
        with col2:
            dedupe_include_direction = st.checkbox("Include direction in dedupe key")
            force_dirpos_list = st.text_input(
                "Force direction only for route/service pairs",
                placeholder="10:Mon-Fri;10A:Sat",
                help="Leave blank when forcing direction for all uploaded data.",
            )

    if st.button("Build Planning Summary", type="primary", disabled=not planning_files):
        try:
            with st.spinner("Processing Planning OTP files..."):
                summary, detail, merged, payload_count = build_planning_reports(
                    planning_files,
                    complete_grid=complete_grid,
                    force_dirpos=force_dirpos,
                    force_dirpos_list=force_dirpos_list,
                    use_dirpos_labels=use_dirpos_labels,
                    dedupe_include_direction=dedupe_include_direction,
                )

            metric_cols = st.columns(3)
            metric_cols[0].metric("Payloads", f"{payload_count:,}")
            metric_cols[1].metric("Trip rows", f"{len(detail):,}")
            metric_cols[2].metric("Summary rows", f"{len(summary):,}")

            st.dataframe(summary, use_container_width=True, hide_index=True)
            render_downloads("otp_planning_summary", summary, detail)

            with st.expander("Trip-level detail"):
                st.dataframe(format_timepoint_column(detail), use_container_width=True, hide_index=True)

            st.download_button(
                "Download merged Planning JSON",
                json.dumps(merged, ensure_ascii=False, indent=2).encode("utf-8"),
                file_name="otp_planning_merged.json",
                mime="application/json",
            )
        except Exception as exc:  # noqa: BLE001
            st.error(str(exc))


def render_schedule_fixer_tab() -> None:
    st.subheader("Schedule Import Fixer")
    st.caption(
        "Fix a full/crew/vehicle schedule workbook for Optibus import while preserving sign-on, sign-off, pull, and deadhead events."
    )

    schedule_file = st.file_uploader(
        "Schedule workbook to fix",
        type=["xlsx", "xls"],
        key="schedule_fixer_schedule",
    )

    st.markdown("Upload either the full Optibus export ZIP, or the two reference files separately.")
    export_zip = st.file_uploader(
        "Full Optibus export ZIP",
        type=["zip"],
        key="schedule_fixer_export_zip",
        help="Must contain data_set.json and an export_full_schedule.xlsx file.",
    )

    with st.expander("Separate reference files", expanded=False):
        data_set_json = st.file_uploader(
            "data_set.json",
            type=["json"],
            key="schedule_fixer_dataset",
        )
        full_schedule_ref = st.file_uploader(
            "Reference full schedule workbook",
            type=["xlsx", "xls"],
            key="schedule_fixer_full_ref",
        )

    has_reference = export_zip is not None or (data_set_json is not None and full_schedule_ref is not None)
    if st.button("Fix Schedule", type="primary", disabled=schedule_file is None or not has_reference):
        try:
            with st.spinner("Fixing schedule workbook..."):
                result = fix_schedule(
                    schedule_file.getvalue(),
                    export_zip_bytes=export_zip.getvalue() if export_zip is not None else None,
                    data_set_json_bytes=data_set_json.getvalue() if data_set_json is not None else None,
                    full_schedule_bytes=full_schedule_ref.getvalue() if full_schedule_ref is not None else None,
                )

            st.success("Schedule fixed and validated.")

            metric_cols = st.columns(4)
            metric_cols[0].metric("Input rows", f"{result.stats['input_rows']:,}")
            metric_cols[1].metric("Output rows", f"{result.stats['output_rows']:,}")
            metric_cols[2].metric("Expanded parents", f"{result.stats['expanded_parent_rows']:,}")
            metric_cols[3].metric("Duplicates removed", f"{result.stats['duplicate_rows_removed']:,}")

            validation_df = pd.DataFrame(
                [{"Check": key.replace("_", " ").title(), "Value": value} for key, value in result.validation.items()]
            )
            stats_df = pd.DataFrame(
                [{"Fix Step": key.replace("_", " ").title(), "Value": value} for key, value in result.stats.items()]
            )
            col1, col2 = st.columns(2)
            with col1:
                st.write("Validation")
                st.dataframe(validation_df, use_container_width=True, hide_index=True)
            with col2:
                st.write("Fix summary")
                st.dataframe(stats_df, use_container_width=True, hide_index=True)

            base_name = schedule_file.name.rsplit(".", 1)[0]
            st.download_button(
                "Download fixed schedule",
                result.workbook_bytes,
                file_name=f"{base_name}_fixed_for_optibus.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:  # noqa: BLE001
            st.error(str(exc))


def main() -> None:
    st.set_page_config(page_title="Optibus Tools", layout="wide")
    st.title("Optibus Tools")
    st.write("Build OTP reports, repair schedule import workbooks, and manage driver custom attributes.")

    schedule_tab, planning_tab, fixer_tab, driver_attrs_tab = st.tabs(
        ["Scheduling OTP Extract", "Planning OTP Summary", "Schedule Import Fixer", "Driver Custom Attributes"]
    )
    with schedule_tab:
        render_schedule_tab()
    with planning_tab:
        render_planning_tab()
    with fixer_tab:
        render_schedule_fixer_tab()
    with driver_attrs_tab:
        render_driver_custom_attributes_tab()


if __name__ == "__main__":
    main()
