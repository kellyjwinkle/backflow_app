"""
batch_generate.py
------------------
Drop this file next to app.py in the backflow_app repo.

Two ways to bulk-generate reports:

1. "Build In-App" (no spreadsheet) -- set shared defaults once (address,
   size, device type, branch, etc.), specify how many rows you need, then
   edit only the fields that differ per building/device directly in an
   in-app table. This is the primary/default mode.

2. "Upload Spreadsheet" -- the original spreadsheet-import flow, kept for
   compatibility with existing Backflow-template.xlsx files.

Field keys below were pulled directly from app.py's UNITED_TEXT_FIELDS,
JAX_TEXT_FIELDS, UNITED_CHECKBOXES and JAX_CHECKBOXES dicts, and from
generate_united_pdf() / generate_jax_pdf() -- so this maps 1:1 onto your
existing PDF-drawing logic. No PDF drawing code is duplicated here.

INTEGRATION (in app.py) -- unchanged from before:
    from batch_generate import render_batch_tab

    tab_united, tab_jax, tab_jobs, tab_batch = st.tabs(
        ["United Fire", "Jacksonville", "Jobs", "Batch Generate"]
    )
    ...
    with tab_batch:
        _active_tech = st.session_state.get("_sidebar_tech_sel", "")
        _tech_profile = get_technician_profile(_active_tech) if _active_tech else {}
        render_batch_tab(generate_united_pdf, generate_jax_pdf, add_job_to_session, _tech_profile)
"""

import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# JACKSONVILLE (JAX) column -> internal key map (spreadsheet-upload mode)
# ---------------------------------------------------------------------------
JAX_COLUMN_MAP = {
    "premise name": "premises_name",
    "owner name": "owner_name",
    "service address": "service_address",
    "mailing address": "mailing_address",
    "jea account #": "jea_account",
    "contact phone": "contact_phone",
    "meter number": "meter_number",
    "device type": "device_type",
    "serial number": "serial_number",
    "install date": "install_date",
    "physical location": "physical_location",
    "manufacturer": "manufacturer",
    "model number": "model_number",
    "size": "size",
    "cv1": "init_cv1_result",
    "cv1 psi": "init_cv1_psi",
    "cv2": "init_cv2_result",
    "cv2 psi": "init_cv2_psi",
    "init rv": "init_rv_result",
    "init psi": "init_rv_psi",
    "init pvb": "init_pvb_result",
    "init pvb psi": "init_pvb_psi",
    "assembly results": "assembly_result",
    "signature date": "signature_date",
}

JAX_DUPLICATE_BLOCKS = [
    ("test purpose", ["comm_test_purpose", "res_test_purpose"]),
    ("service type", ["comm_service_type", "res_service_type"]),
    ("reclaim", ["comm_reclaim", "res_reclaim"]),
    ("fire service bypass", ["comm_fire_bypass", None]),
]

UNITED_COLUMN_MAP = {
    "customer name": "customer_name",
    "street address": "street_address",
    "branch": "branch",
    "ahj": "ahj",
    "manufacturer": "manufacturer",
    "model": "model",
    "size": "size",
    "test date": "date",
    "serial number": "serial_number",
    "location/building": "location",
    "assembly type": "assembly_type",
    "system service": "system_service",
    "bypass": "bypass",
    "cv1 result": "cv1_result",
    "cv1 differential pressure": "cv1_dp",
    "cv2 result": "cv2_result",
    "cv2 differential pressure": "cv2_dp",
    "rv result": "rv_result",
    "rv opened at psi": "rv_psi",
    "rv outlet": "rv_out_result",
    "rv inlet": "rv_in_result",
    "air inlet result": "pvb_ai_result",
    "air inlet psi": "pvb_ai_psi",
    "cv result": "pvb_cv_result",
    "cv psi": "pvb_cv_psi",
    "assembly result": "assembly_result",
}

BOOL_KEYS = {"comm_fire_bypass"}

RESULT_NORMALIZE = {
    "closed tight": "Closed Tight",
    "leaked": "Leaked",
    "opened": "Opened",
    "did not open": "Did Not Open",
    "n/a": "N/A",
    "air inlet opened": "Air Inlet Opened",
    "air inlet did not": "Air Inlet Did Not",
    "satisfactory": "Satisfactory",
    "held": "Held",
}

# ---------------------------------------------------------------------------
# Field definitions for "Build In-App" mode: (internal_key, label, kind, options)
# kind is "text" or "select" (select is just a hint for the defaults form --
# the in-app table itself uses plain editable text cells for simplicity).
# ---------------------------------------------------------------------------
UNITED_FIELD_DEFS = [
    ("customer_name", "Customer Name", "text", None),
    ("street_address", "Street Address", "text", None),
    ("branch", "Branch", "text", None),
    ("ahj", "AHJ", "text", None),
    ("manufacturer", "Manufacturer", "text", None),
    ("model", "Model", "text", None),
    ("size", "Size", "text", None),
    ("date", "Test Date", "text", None),
    ("location", "Location/Building", "text", None),
    ("assembly_type", "Assembly Type", "select", ["", "RP", "DC", "PVB", "SVB"]),
    ("system_service", "System Service", "select", ["", "Fire", "Domestic", "Irrigation", "Attraction"]),
    ("bypass", "Bypass", "select", ["", "Yes", "No"]),
    ("serial_number", "Serial Number", "text", None),
    ("cv1_result", "CV1 Result", "select", ["", "Closed Tight", "Leaked"]),
    ("cv1_dp", "CV1 Differential Pressure", "text", None),
    ("cv2_result", "CV2 Result", "select", ["", "Closed Tight", "Leaked"]),
    ("cv2_dp", "CV2 Differential Pressure", "text", None),
    ("rv_result", "RV Result", "select", ["", "Opened", "Did Not Open"]),
    ("rv_psi", "RV Opened At PSI", "text", None),
    ("rv_out_result", "RV Outlet", "select", ["", "Closed Tight", "Leaked"]),
    ("rv_in_result", "RV Inlet", "select", ["", "Closed Tight", "Leaked"]),
    ("pvb_ai_result", "Air Inlet Result", "select", ["", "Opened", "Did Not Open"]),
    ("pvb_ai_psi", "Air Inlet PSI", "text", None),
    ("pvb_cv_result", "CV Result (PVB)", "select", ["", "Held", "Leaked"]),
    ("pvb_cv_psi", "CV PSI (PVB)", "text", None),
    ("assembly_result", "Assembly Result", "select", ["", "PASSED", "FAILED"]),
    ("repair_desc", "Repair Description", "text", None),
]

JAX_FIELD_DEFS = [
    ("premises_name", "Premises Name", "text", None),
    ("owner_name", "Owner Name", "text", None),
    ("service_address", "Service Address", "text", None),
    ("mailing_address", "Mailing Address", "text", None),
    ("jea_account", "JEA Account #", "text", None),
    ("contact_phone", "Contact Phone", "text", None),
    ("meter_number", "Meter Number", "text", None),
    ("comm_test_purpose", "Comm Test Purpose", "select", ["", "Annual", "Repair", "Replacement", "New Install"]),
    ("comm_service_type", "Comm Service Type", "select", ["", "Fire", "Irrigation", "Process", "Potable"]),
    ("comm_reclaim", "Comm Reclaim", "select", ["", "Yes", "No"]),
    ("device_type", "Device Type", "select", ["", "RP", "DC", "PVB", "SVB"]),
    ("manufacturer", "Manufacturer", "text", None),
    ("model_number", "Model Number", "text", None),
    ("size", "Size", "text", None),
    ("install_date", "Install Date", "text", None),
    ("physical_location", "Physical Location", "text", None),
    ("serial_number", "Serial Number", "text", None),
    ("init_cv1_result", "Init CV1", "select", ["", "Closed Tight", "Leaked"]),
    ("init_cv1_psi", "Init CV1 PSI", "text", None),
    ("init_cv2_result", "Init CV2", "select", ["", "Closed Tight", "Leaked"]),
    ("init_cv2_psi", "Init CV2 PSI", "text", None),
    ("init_rv_result", "Init RV", "select", ["", "Opened", "Did Not Open"]),
    ("init_rv_psi", "Init RV PSI", "text", None),
    ("init_pvb_result", "Init PVB", "select", ["", "Air Inlet Opened", "Air Inlet Did Not"]),
    ("init_pvb_psi", "Init PVB PSI", "text", None),
    ("assembly_result", "Assembly Result", "select", ["", "PASSED", "FAILED"]),
    ("signature_date", "Signature Date", "text", None),
]


def _normalize_result(val) -> str:
    key = str(val).strip().lower()
    return RESULT_NORMALIZE.get(key, str(val).strip())


def _normalize(col: str) -> str:
    return str(col).strip().lower()


def _resolve_duplicate_columns(columns: list) -> dict:
    seen = {}
    resolved = {}
    for idx, col in enumerate(columns):
        key = _normalize(col)
        for base, targets in JAX_DUPLICATE_BLOCKS:
            if key == base:
                n = seen.get(base, 0)
                if n < len(targets) and targets[n]:
                    resolved[idx] = targets[n]
                seen[base] = n + 1
                break
    return resolved


def row_to_form_data(row, columns: list, fmt: str) -> dict:
    col_map = JAX_COLUMN_MAP if fmt == "jax" else UNITED_COLUMN_MAP
    dup_resolved = _resolve_duplicate_columns(columns) if fmt == "jax" else {}

    form_data = {}
    for idx, (col, val) in enumerate(row.items()):
        if pd.isna(val):
            continue
        key = _normalize(col)
        if idx in dup_resolved:
            mapped_key = dup_resolved[idx]
        else:
            mapped_key = col_map.get(key)
        if not mapped_key:
            continue
        if isinstance(val, (pd.Timestamp, datetime)):
            val = val.strftime("%m/%d/%Y")
        elif mapped_key in BOOL_KEYS:
            val = str(val).strip().lower() in ("yes", "true", "1", "x", "y")
        elif mapped_key.endswith("_result") or mapped_key.endswith("_reclaim"):
            val = _normalize_result(val)
        else:
            val = str(val).strip()
        form_data[mapped_key] = val

    if fmt == "jax" and "signature_date" in form_data:
        form_data.setdefault("init_test_date", form_data["signature_date"])
        form_data.setdefault("final_test_date", form_data["signature_date"])
    if fmt == "united" and "date" in form_data:
        form_data.setdefault("test_date", form_data["date"])

    return form_data


def detect_format(sheet_name: str, headers: list) -> str:
    name = sheet_name.strip().lower()
    if "jack" in name or "jax" in name:
        return "jax"
    if "united" in name:
        return "united"
    normalized = {_normalize(h) for h in headers}
    if "jea account #" in normalized or "premise name" in normalized:
        return "jax"
    return "united"


def _apply_tester_profile(form_data: dict, fmt: str, tester_profile: dict):
    if not tester_profile:
        return
    if fmt == "jax":
        form_data.setdefault("init_tester_name", tester_profile.get("technician", ""))
        form_data.setdefault("init_company", tester_profile.get("company", ""))
        form_data.setdefault("init_cert", tester_profile.get("cert_no", ""))
        form_data.setdefault("final_tester_name", tester_profile.get("technician", ""))
        form_data.setdefault("final_company", tester_profile.get("company", ""))
        form_data.setdefault("final_cert", tester_profile.get("cert_no", ""))
    else:
        form_data.setdefault("technician", tester_profile.get("technician", ""))
        form_data.setdefault("gauge_mfg", tester_profile.get("gauge_mfg", ""))
        form_data.setdefault("gauge_serial", tester_profile.get("gauge_serial", ""))
        form_data.setdefault("date_cal", tester_profile.get("date_cal", ""))
        form_data.setdefault("cert_no", tester_profile.get("cert_no", ""))
        form_data.setdefault("recert", tester_profile.get("recert", ""))
    form_data.setdefault("signature_b64", tester_profile.get("signature_b64", ""))


# ---------------------------------------------------------------------------
# Spreadsheet-upload mode (original flow, kept for compatibility)
# ---------------------------------------------------------------------------

def _process_sheet(df, fmt: str, generate_united_pdf, generate_jax_pdf, tester_profile=None):
    generator = generate_jax_pdf if fmt == "jax" else generate_united_pdf
    columns = list(df.columns)
    name_col = "premise name" if fmt == "jax" else "customer name"
    serial_col = "serial number"

    results = []
    errors = []
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        form_data = row_to_form_data(row, columns, fmt)
        if not form_data:
            continue
        _apply_tester_profile(form_data, fmt, tester_profile)
        try:
            pdf_bytes = generator(form_data)
        except Exception as e:
            errors.append(f"{fmt.upper()} row {i}: {e}")
            continue

        label = "unknown"
        for col, val in row.items():
            if _normalize(col) == name_col and pd.notna(val):
                label = str(val).strip().replace("/", "-")[:60]
                break
        serial = "NA"
        for col, val in row.items():
            if _normalize(col) == serial_col and pd.notna(val):
                serial = str(val).strip()
                break

        filename = f"{fmt}_{label}_{serial}.pdf".replace(" ", "_")
        results.append((filename, pdf_bytes, form_data, fmt))

    return results, errors


def build_zip_all(sheets: dict, generate_united_pdf, generate_jax_pdf, tester_profile=None):
    buf = io.BytesIO()
    all_results = []
    all_errors = []
    seen_names = {}

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for sheet_name, raw_df in sheets.items():
            df = raw_df.dropna(how="all")
            if df.empty:
                continue
            fmt = detect_format(sheet_name, list(df.columns))
            results, errors = _process_sheet(df, fmt, generate_united_pdf, generate_jax_pdf, tester_profile)
            all_errors.extend(errors)
            for filename, pdf_bytes, form_data, r_fmt in results:
                n = seen_names.get(filename, 0)
                final_name = filename if n == 0 else filename.replace(".pdf", f"_{n}.pdf")
                seen_names[filename] = n + 1
                zf.writestr(final_name, pdf_bytes)
                all_results.append((final_name, pdf_bytes, form_data, r_fmt))

    buf.seek(0)
    return buf.getvalue(), all_results, all_errors


def _render_spreadsheet_upload(generate_united_pdf, generate_jax_pdf, add_job_to_session=None, tester_profile=None):
    st.caption(
        "Upload your spreadsheet. Every populated sheet (Jacksonville, United, etc.) "
        "is detected and processed automatically -- no need to pick a sheet or "
        "confirm a template."
    )

    uploaded = st.file_uploader("Spreadsheet (.xlsx)", type=["xlsx"], key="batch_upload")
    if not uploaded:
        return

    sheets = pd.read_excel(uploaded, sheet_name=None, header=0)

    total_rows = 0
    sheet_summaries = []
    for sheet_name, raw_df in sheets.items():
        df = raw_df.dropna(how="all")
        if df.empty:
            continue
        fmt = detect_format(sheet_name, list(df.columns))
        total_rows += len(df)
        sheet_summaries.append((sheet_name, fmt, df))

    if total_rows == 0:
        st.warning("No data rows found in any sheet of this file.")
        return

    st.write(f"**{total_rows}** report row(s) detected across **{len(sheet_summaries)}** sheet(s).")
    for sheet_name, fmt, df in sheet_summaries:
        with st.expander(f"\U0001f4c4 {sheet_name} \u2014 {len(df)} rows \u2192 {fmt.upper()} template"):
            st.dataframe(df.head(5).astype(str))

    if st.button(f"\u2705 Generate All {total_rows} Reports", type="primary", use_container_width=True, key="ss_generate"):
        with st.spinner(f"Generating {total_rows} PDFs..."):
            zip_bytes, results, errors = build_zip_all(sheets, generate_united_pdf, generate_jax_pdf, tester_profile)
            if add_job_to_session:
                for _, pdf_bytes, form_data, fmt in results:
                    try:
                        add_job_to_session(form_data, pdf_bytes, fmt)
                    except Exception:
                        pass

        st.success(f"Generated {len(results)} of {total_rows} reports.")
        if errors:
            st.warning("Some rows failed:\n" + "\n".join(errors))

        st.download_button(
            f"\u2b07\ufe0f Download ZIP ({len(results)} PDFs)",
            data=zip_bytes,
            file_name=f"backflow_batch_{datetime.now():%Y%m%d_%H%M}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
            key="ss_download",
        )


# ---------------------------------------------------------------------------
# Build In-App mode (no spreadsheet)
# ---------------------------------------------------------------------------

def _row_dict_to_form_data(row: dict, key_by_label: dict, fmt: str) -> dict:
    form_data = {}
    for label, val in row.items():
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        val = str(val).strip()
        if not val:
            continue
        key = key_by_label.get(label)
        if not key:
            continue
        if key.endswith("_result") or key.endswith("_reclaim"):
            val = _normalize_result(val)
        form_data[key] = val

    if fmt == "jax" and "signature_date" in form_data:
        form_data.setdefault("init_test_date", form_data["signature_date"])
        form_data.setdefault("final_test_date", form_data["signature_date"])
    if fmt == "united" and "date" in form_data:
        form_data.setdefault("test_date", form_data["date"])

    return form_data


def _build_zip_from_table(df, key_by_label: dict, fmt: str, generate_united_pdf, generate_jax_pdf, tester_profile=None):
    generator = generate_jax_pdf if fmt == "jax" else generate_united_pdf
    name_label = "Premises Name" if fmt == "jax" else "Customer Name"
    serial_label = "Serial Number"

    buf = io.BytesIO()
    results = []
    errors = []
    seen_names = {}
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, row in enumerate(df.to_dict(orient="records"), start=1):
            form_data = _row_dict_to_form_data(row, key_by_label, fmt)
            if not form_data:
                continue
            _apply_tester_profile(form_data, fmt, tester_profile)
            try:
                pdf_bytes = generator(form_data)
            except Exception as e:
                errors.append(f"Row {i}: {e}")
                continue

            label_val = str(row.get(name_label) or "").strip().replace("/", "-")[:60] or "unknown"
            serial_val = str(row.get(serial_label) or "").strip() or "NA"
            filename = f"{fmt}_{label_val}_{serial_val}.pdf".replace(" ", "_")
            n = seen_names.get(filename, 0)
            final_name = filename if n == 0 else filename.replace(".pdf", f"_{n}.pdf")
            seen_names[filename] = n + 1
            zf.writestr(final_name, pdf_bytes)
            results.append((final_name, pdf_bytes, form_data, fmt))

    buf.seek(0)
    return buf.getvalue(), results, errors


def _render_group_builder(generate_united_pdf, generate_jax_pdf, add_job_to_session=None, tester_profile=None):
    st.caption(
        "Set the shared defaults once below, then add rows for each building/device. "
        "Every new row starts pre-filled with your defaults -- just edit the cells "
        "that are different (address, serial number, PSI, etc.) directly in the table."
    )

    fmt_choice = st.radio("Template", ["United", "Jacksonville"], horizontal=True, key="grp_fmt_choice")
    fmt_key = "united" if fmt_choice == "United" else "jax"
    field_defs = UNITED_FIELD_DEFS if fmt_key == "united" else JAX_FIELD_DEFS
    key_by_label = {label: key for key, label, _, _ in field_defs}

    defaults_key = f"grp_defaults_{fmt_key}"
    if defaults_key not in st.session_state:
        st.session_state[defaults_key] = {key: "" for key, _, _, _ in field_defs}
    defaults = st.session_state[defaults_key]

    with st.expander("\u2699\ufe0f Shared Defaults \u2014 set once, applies to every new row", expanded=True):
        cols = st.columns(3)
        for i, (key, label, kind, opts) in enumerate(field_defs):
            with cols[i % 3]:
                widget_key = f"{defaults_key}_{key}"
                if kind == "select":
                    current = defaults.get(key, "")
                    idx = opts.index(current) if current in opts else 0
                    val = st.selectbox(label, opts, index=idx, key=widget_key)
                else:
                    val = st.text_input(label, value=defaults.get(key, ""), key=widget_key)
                defaults[key] = val

    table_key = f"grp_table_{fmt_key}"
    if table_key not in st.session_state:
        st.session_state[table_key] = pd.DataFrame(
            [{label: defaults.get(key, "") for key, label, _, _ in field_defs}]
        )

    c1, c2 = st.columns([3, 1])
    with c1:
        n_new = st.number_input(
            "Add how many rows using current defaults?", min_value=1, max_value=200, value=1,
            key=f"{table_key}_n",
        )
    with c2:
        st.write("")
        if st.button("\u2795 Add Rows", key=f"{table_key}_add", use_container_width=True):
            new_rows = pd.DataFrame(
                [{label: defaults.get(key, "") for key, label, _, _ in field_defs} for _ in range(int(n_new))]
            )
            st.session_state[table_key] = pd.concat([st.session_state[table_key], new_rows], ignore_index=True)
            st.rerun()

    edited_df = st.data_editor(
        st.session_state[table_key],
        num_rows="dynamic",
        use_container_width=True,
        key=f"{table_key}_editor",
    )
    st.session_state[table_key] = edited_df

    if st.button(
        f"\u2705 Generate All {len(edited_df)} Reports", type="primary",
        use_container_width=True, key=f"{table_key}_gen",
    ):
        with st.spinner(f"Generating {len(edited_df)} PDFs..."):
            zip_bytes, results, errors = _build_zip_from_table(
                edited_df, key_by_label, fmt_key, generate_united_pdf, generate_jax_pdf, tester_profile
            )
            if add_job_to_session:
                for _, pdf_bytes, form_data, r_fmt in results:
                    try:
                        add_job_to_session(form_data, pdf_bytes, r_fmt)
                    except Exception:
                        pass

        st.success(f"Generated {len(results)} of {len(edited_df)} reports.")
        if errors:
            st.warning("Some rows failed:\n" + "\n".join(errors))

        st.download_button(
            f"\u2b07\ufe0f Download ZIP ({len(results)} PDFs)",
            data=zip_bytes,
            file_name=f"backflow_group_{fmt_key}_{datetime.now():%Y%m%d_%H%M}.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
            key=f"{table_key}_download",
        )


def render_batch_tab(generate_united_pdf, generate_jax_pdf, add_job_to_session=None, tester_profile=None):
    """Streamlit UI: call this inside a tab in app.py's main().

    Two modes: build the batch directly in the app (default, no spreadsheet
    needed), or upload a spreadsheet (kept for backward compatibility).
    """
    st.subheader("Batch Generate Reports")
    mode = st.radio(
        "How do you want to provide the data?",
        ["\U0001f3e2 Build In-App (no spreadsheet)", "\U0001f4c4 Upload Spreadsheet"],
        horizontal=True,
        key="batch_mode",
    )

    if mode.startswith("\U0001f3e2"):
        _render_group_builder(generate_united_pdf, generate_jax_pdf, add_job_to_session, tester_profile)
    else:
        _render_spreadsheet_upload(generate_united_pdf, generate_jax_pdf, add_job_to_session, tester_profile)
