"""
batch_generate.py
------------------
Drop this file next to app.py in the backflow_app repo.

Adds a "Batch Generate" mode: the inspector fills one spreadsheet
(matching your Backflow-template.xlsx headers, one tab per template)
and the app produces ALL the PDFs at once, zipped for download.

Every populated sheet in the uploaded workbook is processed automatically
in a single pass -- no sheet picker, no manual template confirmation.
Multiple devices at the same address are handled per-row, so a single
upload can mix device types/serials under one customer without issue.

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
# JACKSONVILLE (JAX) column -> internal key map
# Matches JAX_TEXT_FIELDS / JAX_CHECKBOXES in app.py exactly.
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

# Ordered pair for the duplicated "Test purpose / Service type / Reclaim /
# Fire service bypass" column blocks -- first occurrence = Commercial,
# second occurrence = Residential.
JAX_DUPLICATE_BLOCKS = [
    ("test purpose", ["comm_test_purpose", "res_test_purpose"]),
    ("service type", ["comm_service_type", "res_service_type"]),
    ("reclaim", ["comm_reclaim", "res_reclaim"]),
    ("fire service bypass", ["comm_fire_bypass", None]),  # no res equivalent field
]

# ---------------------------------------------------------------------------
# UNITED (backflow_template.pdf) column -> internal key map
# Matches UNITED_TEXT_FIELDS / UNITED_CHECKBOXES in app.py exactly.
# ---------------------------------------------------------------------------
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


def _normalize_result(val) -> str:
    key = str(val).strip().lower()
    return RESULT_NORMALIZE.get(key, str(val).strip())


def _normalize(col: str) -> str:
    return str(col).strip().lower()


def _resolve_duplicate_columns(columns: list) -> dict:
    """Handle the JAX sheet's repeated headers (Test purpose, Service type,
    Reclaim, Fire service bypass appear twice: commercial then residential).
    Returns {column_position_index: internal_key}."""
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


def _process_sheet(df, fmt: str, generate_united_pdf, generate_jax_pdf, tester_profile=None):
    """Turn every row of one sheet into a generated PDF. Returns
    (results, errors) where results is a list of
    (filename, pdf_bytes, form_data, fmt) tuples."""
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
    """Process every non-empty sheet in the uploaded workbook and bundle
    every generated PDF into a single ZIP. Returns (zip_bytes, results, errors)."""
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


def render_batch_tab(generate_united_pdf, generate_jax_pdf, add_job_to_session=None, tester_profile=None):
    """Streamlit UI: call this inside a tab in app.py's main().

    Every populated sheet in the uploaded workbook is processed
    automatically. One click generates every report across every sheet
    and produces a single ZIP for download.
    """
    st.subheader("Batch Generate Reports")
    st.caption(
        "Upload your spreadsheet. Every populated sheet (Jacksonville, United, etc.) "
        "is detected and processed automatically -- no need to pick a sheet or "
        "confirm a template. Multiple devices at the same address are handled fine."
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

    if st.button(f"\u2705 Generate All {total_rows} Reports", type="primary", use_container_width=True):
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
        )
