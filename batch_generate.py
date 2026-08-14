"""
batch_generate.py
------------------
Drop this file next to app.py in the backflow_app repo.

Adds a "Batch Generate" mode: the inspector fills one spreadsheet
(matching your Backflow-template.xlsx headers, one tab per template)
and the app produces ALL the PDFs at once, zipped for download.

Field keys below were pulled directly from app.py's UNITED_TEXT_FIELDS,
JAX_TEXT_FIELDS, UNITED_CHECKBOXES and JAX_CHECKBOXES dicts, and from
generate_united_pdf() / generate_jax_pdf() -- so this maps 1:1 onto your
existing PDF-drawing logic. No PDF drawing code is duplicated here.

INTEGRATION (in app.py):
    from batch_generate import render_batch_tab

    tab_united, tab_jax, tab_jobs, tab_batch = st.tabs(
        ["United Fire", "Jacksonville", "Jobs", "Batch Generate"]
    )
    ...
    with tab_batch:
        render_batch_tab(generate_united_pdf, generate_jax_pdf, add_job_to_session)
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
    "signature": "final_tester_name",
    "date": "signature_date",
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
        if mapped_key in BOOL_KEYS:
            val = str(val).strip().lower() in ("yes", "true", "1", "x", "y")
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


def build_zip(df, fmt: str, generate_united_pdf, generate_jax_pdf):
    generator = generate_jax_pdf if fmt == "jax" else generate_united_pdf
    columns = list(df.columns)

    name_col = "premise name" if fmt == "jax" else "customer name"
    serial_col = "serial number"

    buf = io.BytesIO()
    errors = []
    count = 0
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (_, row) in enumerate(df.iterrows(), start=1):
            form_data = row_to_form_data(row, columns, fmt)
            if not form_data:
                continue
            try:
                pdf_bytes = generator(form_data)
            except Exception as e:
                errors.append(f"Row {i}: {e}")
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

            filename = f"{label}_{serial}.pdf".replace(" ", "_")
            zf.writestr(filename, pdf_bytes)
            count += 1

    buf.seek(0)
    return buf.getvalue(), count, errors


def render_batch_tab(generate_united_pdf, generate_jax_pdf, add_job_to_session=None):
    """Streamlit UI: call this inside a tab in app.py's main()."""
    st.subheader("Batch Generate Reports")
    st.caption(
        "Upload a spreadsheet matching your Backflow-template.xlsx layout. "
        "One tab for Jacksonville-format devices, one for United-format devices. "
        "Every row becomes one PDF."
    )

    uploaded = st.file_uploader("Spreadsheet (.xlsx)", type=["xlsx"], key="batch_upload")
    if not uploaded:
        return

    sheets = pd.read_excel(uploaded, sheet_name=None, header=0)
    sheet_name = st.selectbox("Sheet to process", list(sheets.keys()), key="batch_sheet_select")
    df = sheets[sheet_name].dropna(how="all")

    fmt = detect_format(sheet_name, list(df.columns))
    fmt_override = st.radio(
        "Template to use", ["Auto-detected: " + fmt.upper(), "Jacksonville", "United"],
        index=0, key="batch_fmt_override",
    )
    if fmt_override == "Jacksonville":
        fmt = "jax"
    elif fmt_override == "United":
        fmt = "united"

    st.write(f"Detected **{len(df)}** report rows. Using **{fmt.upper()}** template.")
    st.dataframe(df.head(10).astype(str))


    save_to_jobs = False
    if add_job_to_session:
        save_to_jobs = st.checkbox("Also add each report to today's Jobs tab", value=True)

    if st.button("Generate all reports as ZIP", type="primary"):
        with st.spinner(f"Generating {len(df)} PDFs..."):
            zip_bytes, count, errors = build_zip(df, fmt, generate_united_pdf, generate_jax_pdf)

            if save_to_jobs and add_job_to_session:
                columns = list(df.columns)
                for _, row in df.iterrows():
                    fd = row_to_form_data(row, columns, fmt)
                    if not fd:
                        continue
                    try:
                        pdf_bytes = (generate_jax_pdf if fmt == "jax" else generate_united_pdf)(fd)
                        add_job_to_session(fd, pdf_bytes, fmt)
                    except Exception:
                        pass

        st.success(f"Generated {count} of {len(df)} reports.")
        if errors:
            st.warning("Some rows failed:\n" + "\n".join(errors))

        st.download_button(
            "Download ZIP of all reports",
            data=zip_bytes,
            file_name=f"backflow_batch_{fmt}_{datetime.now():%Y%m%d_%H%M}.zip",
            mime="application/zip",
        )
