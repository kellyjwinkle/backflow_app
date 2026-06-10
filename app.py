import streamlit as st
import streamlit.components.v1 as components
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import json, os, re, base64, tempfile, zipfile, requests
from datetime import date
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream
from PIL import Image
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ---------------------------------------------------------------------------
# Paths & constants
# ---------------------------------------------------------------------------
TEMPLATE_UNITED   = "backflow_template.pdf"
TEMPLATE_JAX      = "jacksonville_template.pdf"
TECHNICIANS_FILE  = "technicians.json"
SIG_FILE          = "signature_b64.txt"   # legacy — no longer used for persistence
PAGE_W, PAGE_H     = 612, 792   # United Fire (US Letter)

GITHUB_REPO       = "kellyjwinkle/backflow_app"
GITHUB_API_BASE   = "https://api.github.com"

def _get_pdf_page_size(path):
    try:
        reader = PdfReader(path)
        page = reader.pages[0]
        mb = page.mediabox
        return float(mb.width), float(mb.height)
    except Exception:
        return 595, 842

if os.path.exists(TEMPLATE_JAX):
    JAX_PAGE_W, JAX_PAGE_H = _get_pdf_page_size(TEMPLATE_JAX)
else:
    JAX_PAGE_W, JAX_PAGE_H = 612, 792

# ---------------------------------------------------------------------------
# Technician profile helpers (GitHub-backed)
# ---------------------------------------------------------------------------

def _github_token():
    try:
        return st.secrets["GITHUB_TOKEN"]
    except Exception:
        return None


def _github_headers():
    token = _github_token()
    h = {"Accept": "application/vnd.github+json"}
    if token:
        h["Authorization"] = f"Bearer {token}"
    return h


def load_technicians_from_github():
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO}/contents/{TECHNICIANS_FILE}"
    try:
        r = requests.get(url, headers=_github_headers(), timeout=8)
        if r.status_code == 200:
            data = r.json()
            content = base64.b64decode(data["content"]).decode("utf-8")
            return json.loads(content), data["sha"]
    except Exception:
        pass
    if os.path.exists(TECHNICIANS_FILE):
        with open(TECHNICIANS_FILE, "r") as fh:
            return json.load(fh), None
    return {}, None


def save_technicians_to_github(techs: dict, current_sha):
    token = _github_token()
    if not token:
        with open(TECHNICIANS_FILE, "w") as fh:
            json.dump(techs, fh, indent=2)
        return False, "No GITHUB_TOKEN secret configured — saved locally only."

    url = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO}/contents/{TECHNICIANS_FILE}"
    content_b64 = base64.b64encode(json.dumps(techs, indent=2).encode()).decode()
    payload = {
        "message": "Update technician profiles via app",
        "content": content_b64,
    }
    if current_sha:
        payload["sha"] = current_sha

    try:
        r = requests.put(url, headers=_github_headers(), json=payload, timeout=10)
        if r.status_code in (200, 201):
            return True, "Profile saved to GitHub ✓"
        else:
            return False, f"GitHub API error {r.status_code}: {r.text[:200]}"
    except Exception as e:
        return False, f"Network error: {e}"


def _init_technicians():
    if "technicians" not in st.session_state:
        techs, sha = load_technicians_from_github()
        st.session_state["technicians"] = techs
        st.session_state["technicians_sha"] = sha


def get_technician_names():
    _init_technicians()
    return [""] + list(st.session_state["technicians"].keys())


def get_technician_profile(name: str) -> dict:
    _init_technicians()
    return dict(st.session_state["technicians"].get(name, {}))


def upsert_technician_profile(name: str, profile: dict):
    _init_technicians()
    st.session_state["technicians"][name] = profile
    ok, msg = save_technicians_to_github(
        st.session_state["technicians"],
        st.session_state.get("technicians_sha")
    )
    if ok:
        _, new_sha = load_technicians_from_github()
        st.session_state["technicians_sha"] = new_sha
    return ok, msg


# ---------------------------------------------------------------------------
# United Fire form config
# ---------------------------------------------------------------------------
UNITED_TEXT_FIELDS = {
    "date":           (135, 583, 8),
    "branch":         (235, 583, 8),
    "ahj":            (437, 583, 8),
    "customer_name":  (200, 567, 8),
    "street_address": (200, 551, 8),
    "location":       (200, 533, 8),
    "serial_number":  (205, 507, 8),
    "manufacturer":   (205, 490, 8),
    "model":          (205, 475, 8),
    "size":           (390, 507, 8),
    "rv_psi":         (300, 398, 8),
    "cv1_dp":         (183, 320, 8),
    "cv2_dp":         (395, 312, 8),
    "pvb_ai_psi":     (495, 378, 8),
    "pvb_cv_psi":     (495, 320, 8),
    "test_date":      (168, 290, 8),
    "gauge_mfg":      (215, 178, 8),
    "gauge_serial":   (313, 178, 8),
    "date_cal":       (455, 178, 8),
    "technician":     (176, 155, 8),
    "cert_no":        (407, 165, 8),
    "recert":         (407, 150, 8),
}

UNITED_SIG_X, UNITED_SIG_Y, UNITED_SIG_W, UNITED_SIG_H = 170, 118, 130, 28

UNITED_CHECKBOXES = {
    "RP": (210,460), "DC": (270,460), "PVB": (210,441), "SVB": (270,441),
    "FIRE": (395,490), "DOMESTIC": (395,475), "IRRIGATION": (395,460), "ATTRACTION": (395,442),
    "BYPASS_YES": (500,470), "BYPASS_NO": (500,450),
    "CV1_CLOSED": (130,390), "CV1_LEAKED": (130,375),
    "CV2_CLOSED": (330,390), "CV2_LEAKED": (330,375),
    "PVB_AI_CLOSED": (426,398), "PVB_AI_OPENED": (426,378),
    "PVB_CV_LEAKED": (426,350), "PVB_CV_HELD": (426,323),
    "RV_OPENED": (225,398), "RV_DIDNOTOPEN": (225,378),
    "RV_OUT_CLOSED": (225,334), "RV_OUT_LEAKED": (272,334),
    "RV_IN_CLOSED": (225,310), "RV_IN_LEAKED": (272,310),
    "PASSED": (360,292), "FAILED": (415,292),
}

UNITED_REPAIR_BOX = (228, 200, 10, 3, 70)

# ---------------------------------------------------------------------------
# Field color groups for United Fire form
# ---------------------------------------------------------------------------
UNITED_YELLOW = {
    "gauge_mfg", "gauge_serial", "date_cal", "technician", "cert_no", "recert",
}
UNITED_BLUE = {
    "date", "branch", "ahj", "customer_name", "street_address",
    "serial_number", "manufacturer", "model", "size",
    "assembly_type", "system_service", "bypass",
}
UNITED_GREEN = {
    "location", "rv_psi", "cv1_dp", "cv2_dp", "pvb_ai_psi", "pvb_cv_psi",
    "test_date", "cv1_result", "cv2_result", "rv_result", "rv_out_result",
    "rv_in_result", "pvb_ai_result", "pvb_cv_result", "assembly_result", "repair_desc",
}

UNITED_NEXT_REPORT_KEEP = UNITED_YELLOW | UNITED_BLUE
UNITED_NEW_JOB_KEEP = UNITED_YELLOW

# Widget keys for each color group — must be popped from session_state when clearing
UNITED_GREEN_WIDGET_KEYS = [
    "u_loc", "u_cv1dp", "u_cv2dp", "u_rvpsi", "u_aipsi", "u_cvpsi",
    "u_tdate", "u_rep", "u_res",
    "cv1_result", "cv2_result", "rv_result", "rv_out_result",
    "rv_in_result", "pvb_ai_result", "pvb_cv_result", "assembly_result",
]
UNITED_TESTER_WIDGET_KEYS = ["u_gmfg", "u_gsn", "u_cal", "u_tech", "u_cert", "u_recert"]
UNITED_BLUE_WIDGET_KEYS = [
    "u_date", "u_branch", "u_ahj", "u_cust", "u_addr",
    "u_sn", "u_mfg", "u_mdl", "u_sz", "u_atype", "u_ss", "u_bp",
]

TESTER_KEYS = ["gauge_mfg", "gauge_serial", "date_cal", "technician", "cert_no", "recert"]

# ---------------------------------------------------------------------------
# Jacksonville (JEA) form config
# ---------------------------------------------------------------------------
JAX_TEXT_FIELDS = {
    "premises_name":           (102, 696, 9),
    "owner_name":              (333, 696, 9),
    "service_address":         (104, 652, 9),
    "mailing_address":         (335, 652, 9),
    "physical_location":       (105, 614, 9),
    "contact_phone":           (333, 612, 9),
    "jea_account":             (103, 569, 9),
    "meter_number":            (335, 571, 9),
    "device_type":             (90, 419, 9),
    "manufacturer":            (151, 421, 9),
    "size":                    (231, 421, 9),
    "model_number":            (283, 420, 9),
    "serial_number":           (358, 420, 9),
    "install_date":            (459, 418, 9),
    "init_cv1_psi":            (151, 338, 9),
    "init_cv2_psi":            (249, 338, 9),
    "init_rv_psi":             (396, 356, 9),
    "init_pvb_psi":            (471, 335, 9),
    "final_cv1_psi":           (165, 282, 9),
    "final_cv2_psi":           (263, 283, 9),
    "final_rv_psi":            (404, 299, 9),
    "repairs":                 (99, 244, 9),
    "init_tester_name":        (94, 182, 9),
    "init_company":            (236, 183, 9),
    "init_cert":               (356, 183, 9),
    "init_test_date":          (464, 183, 9),
    "repaired_by":             (95, 158, 9),
    "repair_company":          (239, 156, 9),
    "repair_cert":             (355, 159, 9),
    "repair_date":             (464, 162, 9),
    "final_tester_name":       (93, 135, 9),
    "final_company":           (239, 135, 9),
    "final_cert":              (354, 136, 9),
    "final_test_date":         (468, 139, 9),
    "signature_date":          (433, 84, 9),
}

JAX_SIG_X, JAX_SIG_Y, JAX_SIG_W, JAX_SIG_H = 161, 68, 160, 22

JAX_CHECKBOXES = {
    "COMM_ANNUAL":           (214, 545),
    "COMM_REPAIR":           (286, 544),
    "COMM_REPLACEMENT":      (358, 545),
    "COMM_NEW_INSTALL":      (463, 545),
    "COMM_FIRE":             (214, 523),
    "COMM_IRRIGATION":       (294, 522),
    "COMM_PROCESS":          (362, 521),
    "COMM_POTABLE":          (472, 523),
    "COMM_FIRE_BYPASS":      (215, 510),
    "RECLAIM_YES":           (421, 511),
    "RECLAIM_NO":            (459, 510),
    "RES_ANNUAL":            (210, 489),
    "RES_REPAIR":            (280, 488),
    "RES_REPLACEMENT":       (358, 489),
    "RES_NEW_INSTALL":       (462, 489),
    "RES_POTABLE":           (202, 466),
    "RES_IRRIGATION":        (255, 465),
    "RES_RECLAIM_YES":       (434, 466),
    "RES_RECLAIM_NO":        (472, 464),
    "INIT_CV1_CLOSED":       (139, 363),
    "INIT_CV2_CLOSED":       (235, 362),
    "INIT_RV_OPENED":        (331, 356),
    "INIT_RV_DIDNOT":        (336, 329),
    "INIT_PVB_AIOPEN":       (445, 359),
    "INIT_PVB_AIDNOT":       (451, 323),
    "FINAL_CV1_CLOSED":      (138, 306),
    "FINAL_CV2_CLOSED":      (236, 301),
    "FINAL_RV_OPENED":       (331, 296),
    "FINAL_PVB_SAT":         (450, 290),
    "JAX_PASSED":            (300, 106),
    "JAX_FAILED":            (358, 108),
}

JAX_NEXT_REPORT_KEEP = {
    "premises_name", "owner_name", "service_address", "mailing_address",
    "physical_location", "contact_phone", "jea_account", "meter_number",
    "comm_test_purpose", "comm_service_type", "comm_reclaim",
    "res_test_purpose", "res_service_type", "res_reclaim",
    "manufacturer", "model_number", "size", "device_type",
    "init_tester_name", "init_company", "init_cert",
    "final_tester_name", "final_company", "final_cert",
}
JAX_NEW_JOB_KEEP = {
    "init_tester_name", "init_company", "init_cert",
    "final_tester_name", "final_company", "final_cert",
}

JAX_GREEN_WIDGET_KEYS = [
    "j_sn", "j_id",
    "j_icv1p", "j_icv2p", "j_irvp", "j_ipvbp", "j_itd",
    "j_fcv1p", "j_fcv2p", "j_frvp",
    "j_rep", "j_sd", "j_ares",
    "init_cv1_result", "init_cv2_result", "init_rv_result", "init_pvb_result",
    "final_cv1_result", "final_cv2_result", "final_rv_result", "final_pvb_result",
    "assembly_result",
]
JAX_BLUE_WIDGET_KEYS = [
    "j_prem", "j_own", "j_sa", "j_ma", "j_pl", "j_ph", "j_acct", "j_meter",
    "j_dt", "j_mfg", "j_sz", "j_mn",
    "j_ctp", "j_cst", "j_crc", "j_rtp", "j_rst", "j_rrc",
]
JAX_TESTER_WIDGET_KEYS = [
    "j_itn", "j_ico", "j_ic", "j_itd2",
    "j_rb", "j_rco", "j_rc", "j_rd",
    "j_ftn", "j_fco", "j_fc", "j_ftd2",
]

# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def draw_x(c, bx, by, size=3.8):
    c.setStrokeColorRGB(1, 0, 0)
    c.setLineWidth(2.0)
    c.line(bx-size, by-size, bx+size, by+size)
    c.line(bx+size, by-size, bx-size, by+size)


def put_text(c, val, x, y, sz=8):
    if val:
        c.setFont("Helvetica-Bold", sz)
        c.setFillColorRGB(1, 0, 0)
        c.drawString(x, y, str(val))


def wrap_text(text, w=58):
    words = text.split()
    lines, line = [], ""
    for word in words:
        t = (line + " " + word).strip()
        if len(t) > w:
            lines.append(line.strip())
            line = word
        else:
            line = t
    if line:
        lines.append(line)
    return lines


# ---------------------------------------------------------------------------
# Tap-to-clear input helper  (BLUE fields)
# ---------------------------------------------------------------------------

def tap_clear_input(label, form_key, field_key, widget_key, **kwargs):
    form = st.session_state[form_key]
    focus_flag = f"{widget_key}_focused"

    if st.session_state.get(f"{widget_key}_do_clear"):
        form[field_key] = ""
        st.session_state.pop(f"{widget_key}_do_clear", None)
        st.session_state.pop(widget_key, None)

    def _on_change():
        if not st.session_state.get(focus_flag):
            st.session_state[f"{widget_key}_do_clear"] = True
            st.session_state[focus_flag] = True

    col_input, col_btn = st.columns([5, 1])
    with col_input:
        val = st.text_input(
            label,
            value=form.get(field_key, ""),
            key=widget_key,
            on_change=_on_change,
            **kwargs,
        )
        form[field_key] = st.session_state.get(widget_key, "")

    with col_btn:
        st.write("")
        if st.button("✕", key=f"clr_{widget_key}", help=f"Clear {label}"):
            form[field_key] = ""
            st.session_state.pop(widget_key, None)
            st.session_state.pop(focus_flag, None)
            st.rerun()

    return val


def reset_blue_focus_flags():
    for wk in UNITED_BLUE_WIDGET_KEYS:
        st.session_state.pop(f"{wk}_focused", None)
        st.session_state.pop(f"{wk}_do_clear", None)


# ---------------------------------------------------------------------------
# Standard clearable input (GREEN / YELLOW fields)
# ---------------------------------------------------------------------------

def clearable_input(label, form_key, field_key, widget_key, **kwargs):
    form = st.session_state[form_key]

    col_input, col_btn = st.columns([5, 1])
    with col_input:
        val = st.text_input(
            label,
            value=form.get(field_key, ""),
            key=widget_key,
            **kwargs,
        )
        form[field_key] = st.session_state.get(widget_key, "")

    with col_btn:
        st.write("")
        if st.button("✕", key=f"clr_{widget_key}", help=f"Clear {label}"):
            st.session_state[form_key][field_key] = ""
            st.session_state.pop(widget_key, None)
            st.rerun()

    return val


# ---------------------------------------------------------------------------
# Signature helpers
# ---------------------------------------------------------------------------

def get_signature_image_reader():
    b64 = st.session_state.get("signature_b64")
    if b64:
        try:
            buf = BytesIO(base64.b64decode(b64))
            buf.seek(0)
            return ImageReader(buf)
        except Exception:
            return None
    return None


def clear_signature():
    st.session_state.pop("signature_b64", None)


def save_signature(img_array):
    buf = BytesIO()
    img = Image.fromarray(img_array.astype("uint8"), "RGBA")
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    st.session_state["signature_b64"] = b64
    return b64


# ===========================================================================
# Excel export helpers
# ===========================================================================

UNITED_EXCEL_COLS = [
    ("Date",              "date"),
    ("Branch",            "branch"),
    ("AHJ",              "ahj"),
    ("Customer Name",     "customer_name"),
    ("Street Address",    "street_address"),
    ("Location",         "location"),
    ("Serial Number",    "serial_number"),
    ("Manufacturer",     "manufacturer"),
    ("Model",            "model"),
    ("Size",             "size"),
    ("Assembly Type",    "assembly_type"),
    ("System Service",   "system_service"),
    ("Bypass",           "bypass"),
    ("Test Date",        "test_date"),
    ("CV1 Result",       "cv1_result"),
    ("CV1 DP (psi)",     "cv1_dp"),
    ("CV2 Result",       "cv2_result"),
    ("CV2 DP (psi)",     "cv2_dp"),
    ("RV Result",        "rv_result"),
    ("RV Opened At (psi)","rv_psi"),
    ("RV Outlet",        "rv_out_result"),
    ("RV Inlet",         "rv_in_result"),
    ("PVB Air Inlet",    "pvb_ai_result"),
    ("PVB AI (psi)",     "pvb_ai_psi"),
    ("PVB CV Result",    "pvb_cv_result"),
    ("PVB CV (psi)",     "pvb_cv_psi"),
    ("Pass / Fail",      "assembly_result"),
    ("Repairs",          "repair_desc"),
    ("Technician",       "technician"),
    ("Cert No.",         "cert_no"),
    ("Re-Cert Date",     "recert"),
    ("Gauge Mfg",        "gauge_mfg"),
    ("Gauge Serial",     "gauge_serial"),
    ("Date Calibrated",  "date_cal"),
]

JAX_EXCEL_COLS = [
    ("Premises Name",      "premises_name"),
    ("Owner Name",         "owner_name"),
    ("Service Address",    "service_address"),
    ("Mailing Address",    "mailing_address"),
    ("Physical Location",  "physical_location"),
    ("Contact Phone",      "contact_phone"),
    ("JEA Account",        "jea_account"),
    ("Meter Number",       "meter_number"),
    ("Device Type",        "device_type"),
    ("Manufacturer",       "manufacturer"),
    ("Size",               "size"),
    ("Model Number",       "model_number"),
    ("Serial Number",      "serial_number"),
    ("Install Date",       "install_date"),
    ("Comm Test Purpose",  "comm_test_purpose"),
    ("Comm Service Type",  "comm_service_type"),
    ("Comm Reclaim",       "comm_reclaim"),
    ("Res Test Purpose",   "res_test_purpose"),
    ("Res Service Type",   "res_service_type"),
    ("Res Reclaim",        "res_reclaim"),
    ("Init CV1 Result",    "init_cv1_result"),
    ("Init CV1 (psi)",     "init_cv1_psi"),
    ("Init CV2 Result",    "init_cv2_result"),
    ("Init CV2 (psi)",     "init_cv2_psi"),
    ("Init RV Result",     "init_rv_result"),
    ("Init RV (psi)",      "init_rv_psi"),
    ("Init PVB Result",    "init_pvb_result"),
    ("Init PVB (psi)",     "init_pvb_psi"),
    ("Final CV1 Result",   "final_cv1_result"),
    ("Final CV1 (psi)",    "final_cv1_psi"),
    ("Final CV2 Result",   "final_cv2_result"),
    ("Final CV2 (psi)",    "final_cv2_psi"),
    ("Final RV Result",    "final_rv_result"),
    ("Final RV (psi)",     "final_rv_psi"),
    ("Final PVB Result",   "final_pvb_result"),
    ("Pass / Fail",        "assembly_result"),
    ("Repairs",            "repairs"),
    ("Init Tester",        "init_tester_name"),
    ("Init Company",       "init_company"),
    ("Init Cert",          "init_cert"),
    ("Init Test Date",     "init_test_date"),
    ("Repaired By",        "repaired_by"),
    ("Repair Company",     "repair_company"),
    ("Repair Cert",        "repair_cert"),
    ("Repair Date",        "repair_date"),
    ("Final Tester",       "final_tester_name"),
    ("Final Company",      "final_company"),
    ("Final Cert",         "final_cert"),
    ("Final Test Date",    "final_test_date"),
    ("Signature Date",     "signature_date"),
]


def _style_header_row(ws, num_cols):
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    thin = Side(style="thin", color="AAAAAA")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
    ws.row_dimensions[1].height = 30


def _style_data_row(ws, row_idx, num_cols, pass_fail_col_idx):
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    bg = "EEF2F7" if row_idx % 2 == 0 else "FFFFFF"
    fill = PatternFill("solid", fgColor=bg)

    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = fill
        cell.font = Font(size=9)
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        cell.border = border

    pf_cell = ws.cell(row=row_idx, column=pass_fail_col_idx)
    val = str(pf_cell.value or "").upper()
    if val == "PASSED":
        pf_cell.fill = PatternFill("solid", fgColor="C6EFCE")
        pf_cell.font = Font(bold=True, color="276221", size=9)
    elif val == "FAILED":
        pf_cell.fill = PatternFill("solid", fgColor="FFC7CE")
        pf_cell.font = Font(bold=True, color="9C0006", size=9)


def _build_summary_sheet(wb, job_folder):
    rows = []
    for item in job_folder:
        fd = item.get("form_data", {})
        ft = item.get("form_type", "united")

        if ft == "united":
            form_label = "United Fire"
            location = (
                fd.get("customer_name", "")
                or fd.get("street_address", "")
                or fd.get("location", "")
            ).strip()
        else:
            form_label = "Jacksonville"
            location = (
                fd.get("premises_name", "")
                or fd.get("service_address", "")
                or fd.get("physical_location", "")
            ).strip()

        result = str(fd.get("assembly_result", "")).strip().upper() or "—"
        rows.append((form_label, location or "—", result))

    rows.sort(key=lambda r: (r[0], r[1]))

    ws = wb.create_sheet("Summary", 0)

    headers = ["Form Type", "Location / Customer", "Pass / Fail"]
    ws.append(headers)
    _style_header_row(ws, len(headers))

    passed_count = 0
    failed_count = 0

    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for i, (form_label, location, result) in enumerate(rows, start=2):
        ws.append([form_label, location, result])

        bg = "EEF2F7" if i % 2 == 0 else "FFFFFF"
        base_fill = PatternFill("solid", fgColor=bg)

        ft_cell = ws.cell(row=i, column=1)
        ft_cell.fill = base_fill
        ft_cell.font = Font(size=10)
        ft_cell.alignment = Alignment(vertical="center")
        ft_cell.border = border

        loc_cell = ws.cell(row=i, column=2)
        loc_cell.fill = base_fill
        loc_cell.font = Font(size=10)
        loc_cell.alignment = Alignment(vertical="center")
        loc_cell.border = border

        pf_cell = ws.cell(row=i, column=3)
        pf_cell.border = border
        pf_cell.alignment = Alignment(horizontal="center", vertical="center")
        if result == "PASSED":
            pf_cell.fill = PatternFill("solid", fgColor="C6EFCE")
            pf_cell.font = Font(bold=True, color="276221", size=10)
            passed_count += 1
        elif result == "FAILED":
            pf_cell.fill = PatternFill("solid", fgColor="FFC7CE")
            pf_cell.font = Font(bold=True, color="9C0006", size=10)
            failed_count += 1
        else:
            pf_cell.fill = base_fill
            pf_cell.font = Font(size=10, color="888888")

    spacer_row = len(rows) + 2
    ws.append([])
    total_row = spacer_row + 1
    ws.cell(row=total_row, column=1).value = f"Total Reports: {len(rows)}"
    ws.cell(row=total_row, column=1).font = Font(bold=True, size=10)
    ws.cell(row=total_row, column=2).value = f"Passed: {passed_count}"
    ws.cell(row=total_row, column=2).font = Font(bold=True, color="276221", size=10)
    ws.cell(row=total_row, column=3).value = f"Failed: {failed_count}"
    ws.cell(row=total_row, column=3).font = Font(bold=True, color="9C0006", size=10)

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 14
    ws.row_dimensions[1].height = 28


def build_excel_for_job(job_folder: list, job_name: str) -> bytes:
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    _build_summary_sheet(wb, job_folder)

    ws_u = wb.create_sheet("United Fire")
    u_headers = [h for h, _ in UNITED_EXCEL_COLS]
    ws_u.append(u_headers)
    pf_u = next((i+1 for i, (h, _) in enumerate(UNITED_EXCEL_COLS) if h == "Pass / Fail"), 1)
    _style_header_row(ws_u, len(u_headers))

    ws_j = wb.create_sheet("Jacksonville")
    j_headers = [h for h, _ in JAX_EXCEL_COLS]
    ws_j.append(j_headers)
    pf_j = next((i+1 for i, (h, _) in enumerate(JAX_EXCEL_COLS) if h == "Pass / Fail"), 1)
    _style_header_row(ws_j, len(j_headers))

    u_row = 2
    j_row = 2

    for item in job_folder:
        form_data = item.get("form_data", {})
        form_type = item.get("form_type", "united")

        if form_type == "united":
            row = [form_data.get(fk, "") for _, fk in UNITED_EXCEL_COLS]
            ws_u.append(row)
            _style_data_row(ws_u, u_row, len(u_headers), pf_u)
            u_row += 1
        else:
            row = [form_data.get(fk, "") for _, fk in JAX_EXCEL_COLS]
            ws_j.append(row)
            _style_data_row(ws_j, j_row, len(j_headers), pf_j)
            j_row += 1

    for ws in [ws_u, ws_j]:
        for col_cells in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col_cells), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 40)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ---------------------------------------------------------------------------
# Job folder
# ---------------------------------------------------------------------------

def _init_job_folder():
    if "job_folder" not in st.session_state:
        st.session_state.job_folder = []


def add_to_job_folder(pdf_bytes: bytes, filename: str, form_data: dict, form_type: str):
    _init_job_folder()
    st.session_state.job_folder = [
        f for f in st.session_state.job_folder if f["name"] != filename
    ]
    st.session_state.job_folder.append({
        "name": filename,
        "bytes": pdf_bytes,
        "form_data": dict(form_data),
        "form_type": form_type,
    })


def build_zip() -> bytes:
    folder = st.session_state.job_folder
    job_name = st.session_state.get("job_folder_name", "").strip() or "Job"

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in folder:
            zf.writestr(item["name"], item["bytes"])
        try:
            excel_bytes = build_excel_for_job(folder, job_name)
            excel_name = re.sub(r"[^\w\s\-]", "", job_name).strip() or "Job"
            zf.writestr(f"{excel_name} - Report Summary.xlsx", excel_bytes)
        except Exception as e:
            zf.writestr("excel_error.txt", f"Excel export failed: {e}")

    buf.seek(0)
    return buf.read()


def render_job_folder_sidebar():
    _init_job_folder()
    folder = st.session_state.job_folder
    st.sidebar.markdown("---")
    st.sidebar.subheader("📁 Job Folder")

    if "job_folder_name" not in st.session_state:
        st.session_state.job_folder_name = ""
    st.session_state.job_folder_name = st.sidebar.text_input(
        "Job name (for ZIP file)",
        value=st.session_state.job_folder_name,
        placeholder="e.g. Smith Residence",
        key="job_name_input",
    )

    if not folder:
        st.sidebar.caption("No reports yet. Generate a PDF to add it here.")
        return

    st.sidebar.caption(f"{len(folder)} report(s) in this job folder:")
    for item in folder:
        st.sidebar.markdown(f"• {item['name']}")

    zip_bytes = build_zip()
    entered_name = st.session_state.job_folder_name.strip()
    if entered_name:
        zip_name = re.sub(r"[^\w\s\-]", "", entered_name).strip()
    else:
        first_name = folder[0]["name"].replace(".pdf", "")
        zip_name = re.sub(r" - [^-]+$", "", first_name)
        zip_name = re.sub(r"[^\w\s\-]", "", zip_name).strip() or "Job"
    zip_filename = f"{zip_name} - Reports.zip"

    st.sidebar.download_button(
        label=f"📦 Download ZIP ({len(folder)} report{'s' if len(folder)!=1 else ''} + Excel)",
        data=zip_bytes,
        file_name=zip_filename,
        mime="application/zip",
        use_container_width=True,
        key=f"zip_dl_{len(folder)}",
    )

    st.sidebar.caption("📱 iPhone/iPad: tap ZIP button → Share → Save to Files")
    st.sidebar.caption("📊 ZIP includes an Excel summary of all reports in this job.")

    if st.sidebar.button("🗑️ Clear Job Folder", use_container_width=True, key="clear_folder"):
        st.session_state.job_folder = []
        st.session_state.job_folder_name = ""
        st.rerun()


# ---------------------------------------------------------------------------
# PDF merge helper
# ---------------------------------------------------------------------------

def _merge_overlay(template_path: str, overlay_buf: BytesIO) -> bytes:
    overlay_buf.seek(0)
    overlay_reader = PdfReader(overlay_buf)
    template_reader = PdfReader(template_path)

    writer = PdfWriter()
    template_page = template_reader.pages[0]
    overlay_page = overlay_reader.pages[0]
    template_page.merge_page(overlay_page)
    writer.add_page(template_page)

    out = BytesIO()
    writer.write(out)
    out.seek(0)
    return out.read()


# ---------------------------------------------------------------------------
# United Fire PDF generator
# ---------------------------------------------------------------------------

def generate_united_pdf(form):
    overlay_buf = BytesIO()
    c = canvas.Canvas(overlay_buf, pagesize=(PAGE_W, PAGE_H))

    for field, (x, y, sz) in UNITED_TEXT_FIELDS.items():
        put_text(c, form.get(field, ""), x, y, sz)

    atype = form.get("assembly_type", "")
    if atype in UNITED_CHECKBOXES:
        draw_x(c, *UNITED_CHECKBOXES[atype])

    svc = form.get("system_service", "")
    if svc in UNITED_CHECKBOXES:
        draw_x(c, *UNITED_CHECKBOXES[svc])

    bp = form.get("bypass", "")
    if bp == "YES": draw_x(c, *UNITED_CHECKBOXES["BYPASS_YES"])
    elif bp == "NO": draw_x(c, *UNITED_CHECKBOXES["BYPASS_NO"])

    cv1 = form.get("cv1_result", "")
    if cv1 == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["CV1_CLOSED"])
    elif cv1 == "Leaked": draw_x(c, *UNITED_CHECKBOXES["CV1_LEAKED"])

    cv2 = form.get("cv2_result", "")
    if cv2 == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["CV2_CLOSED"])
    elif cv2 == "Leaked": draw_x(c, *UNITED_CHECKBOXES["CV2_LEAKED"])

    rv = form.get("rv_result", "")
    if rv == "Opened At": draw_x(c, *UNITED_CHECKBOXES["RV_OPENED"])
    elif rv == "Did Not Open": draw_x(c, *UNITED_CHECKBOXES["RV_DIDNOTOPEN"])

    rvo = form.get("rv_out_result", "")
    if rvo == "Closed": draw_x(c, *UNITED_CHECKBOXES["RV_OUT_CLOSED"])
    elif rvo == "Leaked": draw_x(c, *UNITED_CHECKBOXES["RV_OUT_LEAKED"])

    rvi = form.get("rv_in_result", "")
    if rvi == "Closed": draw_x(c, *UNITED_CHECKBOXES["RV_IN_CLOSED"])
    elif rvi == "Leaked": draw_x(c, *UNITED_CHECKBOXES["RV_IN_LEAKED"])

    pvb_ai = form.get("pvb_ai_result", "")
    if pvb_ai == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["PVB_AI_CLOSED"])
    elif pvb_ai == "Opened At": draw_x(c, *UNITED_CHECKBOXES["PVB_AI_OPENED"])

    pvb_cv = form.get("pvb_cv_result", "")
    if pvb_cv == "Leaked": draw_x(c, *UNITED_CHECKBOXES["PVB_CV_LEAKED"])
    elif pvb_cv == "Held At": draw_x(c, *UNITED_CHECKBOXES["PVB_CV_HELD"])

    result = form.get("assembly_result", "")
    if result == "PASSED": draw_x(c, *UNITED_CHECKBOXES["PASSED"])
    elif result == "FAILED": draw_x(c, *UNITED_CHECKBOXES["FAILED"])

    rx, ry, rh, rmax, rw = UNITED_REPAIR_BOX
    for i, ln in enumerate(wrap_text(form.get("repair_desc", ""), rw)[:rmax]):
        put_text(c, ln, rx, ry - i * rh, 7)

    sig_ir = get_signature_image_reader()
    if sig_ir:
        c.drawImage(sig_ir, UNITED_SIG_X, UNITED_SIG_Y,
                    width=UNITED_SIG_W, height=UNITED_SIG_H, mask="auto")

    c.save()
    return _merge_overlay(TEMPLATE_UNITED, overlay_buf)


def generate_jax_pdf(form):
    overlay_buf = BytesIO()
    c = canvas.Canvas(overlay_buf, pagesize=(JAX_PAGE_W, JAX_PAGE_H))

    for field, (x, y, sz) in JAX_TEXT_FIELDS.items():
        put_text(c, form.get(field, ""), x, y, sz)

    ctp = form.get("comm_test_purpose", "")
    for key, label in [("COMM_ANNUAL","Annual"),("COMM_REPAIR","Repair"),
                       ("COMM_REPLACEMENT","Replacement"),("COMM_NEW_INSTALL","New Installation")]:
        if ctp == label:
            draw_x(c, *JAX_CHECKBOXES[key])

    cst = form.get("comm_service_type", "")
    for key, label in [("COMM_FIRE","Fire"),("COMM_IRRIGATION","Irrigation"),
                       ("COMM_PROCESS","Process/Isolation"),("COMM_POTABLE","Potable"),
                       ("COMM_FIRE_BYPASS","Fire bypass")]:
        if cst == label:
            draw_x(c, *JAX_CHECKBOXES[key])

    rcl = form.get("comm_reclaim", "")
    if rcl == "Yes": draw_x(c, *JAX_CHECKBOXES["RECLAIM_YES"])
    elif rcl == "No": draw_x(c, *JAX_CHECKBOXES["RECLAIM_NO"])

    rtp = form.get("res_test_purpose", "")
    for key, label in [("RES_ANNUAL","Annual"),("RES_REPAIR","Repair"),
                       ("RES_REPLACEMENT","Replacement"),("RES_NEW_INSTALL","New Installation")]:
        if rtp == label:
            draw_x(c, *JAX_CHECKBOXES[key])

    rst = form.get("res_service_type", "")
    for key, label in [("RES_POTABLE","Potable"),("RES_IRRIGATION","Irrigation / Is reclaimed")]:
        if rst == label:
            draw_x(c, *JAX_CHECKBOXES[key])

    res_rcl = form.get("res_reclaim", "")
    if res_rcl == "Yes": draw_x(c, *JAX_CHECKBOXES["RES_RECLAIM_YES"])
    elif res_rcl == "No": draw_x(c, *JAX_CHECKBOXES["RES_RECLAIM_NO"])

    icv1 = form.get("init_cv1_result", "")
    if icv1 == "Closed Tight":
        draw_x(c, *JAX_CHECKBOXES["INIT_CV1_CLOSED"])

    icv2 = form.get("init_cv2_result", "")
    if icv2 == "Closed Tight":
        draw_x(c, *JAX_CHECKBOXES["INIT_CV2_CLOSED"])

    irv = form.get("init_rv_result", "")
    if irv == "Opened At": draw_x(c, *JAX_CHECKBOXES["INIT_RV_OPENED"])
    elif irv == "Did Not Open": draw_x(c, *JAX_CHECKBOXES["INIT_RV_DIDNOT"])

    ipvb = form.get("init_pvb_result", "")
    if ipvb == "Air inlet opened at": draw_x(c, *JAX_CHECKBOXES["INIT_PVB_AIOPEN"])
    elif ipvb == "Did not open": draw_x(c, *JAX_CHECKBOXES["INIT_PVB_AIDNOT"])

    fcv1 = form.get("final_cv1_result", "")
    if fcv1 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["FINAL_CV1_CLOSED"])

    fcv2 = form.get("final_cv2_result", "")
    if fcv2 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["FINAL_CV2_CLOSED"])

    frv = form.get("final_rv_result", "")
    if frv == "Opened At": draw_x(c, *JAX_CHECKBOXES["FINAL_RV_OPENED"])

    fpvb = form.get("final_pvb_result", "")
    if fpvb == "Satisfactory": draw_x(c, *JAX_CHECKBOXES["FINAL_PVB_SAT"])

    result = form.get("assembly_result", "")
    if result == "PASSED": draw_x(c, *JAX_CHECKBOXES["JAX_PASSED"])
    elif result == "FAILED": draw_x(c, *JAX_CHECKBOXES["JAX_FAILED"])

    sig_ir = get_signature_image_reader()
    if sig_ir:
        c.drawImage(sig_ir, JAX_SIG_X, JAX_SIG_Y,
                    width=JAX_SIG_W, height=JAX_SIG_H, mask="auto")

    c.save()
    return _merge_overlay(TEMPLATE_JAX, overlay_buf)


# ---------------------------------------------------------------------------
# Delivery
# ---------------------------------------------------------------------------

def safe_filename(customer, street, location, prefix=""):
    def clean(s): return re.sub(r"[^\w\s\-]", "", s or "").strip()
    parts = [clean(customer) or "Customer",
             clean(street) or "Address",
             clean(location) or "Location"]
    name = " - ".join(parts) + ".pdf"
    return (prefix + " " + name).strip() if prefix else name


def deliver_pdf(pdf_bytes: bytes, filename: str):
    st.download_button(
        label="📥 Download PDF",
        data=pdf_bytes,
        file_name=filename,
        mime="application/pdf",
        use_container_width=True,
        key=f"pdf_dl_{filename}_{len(st.session_state.get('job_folder', []))}",
    )


# ---------------------------------------------------------------------------
# Auto-save helper (used by JAX form)
# ---------------------------------------------------------------------------

def _sync(form, key, widget_key):
    form[key] = st.session_state.get(widget_key, "")


def text_input_autosave(label, form, key, widget_key, **kwargs):
    return st.text_input(
        label,
        value=form.get(key, ""),
        key=widget_key,
        on_change=_sync,
        args=(form, key, widget_key),
        **kwargs,
    )


def _radio(label, options, key, form, **kwargs):
    opts = [""] + list(options)
    current = form.get(key, "")
    idx = opts.index(current) if current in opts else 0

    def _sync_radio(_form=form, _key=key):
        _form[_key] = st.session_state.get(_key, "")

    chosen = st.radio(label, opts, index=idx, key=key,
                      format_func=lambda x: "—" if x == "" else x,
                      on_change=_sync_radio, **kwargs)
    form[key] = chosen
    return chosen


# ===========================================================================
# Technician selector sidebar widget
# ===========================================================================

def render_technician_sidebar():
    _init_technicians()
    st.sidebar.markdown("---")
    st.sidebar.subheader("👷 Technician")

    names = get_technician_names()
    current = st.session_state.get("active_technician", "")
    idx = names.index(current) if current in names else 0

    selected = st.sidebar.selectbox(
        "Select technician",
        names,
        index=idx,
        format_func=lambda x: "— select —" if x == "" else x,
        key="tech_selector",
    )

    if selected != st.session_state.get("active_technician", ""):
        # Technician changed — load profile and inject into both forms
        profile = get_technician_profile(selected) if selected else {}
        st.session_state["active_technician"] = selected
        st.session_state["signature_b64"] = profile.get("signature_b64", "")

        # Inject tester fields into United Fire form (clear widget keys so new values render)
        if "united_form" in st.session_state:
            for k in TESTER_KEYS:
                st.session_state["united_form"][k] = profile.get(k, "")
            for wk in UNITED_TESTER_WIDGET_KEYS:
                st.session_state.pop(wk, None)

        # Inject tester fields into JAX form
        if "jax_form" in st.session_state:
            jax_tech_map = {
                "init_tester_name": "technician",
                "init_cert": "cert_no",
                "final_tester_name": "technician",
                "final_cert": "cert_no",
            }
            for form_key, profile_key in jax_tech_map.items():
                st.session_state["jax_form"][form_key] = profile.get(profile_key, "")
            for wk in JAX_TESTER_WIDGET_KEYS:
                st.session_state.pop(wk, None)

        # Store tester_defaults for fresh form initialization
        st.session_state["tester_defaults"] = {k: profile.get(k, "") for k in TESTER_KEYS}

        st.rerun()

    if not selected:
        st.sidebar.caption("Select your name to load your profile.")
        return

    with st.sidebar.expander("✏️ Edit My Profile", expanded=False):
        profile = get_technician_profile(selected)

        new_name   = st.text_input("Display name",        value=profile.get("technician", selected), key="pe_name")
        new_cert   = st.text_input("Certification No.",   value=profile.get("cert_no", ""),           key="pe_cert")
        new_recert = st.text_input("Re-Cert Due Date",    value=profile.get("recert", ""),             key="pe_recert")
        new_gmfg   = st.text_input("Gauge Manufacturer",  value=profile.get("gauge_mfg", ""),         key="pe_gmfg")
        new_gsn    = st.text_input("Gauge Serial #",      value=profile.get("gauge_serial", ""),      key="pe_gsn")
        new_gcal   = st.text_input("Date Calibrated",     value=profile.get("date_cal", ""),          key="pe_gcal")

        st.markdown("**Signature**")
        if st.session_state.get("signature_b64"):
            st.success("Signature on file ✓")
            if st.button("🗑️ Clear signature", key="pe_clrsig"):
                clear_signature()
                st.rerun()
        else:
            try:
                from streamlit_drawable_canvas import st_canvas
                sig_canvas = st_canvas(
                    fill_color="rgba(255,255,255,0)",
                    stroke_width=2,
                    stroke_color="#cc0000",
                    background_color="#ffffff",
                    height=80, width=220,
                    drawing_mode="freedraw",
                    key="pe_sig_canvas",
                )
                if st.button("💾 Save signature", key="pe_savesig"):
                    if sig_canvas.image_data is not None and sig_canvas.image_data.max() > 0:
                        b64 = save_signature(sig_canvas.image_data)
                        st.session_state["signature_b64"] = b64
                    else:
                        st.warning("Draw your signature first.")
            except ImportError:
                st.warning("`streamlit-drawable-canvas` not installed.")

        if st.button("💾 Save Profile to Server", key="pe_save", use_container_width=True):
            updated_profile = {
                "technician":    new_name,
                "cert_no":       new_cert,
                "recert":        new_recert,
                "gauge_mfg":     new_gmfg,
                "gauge_serial":  new_gsn,
                "date_cal":      new_gcal,
                "signature_b64": st.session_state.get("signature_b64", ""),
            }
            techs = st.session_state["technicians"]
            if new_name != selected and new_name.strip():
                techs.pop(selected, None)
                key_name = new_name
            else:
                key_name = selected
            techs[key_name] = updated_profile
            ok, msg = save_technicians_to_github(
                techs, st.session_state.get("technicians_sha")
            )
            if ok:
                st.success(msg)
                st.session_state["active_technician"] = key_name
                _, new_sha = load_technicians_from_github()
                st.session_state["technicians_sha"] = new_sha
                # Re-inject updated profile into active forms
                for k in TESTER_KEYS:
                    if "united_form" in st.session_state:
                        st.session_state["united_form"][k] = updated_profile.get(k, "")
                for wk in UNITED_TESTER_WIDGET_KEYS:
                    st.session_state.pop(wk, None)
            else:
                st.error(msg)

    with st.sidebar.expander("➕ Add New Technician", expanded=False):
        new_tech_name = st.text_input("Full name", key="new_tech_name")
        if st.button("Add", key="new_tech_add"):
            name = new_tech_name.strip()
            if name and name not in st.session_state["technicians"]:
                blank = {
                    "technician": name, "cert_no": "", "recert": "",
                    "gauge_mfg": "", "gauge_serial": "", "date_cal": "",
                    "signature_b64": "",
                }
                st.session_state["technicians"][name] = blank
                ok, msg = save_technicians_to_github(
                    st.session_state["technicians"],
                    st.session_state.get("technicians_sha")
                )
                if ok:
                    st.success(f"Added {name}")
                    _, new_sha = load_technicians_from_github()
                    st.session_state["technicians_sha"] = new_sha
                    st.rerun()
                else:
                    st.error(msg)
            elif name in st.session_state["technicians"]:
                st.warning("That name already exists.")
            else:
                st.warning("Enter a name first.")


# ===========================================================================
# App layout
# ===========================================================================

st.set_page_config(page_title="Backflow Report", page_icon="🔧", layout="wide")
_init_job_folder()
_init_technicians()

render_technician_sidebar()
render_job_folder_sidebar()

st.title("🔧 Backflow Preventer Test Report")

# --------------------------------------------------------------------------