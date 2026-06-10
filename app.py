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

# Widget keys that correspond to GREEN fields (need session state cleared after PDF)
UNITED_GREEN_WIDGET_KEYS = [
    "u_loc", "u_cv1dp", "u_cv2dp", "u_rvpsi", "u_aipsi", "u_cvpsi",
    "u_tdate", "u_rep",
    "cv1_result", "cv2_result", "rv_result", "rv_out_result",
    "rv_in_result", "pvb_ai_result", "pvb_cv_result", "assembly_result",
    # selectbox / radio widget keys that map to green fields
    "u_res", "u_bp",
]

# Widget keys that correspond to YELLOW (tester) fields
UNITED_TESTER_WIDGET_KEYS = ["u_gmfg", "u_gsn", "u_cal", "u_tech", "u_cert", "u_recert"]

# Widget keys that correspond to BLUE fields
UNITED_BLUE_WIDGET_KEYS = [
    "u_date", "u_branch", "u_ahj", "u_cust", "u_addr",
    "u_sn", "u_mfg", "u_mdl", "u_sz",
    "u_atype", "u_ss",
]

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

# All JAX widget keys — cleared when form resets or tester profile injects
JAX_GREEN_WIDGET_KEYS = [
    "j_sn", "j_id",
    "j_icv1p", "j_icv2p", "j_irvp", "j_ipvbp", "j_itd",
    "j_fcv1p", "j_fcv2p", "j_frvp",
    "j_rep", "j_sd",
    "j_ares",
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
# Tap-to-clear input helper
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
# Standard clearable input (no tap-to-clear)
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
    ("Final Tester",       "final_tester