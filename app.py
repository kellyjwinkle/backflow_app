import streamlit as st
import streamlit.components.v1 as components
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import json, os, re, base64, tempfile, zipfile, requests
from datetime import date, datetime
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

UNITED_GREEN_WIDGET_KEYS = [
    "u_loc", "u_cv1dp", "u_cv2dp", "u_rvpsi", "u_aipsi", "u_cvpsi",
    "u_rep",
    "cv1_result", "cv2_result", "rv_result", "rv_out_result",
    "rv_in_result", "pvb_ai_result", "pvb_cv_result", "assembly_result",
]

UNITED_TESTER_WIDGET_KEYS = ["u_gmfg", "u_gsn", "u_cal", "u_tech", "u_cert", "u_recert"]

UNITED_BLUE_WIDGET_KEYS = [
    "u_branch", "u_ahj", "u_cust", "u_addr",
    "u_sn", "u_mfg", "u_mdl", "u_sz",
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
    "signature_date", "init_test_date", "final_test_date", "repair_date",
}
JAX_NEW_JOB_KEEP = {
    "init_tester_name", "init_company", "init_cert",
    "final_tester_name", "final_company", "final_cert",
}

JAX_GREEN_WIDGET_KEYS = [
    "j_sn", "j_id",
    "j_icv1p", "j_icv2p", "j_irvp", "j_ipvbp",
    "j_fcv1p", "j_fcv2p", "j_frvp",
    "j_rep", "j_ares",
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
    "j_itn", "j_ico", "j_ic",
    "j_rb", "j_rco", "j_rc",
    "j_ftn", "j_fco", "j_fc",
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
# Input helpers
# ---------------------------------------------------------------------------

def tap_clear_input(label, form_key, field_key, widget_key, **kwargs):
    form = st.session_state[form_key]
    val = st.text_input(
        label,
        value=form.get(field_key, ""),
        key=widget_key,
        **kwargs,
    )
    form[field_key] = val
    return val


def reset_blue_focus_flags():
    return


def clearable_input(label, form_key, field_key, widget_key, **kwargs):
    form = st.session_state[form_key]
    val = st.text_input(
        label,
        value=form.get(field_key, ""),
        key=widget_key,
        **kwargs,
    )
    form[field_key] = val
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


def _date_to_string(value):
    if not value:
        return ""
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return value.strftime("%m/%d/%Y")
    return str(value)


def synced_date_input(label, form_key, source_key, widget_key, target_fields):
    form = st.session_state[form_key]
    current = form.get(source_key, "")
    default_date = date.today()

    if current:
        for fmt in ("%m/%d/%Y", "%Y-%m-%d"):
            try:
                default_date = datetime.strptime(str(current), fmt).date()
                break
            except Exception:
                pass

    picked = st.date_input(label, value=default_date, key=widget_key)
    picked_str = _date_to_string(picked)
    form[source_key] = picked_str
    for field in target_fields:
        form[field] = picked_str
    return picked


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

        result = str(fd.get("assembly_result", "")).strip().upper() or "\u2014"
        rows.append((form_label, location or "\u2014", result))

    rows.sort(key=lambda r: (r[0], r[1]))

    ws = wb.create_sheet("Summary", 0)

    headers = ["Form Type", "Location / Customer", "Pass / Fail"]
    ws.append(headers)
    _style_header_row(ws, len(headers))

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
        elif result == "FAILED":
            pf_cell.fill = PatternFill("solid", fgColor="FFC7CE")
            pf_cell.font = Font(bold=True, color="9C0006", size=10)
        else:
            pf_cell.fill = base_fill
            pf_cell.font = Font(size=10)

        ws.row_dimensions[i].height = 18

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 14
    ws.freeze_panes = "A2"


def export_jobs_to_excel(job_folder):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _build_summary_sheet(wb, job_folder)

    united_rows = [item for item in job_folder if item.get("form_type", "united") == "united"]
    jax_rows    = [item for item in job_folder if item.get("form_type") == "jax"]

    if united_rows:
        ws_u = wb.create_sheet("United Fire")
        headers_u = [h for h, _ in UNITED_EXCEL_COLS]
        ws_u.append(headers_u)
        _style_header_row(ws_u, len(headers_u))
        pf_col_u = next(i+1 for i,(h,_) in enumerate(UNITED_EXCEL_COLS) if h == "Pass / Fail")

        for row_idx, item in enumerate(united_rows, start=2):
            fd = item.get("form_data", {})
            row_data = [fd.get(fk, "") for _, fk in UNITED_EXCEL_COLS]
            ws_u.append(row_data)
            _style_data_row(ws_u, row_idx, len(UNITED_EXCEL_COLS), pf_col_u)
            ws_u.row_dimensions[row_idx].height = 18

        for col_idx, (header, _) in enumerate(UNITED_EXCEL_COLS, start=1):
            ws_u.column_dimensions[ws_u.cell(row=1, column=col_idx).column_letter].width = max(
                12, min(30, len(header) + 4)
            )
        ws_u.freeze_panes = "A2"

    if jax_rows:
        ws_j = wb.create_sheet("Jacksonville")
        headers_j = [h for h, _ in JAX_EXCEL_COLS]
        ws_j.append(headers_j)
        _style_header_row(ws_j, len(headers_j))
        pf_col_j = next(i+1 for i,(h,_) in enumerate(JAX_EXCEL_COLS) if h == "Pass / Fail")

        for row_idx, item in enumerate(jax_rows, start=2):
            fd = item.get("form_data", {})
            row_data = [fd.get(fk, "") for _, fk in JAX_EXCEL_COLS]
            ws_j.append(row_data)
            _style_data_row(ws_j, row_idx, len(JAX_EXCEL_COLS), pf_col_j)
            ws_j.row_dimensions[row_idx].height = 18

        for col_idx, (header, _) in enumerate(JAX_EXCEL_COLS, start=1):
            ws_j.column_dimensions[ws_j.cell(row=1, column=col_idx).column_letter].width = max(
                12, min(30, len(header) + 4)
            )
        ws_j.freeze_panes = "A2"

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ===========================================================================
# PDF generation — United Fire
# ===========================================================================

def generate_united_pdf(form_data: dict) -> bytes:
    reader = PdfReader(TEMPLATE_UNITED)
    writer = PdfWriter()
    writer.append(reader)

    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))

    for field, (x, y, sz) in UNITED_TEXT_FIELDS.items():
        put_text(c, form_data.get(field, ""), x, y, sz)

    atype = form_data.get("assembly_type", "")
    if atype == "RP":   draw_x(c, *UNITED_CHECKBOXES["RP"])
    elif atype == "DC": draw_x(c, *UNITED_CHECKBOXES["DC"])
    elif atype == "PVB":draw_x(c, *UNITED_CHECKBOXES["PVB"])
    elif atype == "SVB":draw_x(c, *UNITED_CHECKBOXES["SVB"])

    svc = form_data.get("system_service", "")
    if svc == "Fire":        draw_x(c, *UNITED_CHECKBOXES["FIRE"])
    elif svc == "Domestic":  draw_x(c, *UNITED_CHECKBOXES["DOMESTIC"])
    elif svc == "Irrigation":draw_x(c, *UNITED_CHECKBOXES["IRRIGATION"])
    elif svc == "Attraction":draw_x(c, *UNITED_CHECKBOXES["ATTRACTION"])

    bypass = form_data.get("bypass", "")
    if bypass == "Yes": draw_x(c, *UNITED_CHECKBOXES["BYPASS_YES"])
    elif bypass == "No":draw_x(c, *UNITED_CHECKBOXES["BYPASS_NO"])

    cv1r = form_data.get("cv1_result", "")
    if cv1r == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["CV1_CLOSED"])
    elif cv1r == "Leaked":     draw_x(c, *UNITED_CHECKBOXES["CV1_LEAKED"])

    cv2r = form_data.get("cv2_result", "")
    if cv2r == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["CV2_CLOSED"])
    elif cv2r == "Leaked":     draw_x(c, *UNITED_CHECKBOXES["CV2_LEAKED"])

    rvr = form_data.get("rv_result", "")
    if rvr == "Opened":        draw_x(c, *UNITED_CHECKBOXES["RV_OPENED"])
    elif rvr == "Did Not Open":draw_x(c, *UNITED_CHECKBOXES["RV_DIDNOTOPEN"])

    rvo = form_data.get("rv_out_result", "")
    if rvo == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["RV_OUT_CLOSED"])
    elif rvo == "Leaked":     draw_x(c, *UNITED_CHECKBOXES["RV_OUT_LEAKED"])

    rvi = form_data.get("rv_in_result", "")
    if rvi == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["RV_IN_CLOSED"])
    elif rvi == "Leaked":     draw_x(c, *UNITED_CHECKBOXES["RV_IN_LEAKED"])

    pai = form_data.get("pvb_ai_result", "")
    if pai == "Opened":        draw_x(c, *UNITED_CHECKBOXES["PVB_AI_OPENED"])
    elif pai == "Did Not Open":draw_x(c, *UNITED_CHECKBOXES["PVB_AI_CLOSED"])

    pcv = form_data.get("pvb_cv_result", "")
    if pcv == "Leaked":    draw_x(c, *UNITED_CHECKBOXES["PVB_CV_LEAKED"])
    elif pcv == "Held":    draw_x(c, *UNITED_CHECKBOXES["PVB_CV_HELD"])

    ar = form_data.get("assembly_result", "")
    if ar == "PASSED":  draw_x(c, *UNITED_CHECKBOXES["PASSED"])
    elif ar == "FAILED":draw_x(c, *UNITED_CHECKBOXES["FAILED"])

    repair = form_data.get("repair_desc", "")
    if repair:
        bx, by, bh, bl, bw = UNITED_REPAIR_BOX
        lines = wrap_text(repair, bw)
        c.setFont("Helvetica-Bold", 7)
        c.setFillColorRGB(1, 0, 0)
        for i, ln in enumerate(lines[:bl]):
            c.drawString(bx, by - i * bh, ln)

    sig_reader = get_signature_image_reader()
    if sig_reader:
        c.drawImage(sig_reader, UNITED_SIG_X, UNITED_SIG_Y,
                    width=UNITED_SIG_W, height=UNITED_SIG_H, mask="auto")

    c.save()
    packet.seek(0)

    overlay_reader = PdfReader(packet)
    page = writer.pages[0]
    page.merge_page(overlay_reader.pages[0])

    out = BytesIO()
    writer.write(out)
    return out.getvalue()


# ===========================================================================
# PDF generation — Jacksonville
# ===========================================================================

def generate_jax_pdf(form_data: dict) -> bytes:
    reader = PdfReader(TEMPLATE_JAX)
    writer = PdfWriter()
    writer.append(reader)

    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(JAX_PAGE_W, JAX_PAGE_H))

    for field, (x, y, sz) in JAX_TEXT_FIELDS.items():
        put_text(c, form_data.get(field, ""), x, y, sz)

    comm_tp = form_data.get("comm_test_purpose", "")
    if comm_tp == "Annual":          draw_x(c, *JAX_CHECKBOXES["COMM_ANNUAL"])
    elif comm_tp == "Repair":        draw_x(c, *JAX_CHECKBOXES["COMM_REPAIR"])
    elif comm_tp == "Replacement":   draw_x(c, *JAX_CHECKBOXES["COMM_REPLACEMENT"])
    elif comm_tp == "New Install":   draw_x(c, *JAX_CHECKBOXES["COMM_NEW_INSTALL"])

    comm_st = form_data.get("comm_service_type", "")
    if comm_st == "Fire":        draw_x(c, *JAX_CHECKBOXES["COMM_FIRE"])
    elif comm_st == "Irrigation":draw_x(c, *JAX_CHECKBOXES["COMM_IRRIGATION"])
    elif comm_st == "Process":   draw_x(c, *JAX_CHECKBOXES["COMM_PROCESS"])
    elif comm_st == "Potable":   draw_x(c, *JAX_CHECKBOXES["COMM_POTABLE"])

    comm_fp = form_data.get("comm_fire_bypass", "")
    if comm_fp: draw_x(c, *JAX_CHECKBOXES["COMM_FIRE_BYPASS"])

    comm_rc = form_data.get("comm_reclaim", "")
    if comm_rc == "Yes":  draw_x(c, *JAX_CHECKBOXES["RECLAIM_YES"])
    elif comm_rc == "No": draw_x(c, *JAX_CHECKBOXES["RECLAIM_NO"])

    res_tp = form_data.get("res_test_purpose", "")
    if res_tp == "Annual":          draw_x(c, *JAX_CHECKBOXES["RES_ANNUAL"])
    elif res_tp == "Repair":        draw_x(c, *JAX_CHECKBOXES["RES_REPAIR"])
    elif res_tp == "Replacement":   draw_x(c, *JAX_CHECKBOXES["RES_REPLACEMENT"])
    elif res_tp == "New Install":   draw_x(c, *JAX_CHECKBOXES["RES_NEW_INSTALL"])

    res_st = form_data.get("res_service_type", "")
    if res_st == "Potable":    draw_x(c, *JAX_CHECKBOXES["RES_POTABLE"])
    elif res_st == "Irrigation":draw_x(c, *JAX_CHECKBOXES["RES_IRRIGATION"])

    res_rc = form_data.get("res_reclaim", "")
    if res_rc == "Yes":  draw_x(c, *JAX_CHECKBOXES["RES_RECLAIM_YES"])
    elif res_rc == "No": draw_x(c, *JAX_CHECKBOXES["RES_RECLAIM_NO"])

    icv1 = form_data.get("init_cv1_result", "")
    if icv1 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["INIT_CV1_CLOSED"])

    icv2 = form_data.get("init_cv2_result", "")
    if icv2 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["INIT_CV2_CLOSED"])

    irv = form_data.get("init_rv_result", "")
    if irv == "Opened":         draw_x(c, *JAX_CHECKBOXES["INIT_RV_OPENED"])
    elif irv == "Did Not Open": draw_x(c, *JAX_CHECKBOXES["INIT_RV_DIDNOT"])

    ipvb = form_data.get("init_pvb_result", "")
    if ipvb == "Air Inlet Opened":    draw_x(c, *JAX_CHECKBOXES["INIT_PVB_AIOPEN"])
    elif ipvb == "Air Inlet Did Not": draw_x(c, *JAX_CHECKBOXES["INIT_PVB_AIDNOT"])

    fcv1 = form_data.get("final_cv1_result", "")
    if fcv1 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["FINAL_CV1_CLOSED"])

    fcv2 = form_data.get("final_cv2_result", "")
    if fcv2 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["FINAL_CV2_CLOSED"])

    frv = form_data.get("final_rv_result", "")
    if frv == "Opened": draw_x(c, *JAX_CHECKBOXES["FINAL_RV_OPENED"])

    fpvb = form_data.get("final_pvb_result", "")
    if fpvb == "Satisfactory": draw_x(c, *JAX_CHECKBOXES["FINAL_PVB_SAT"])

    ar = form_data.get("assembly_result", "")
    if ar == "PASSED":   draw_x(c, *JAX_CHECKBOXES["JAX_PASSED"])
    elif ar == "FAILED": draw_x(c, *JAX_CHECKBOXES["JAX_FAILED"])

    sig_reader = get_signature_image_reader()
    if sig_reader:
        c.drawImage(sig_reader, JAX_SIG_X, JAX_SIG_Y,
                    width=JAX_SIG_W, height=JAX_SIG_H, mask="auto")

    c.save()
    packet.seek(0)

    overlay_reader = PdfReader(packet)
    page = writer.pages[0]
    page.merge_page(overlay_reader.pages[0])

    out = BytesIO()
    writer.write(out)
    return out.getvalue()


# ===========================================================================
# Session-state helpers
# ===========================================================================

TESTER_KEYS = ["gauge_mfg", "gauge_serial", "date_cal", "technician", "cert_no", "recert"]

JAX_TESTER_MAP = {
    "init_tester_name":  "technician",
    "init_company":      "company",
    "init_cert":         "cert_no",
    "final_tester_name": "technician",
    "final_company":     "company",
    "final_cert":        "cert_no",
}


def apply_profile_to_forms(profile: dict):
    """Push a technician profile into both active forms + signature.

    Key design principle: we write directly to the form dicts so the values
    survive the next Streamlit render cycle.  We do NOT clear widget keys here
    because clearable_input/tap_clear_input now read value= from the form dict
    every render — clearing the keys would cause them to re-initialise to ""
    and immediately overwrite what we just set.
    """
    if not profile:
        return

    # Load signature
    if profile.get("signature_b64"):
        st.session_state["signature_b64"] = profile["signature_b64"]

    # United Fire form
    _init_form("united_form")
    united = st.session_state["united_form"]
    for tk in TESTER_KEYS:
        united[tk] = profile.get(tk, "")

    # Jacksonville form
    _init_form("jax_form")
    jax = st.session_state["jax_form"]
    for jk, pk in JAX_TESTER_MAP.items():
        jax[jk] = profile.get(pk, "")
    # also push gauge/recert fields into jax form for completeness
    for tk in TESTER_KEYS:
        jax[tk] = profile.get(tk, "")


def _init_form(key, defaults=None):
    if key not in st.session_state:
        st.session_state[key] = defaults or {}


# ===========================================================================
# Sidebar — Technician profile
# ===========================================================================

def render_technician_sidebar():
    st.sidebar.title("\U0001f464 Technician Profile")

    names = get_technician_names()
    prev_sel = st.session_state.get("_sidebar_tech_sel", "")
    if prev_sel not in names:
        prev_sel = ""

    selected = st.sidebar.selectbox(
        "Select technician",
        names,
        index=names.index(prev_sel) if prev_sel in names else 0,
        key="sidebar_tech_select",
    )

    st.session_state["_sidebar_tech_sel"] = selected

    # Auto-load profile on first visit OR when selection changes
    last_loaded = st.session_state.get("_last_loaded_tech", None)
    if selected and selected != last_loaded:
        profile = get_technician_profile(selected)
        apply_profile_to_forms(profile)
        st.session_state["_last_loaded_tech"] = selected

    # Manual reload button (useful after editing a profile)
    if st.sidebar.button(
        "\U0001f4e5 Reload Profile",
        key="load_profile_btn",
        disabled=not bool(selected),
    ):
        profile = get_technician_profile(selected)
        apply_profile_to_forms(profile)
        st.session_state["_last_loaded_tech"] = selected
        st.sidebar.success(f"Loaded: {selected}")
        st.rerun()

    st.sidebar.divider()

    # ── Edit / Add Profile ──────────────────────────────────────────────────
    with st.sidebar.expander("\u270f\ufe0f Edit / Add Profile", expanded=False):
        current = get_technician_profile(selected) if selected else {}
        prof_name = st.text_input("Name (key)", value=selected, key="prof_name")
        prof_co   = st.text_input("Company",    value=current.get("company", ""), key="prof_co")
        prof_cert = st.text_input("Cert No.",   value=current.get("cert_no", ""), key="prof_cert")
        prof_rec  = st.text_input("Re-Cert",    value=current.get("recert", ""),  key="prof_rec")
        prof_gmfg = st.text_input("Gauge Mfg",  value=current.get("gauge_mfg", ""), key="prof_gmfg")
        prof_gsn  = st.text_input("Gauge SN",   value=current.get("gauge_serial", ""), key="prof_gsn")
        prof_cal  = st.text_input("Date Cal",   value=current.get("date_cal", ""), key="prof_cal")

        if current.get("signature_b64"):
            st.caption("Saved profile signature:")
            st.image(base64.b64decode(current["signature_b64"]), width=180)

        if st.button("\U0001f4be Save Profile", key="save_profile_btn"):
            if prof_name.strip():
                new_profile = {
                    "technician":    prof_name.strip(),
                    "company":       prof_co.strip(),
                    "cert_no":       prof_cert.strip(),
                    "recert":        prof_rec.strip(),
                    "gauge_mfg":     prof_gmfg.strip(),
                    "gauge_serial":  prof_gsn.strip(),
                    "date_cal":      prof_cal.strip(),
                    "signature_b64": st.session_state.get("signature_b64", ""),
                }
                ok, msg = upsert_technician_profile(prof_name.strip(), new_profile)
                if ok:
                    st.success(msg)
                else:
                    st.warning(msg)
                st.session_state["_sidebar_tech_sel"] = prof_name.strip()
                st.session_state["_last_loaded_tech"] = None  # force reload on next render
                st.rerun()
            else:
                st.error("Name cannot be empty.")

    st.sidebar.divider()

    # ── Signature section ───────────────────────────────────────────────────
    st.sidebar.markdown("**Signature**")

    sig_b64 = st.session_state.get("signature_b64")
    if sig_b64:
        st.sidebar.image(
            base64.b64decode(sig_b64),
            caption="Current signature",
            use_container_width=True,
        )

    upload = st.sidebar.file_uploader(
        "Upload signature (PNG preferred)",
        type=["png", "jpg", "jpeg"],
        key="sig_upload",
    )
    if upload is not None:
        try:
            img = Image.open(upload).convert("RGBA")
            buf = BytesIO()
            img.save(buf, format="PNG")
            st.session_state["signature_b64"] = base64.b64encode(buf.getvalue()).decode()
            st.sidebar.success("Uploaded signature loaded.")
        except Exception as e:
            st.sidebar.error(f"Could not read uploaded signature: {e}")

    st.sidebar.caption("Or draw your signature below:")
    try:
        from streamlit_drawable_canvas import st_canvas
        canvas_result = st_canvas(
            fill_color="rgba(255,255,255,0)",
            stroke_width=2,
            stroke_color="#000000",
            background_color="#FFFFFF",
            height=80,
            width=220,
            drawing_mode="freedraw",
            key="sb_sig_canvas",
        )
        if canvas_result.image_data is not None:
            arr = canvas_result.image_data
            if arr.max() > 0 and arr[:, :, 3].max() > 0:
                save_signature(arr)
    except ImportError:
        st.sidebar.info("Install streamlit-drawable-canvas for signature support.")

    col_sig1, col_sig2 = st.sidebar.columns(2)
    with col_sig1:
        if st.button("\U0001f5d1\ufe0f Clear Sig", key="sb_clear_sig"):
            clear_signature()
            st.rerun()
    with col_sig2:
        if selected and st.button("\U0001f4be Save Sig to Profile", key="save_sig_to_profile"):
            current = get_technician_profile(selected)
            current["signature_b64"] = st.session_state.get("signature_b64", "")
            ok, msg = upsert_technician_profile(selected, current)
            if ok:
                st.sidebar.success("Signature saved to profile.")
            else:
                st.sidebar.warning(msg)


# ===========================================================================
# Tester info panel — rendered at the bottom of each form tab
# ===========================================================================

def render_tester_panel_united():
    """Read-only display + editable fields showing the active technician data."""
    form = st.session_state["united_form"]
    selected_tech = st.session_state.get("_sidebar_tech_sel", "")

    st.divider()
    st.markdown("**🔧 Tester / Technician Info**")

    if selected_tech:
        st.caption(f"Profile loaded: **{selected_tech}** — edit in sidebar or update fields below.")
    else:
        st.caption("No profile selected. Select one in the sidebar or fill in manually.")

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        clearable_input("Gauge Mfg",       "united_form", "gauge_mfg",    "u_gmfg")
        clearable_input("Gauge Serial",    "united_form", "gauge_serial", "u_gsn")
        clearable_input("Date Calibrated", "united_form", "date_cal",     "u_cal")
    with col_t2:
        clearable_input("Technician",      "united_form", "technician",   "u_tech")
        clearable_input("Cert No.",        "united_form", "cert_no",      "u_cert")
        clearable_input("Re-Cert Date",    "united_form", "recert",       "u_recert")

    # Show signature preview
    sig_b64 = st.session_state.get("signature_b64")
    if sig_b64:
        st.image(base64.b64decode(sig_b64), caption="Signature on file", width=200)
    else:
        st.caption("⚠️ No signature loaded. Upload or draw one in the sidebar.")


def render_tester_panel_jax():
    """Tester panel for Jacksonville form."""
    form = st.session_state["jax_form"]
    selected_tech = st.session_state.get("_sidebar_tech_sel", "")

    st.divider()
    st.markdown("**🔧 Tester / Technician Info**")

    if selected_tech:
        st.caption(f"Profile loaded: **{selected_tech}** — edit in sidebar or update fields below.")
    else:
        st.caption("No profile selected. Select one in the sidebar or fill in manually.")

    col_itn, col_ico, col_ic = st.columns(3)
    with col_itn:
        clearable_input("Init Tester Name", "jax_form", "init_tester_name", "j_itn")
    with col_ico:
        clearable_input("Init Company",     "jax_form", "init_company",     "j_ico")
    with col_ic:
        clearable_input("Init Cert",        "jax_form", "init_cert",        "j_ic")

    col_rb, col_rco, col_rc = st.columns(3)
    with col_rb:
        clearable_input("Repaired By",    "jax_form", "repaired_by",    "j_rb")
    with col_rco:
        clearable_input("Repair Company", "jax_form", "repair_company", "j_rco")
    with col_rc:
        clearable_input("Repair Cert",    "jax_form", "repair_cert",    "j_rc")

    col_ftn, col_fco, col_fc = st.columns(3)
    with col_ftn:
        clearable_input("Final Tester Name", "jax_form", "final_tester_name", "j_ftn")
    with col_fco:
        clearable_input("Final Company",     "jax_form", "final_company",     "j_fco")
    with col_fc:
        clearable_input("Final Cert",        "jax_form", "final_cert",        "j_fc")

    # Show signature preview
    sig_b64 = st.session_state.get("signature_b64")
    if sig_b64:
        st.image(base64.b64decode(sig_b64), caption="Signature on file", width=200)
    else:
        st.caption("⚠️ No signature loaded. Upload or draw one in the sidebar.")


# ===========================================================================
# United Fire tab
# ===========================================================================

def render_united_tab():
    _init_form("united_form")
    form = st.session_state["united_form"]

    st.subheader("\U0001f4cb United Fire \u2014 Backflow Test Report")

    col1, col2 = st.columns([1.15, 1])
    with col1:
        synced_date_input("Inspection Date", "united_form", "date", "u_date_picker", ["date", "test_date"])
    with col2:
        tap_clear_input("Branch", "united_form", "branch", "u_branch")

    tap_clear_input("AHJ", "united_form", "ahj", "u_ahj")

    col4, col5 = st.columns(2)
    with col4:
        tap_clear_input("Customer Name", "united_form", "customer_name", "u_cust")
    with col5:
        tap_clear_input("Street Address", "united_form", "street_address", "u_addr")

    clearable_input("Location / Description", "united_form", "location", "u_loc")

    st.divider()
    st.markdown("**Assembly Info**")
    col6, col7 = st.columns(2)
    with col6:
        tap_clear_input("Serial Number", "united_form", "serial_number", "u_sn")
        tap_clear_input("Manufacturer", "united_form", "manufacturer", "u_mfg")
    with col7:
        tap_clear_input("Model", "united_form", "model", "u_mdl")
        tap_clear_input("Size", "united_form", "size", "u_sz")

    col10, col11, col12 = st.columns(3)
    with col10:
        atype_opts = ["", "RP", "DC", "PVB", "SVB"]
        cur_at = form.get("assembly_type", "")
        at_idx = atype_opts.index(cur_at) if cur_at in atype_opts else 0
        form["assembly_type"] = st.selectbox("Assembly Type", atype_opts, index=at_idx, key="u_atype")
    with col11:
        svc_opts = ["", "Fire", "Domestic", "Irrigation", "Attraction"]
        cur_sv = form.get("system_service", "")
        sv_idx = svc_opts.index(cur_sv) if cur_sv in svc_opts else 0
        form["system_service"] = st.selectbox("System Service", svc_opts, index=sv_idx, key="u_ss")
    with col12:
        bp_opts = ["", "Yes", "No"]
        cur_bp = form.get("bypass", "")
        bp_idx = bp_opts.index(cur_bp) if cur_bp in bp_opts else 0
        form["bypass"] = st.selectbox("Fire Bypass", bp_opts, index=bp_idx, key="u_bp")

    st.divider()
    st.markdown("**Test Results**")

    atype = form.get("assembly_type", "")

    if atype in ("", "RP", "DC"):
        col_cv1a, col_cv1b = st.columns(2)
        with col_cv1a:
            cv1_opts = ["", "Closed Tight", "Leaked"]
            cur_cv1 = form.get("cv1_result", "")
            cv1_idx = cv1_opts.index(cur_cv1) if cur_cv1 in cv1_opts else 0
            form["cv1_result"] = st.selectbox("CV1 Result", cv1_opts, index=cv1_idx, key="cv1_result")
        with col_cv1b:
            clearable_input("CV1 DP (psi)", "united_form", "cv1_dp", "u_cv1dp")

        col_cv2a, col_cv2b = st.columns(2)
        with col_cv2a:
            cv2_opts = ["", "Closed Tight", "Leaked"]
            cur_cv2 = form.get("cv2_result", "")
            cv2_idx = cv2_opts.index(cur_cv2) if cur_cv2 in cv2_opts else 0
            form["cv2_result"] = st.selectbox("CV2 Result", cv2_opts, index=cv2_idx, key="cv2_result")
        with col_cv2b:
            clearable_input("CV2 DP (psi)", "united_form", "cv2_dp", "u_cv2dp")

    if atype in ("", "RP"):
        col_rva, col_rvb = st.columns(2)
        with col_rva:
            rv_opts = ["", "Opened", "Did Not Open"]
            cur_rv = form.get("rv_result", "")
            rv_idx = rv_opts.index(cur_rv) if cur_rv in rv_opts else 0
            form["rv_result"] = st.selectbox("RV Result", rv_opts, index=rv_idx, key="rv_result")
        with col_rvb:
            clearable_input("RV Opened At (psi)", "united_form", "rv_psi", "u_rvpsi")

        col_rvoa, col_rvob = st.columns(2)
        with col_rvoa:
            rvo_opts = ["", "Closed Tight", "Leaked"]
            cur_rvo = form.get("rv_out_result", "")
            rvo_idx = rvo_opts.index(cur_rvo) if cur_rvo in rvo_opts else 0
            form["rv_out_result"] = st.selectbox("RV Outlet", rvo_opts, index=rvo_idx, key="rv_out_result")
        with col_rvob:
            rvi_opts = ["", "Closed Tight", "Leaked"]
            cur_rvi = form.get("rv_in_result", "")
            rvi_idx = rvi_opts.index(cur_rvi) if cur_rvi in rvi_opts else 0
            form["rv_in_result"] = st.selectbox("RV Inlet", rvi_opts, index=rvi_idx, key="rv_in_result")

    if atype in ("", "PVB", "SVB"):
        col_aia, col_aib = st.columns(2)
        with col_aia:
            ai_opts = ["", "Opened", "Did Not Open"]
            cur_ai = form.get("pvb_ai_result", "")
            ai_idx = ai_opts.index(cur_ai) if cur_ai in ai_opts else 0
            form["pvb_ai_result"] = st.selectbox("PVB Air Inlet", ai_opts, index=ai_idx, key="pvb_ai_result")
        with col_aib:
            clearable_input("PVB AI (psi)", "united_form", "pvb_ai_psi", "u_aipsi")

        col_cva, col_cvb = st.columns(2)
        with col_cva:
            pcv_opts = ["", "Leaked", "Held"]
            cur_pcv = form.get("pvb_cv_result", "")
            pcv_idx = pcv_opts.index(cur_pcv) if cur_pcv in pcv_opts else 0
            form["pvb_cv_result"] = st.selectbox("PVB CV Result", pcv_opts, index=pcv_idx, key="pvb_cv_result")
        with col_cvb:
            clearable_input("PVB CV (psi)", "united_form", "pvb_cv_psi", "u_cvpsi")

    st.divider()
    st.caption(f"Test Date will match Inspection Date: {form.get('test_date', '')}")

    ar_opts = ["", "PASSED", "FAILED"]
    cur_ar = form.get("assembly_result", "")
    ar_idx = ar_opts.index(cur_ar) if cur_ar in ar_opts else 0
    form["assembly_result"] = st.selectbox("Assembly Result", ar_opts, index=ar_idx, key="assembly_result")

    clearable_input("Repair Description", "united_form", "repair_desc", "u_rep")

    # Tester panel at the bottom
    render_tester_panel_united()

    st.divider()

    col_gen, col_save = st.columns(2)
    with col_gen:
        if st.button("\U0001f5a8\ufe0f Generate PDF", key="u_gen_pdf"):
            try:
                pdf_bytes = generate_united_pdf(form)
                fname = f"backflow_{form.get('customer_name','report').replace(' ','_')}.pdf"
                st.download_button("\U0001f4e5 Download PDF", pdf_bytes, fname, "application/pdf", key="u_dl_pdf")
                for k in UNITED_GREEN - {"assembly_type", "system_service", "bypass"}:
                    form.pop(k, None)
                for wk in UNITED_GREEN_WIDGET_KEYS:
                    st.session_state.pop(wk, None)
            except Exception as e:
                st.error(f"PDF error: {e}")

    with col_save:
        if st.button("\U0001f4be Save to Job Folder", key="u_save_job"):
            if "job_folder" not in st.session_state:
                st.session_state["job_folder"] = []
            st.session_state["job_folder"].append({
                "form_type": "united",
                "form_data": dict(form),
                "label": form.get("customer_name") or form.get("street_address") or "Report",
            })
            st.success(f"Saved! Job folder has {len(st.session_state['job_folder'])} report(s).")

    st.divider()
    col_nr, col_nj = st.columns(2)
    with col_nr:
        if st.button("\U0001f504 Next Report (same job)", key="u_next_report"):
            keep = {k: v for k, v in form.items() if k in UNITED_NEXT_REPORT_KEEP}
            st.session_state["united_form"] = keep
            for wk in UNITED_GREEN_WIDGET_KEYS:
                st.session_state.pop(wk, None)
            st.rerun()
    with col_nj:
        if st.button("\U0001f195 New Job", key="u_new_job"):
            keep = {k: v for k, v in form.items() if k in UNITED_NEW_JOB_KEEP}
            st.session_state["united_form"] = keep
            for wk in UNITED_GREEN_WIDGET_KEYS + UNITED_BLUE_WIDGET_KEYS:
                st.session_state.pop(wk, None)
            st.rerun()


# ===========================================================================
# Jacksonville tab
# ===========================================================================

def render_jax_tab():
    _init_form("jax_form")
    form = st.session_state["jax_form"]

    st.subheader("\U0001f4cb Jacksonville (JEA) \u2014 Backflow Test Report")

    synced_date_input(
        "Inspection Date",
        "jax_form",
        "signature_date",
        "j_sig_date_picker",
        ["signature_date", "init_test_date", "final_test_date", "repair_date"],
    )

    st.markdown("**Property Info**")
    col1, col2 = st.columns(2)
    with col1:
        clearable_input("Premises Name",    "jax_form", "premises_name",    "j_prem")
        clearable_input("Service Address",  "jax_form", "service_address",  "j_sa")
        clearable_input("Physical Location","jax_form", "physical_location","j_pl")
        clearable_input("JEA Account",      "jax_form", "jea_account",      "j_acct")
    with col2:
        clearable_input("Owner Name",       "jax_form", "owner_name",       "j_own")
        clearable_input("Mailing Address",  "jax_form", "mailing_address",  "j_ma")
        clearable_input("Contact Phone",    "jax_form", "contact_phone",    "j_ph")
        clearable_input("Meter Number",     "jax_form", "meter_number",     "j_meter")

    st.divider()
    st.markdown("**Test Purpose & Service Type**")

    col_ctp, col_cst, col_crc = st.columns(3)
    with col_ctp:
        ctp_opts = ["", "Annual", "Repair", "Replacement", "New Install"]
        cur_ctp = form.get("comm_test_purpose", "")
        ctp_idx = ctp_opts.index(cur_ctp) if cur_ctp in ctp_opts else 0
        form["comm_test_purpose"] = st.selectbox("Comm Test Purpose", ctp_opts, index=ctp_idx, key="j_ctp")
    with col_cst:
        cst_opts = ["", "Fire", "Irrigation", "Process", "Potable"]
        cur_cst = form.get("comm_service_type", "")
        cst_idx = cst_opts.index(cur_cst) if cur_cst in cst_opts else 0
        form["comm_service_type"] = st.selectbox("Comm Service Type", cst_opts, index=cst_idx, key="j_cst")
    with col_crc:
        crc_opts = ["", "Yes", "No"]
        cur_crc = form.get("comm_reclaim", "")
        crc_idx = crc_opts.index(cur_crc) if cur_crc in crc_opts else 0
        form["comm_reclaim"] = st.selectbox("Comm Reclaim Water", crc_opts, index=crc_idx, key="j_crc")

    col_rtp, col_rst, col_rrc = st.columns(3)
    with col_rtp:
        rtp_opts = ["", "Annual", "Repair", "Replacement", "New Install"]
        cur_rtp = form.get("res_test_purpose", "")
        rtp_idx = rtp_opts.index(cur_rtp) if cur_rtp in rtp_opts else 0
        form["res_test_purpose"] = st.selectbox("Res Test Purpose", rtp_opts, index=rtp_idx, key="j_rtp")
    with col_rst:
        rst_opts = ["", "Potable", "Irrigation"]
        cur_rst = form.get("res_service_type", "")
        rst_idx = rst_opts.index(cur_rst) if cur_rst in rst_opts else 0
        form["res_service_type"] = st.selectbox("Res Service Type", rst_opts, index=rst_idx, key="j_rst")
    with col_rrc:
        rrc_opts = ["", "Yes", "No"]
        cur_rrc = form.get("res_reclaim", "")
        rrc_idx = rrc_opts.index(cur_rrc) if cur_rrc in rrc_opts else 0
        form["res_reclaim"] = st.selectbox("Res Reclaim Water", rrc_opts, index=rrc_idx, key="j_rrc")

    st.divider()
    st.markdown("**Assembly Info**")
    col_dt, col_mfg, col_sz, col_mn = st.columns(4)
    with col_dt:
        clearable_input("Device Type",  "jax_form", "device_type",  "j_dt")
    with col_mfg:
        clearable_input("Manufacturer", "jax_form", "manufacturer", "j_mfg")
    with col_sz:
        clearable_input("Size",         "jax_form", "size",         "j_sz")
    with col_mn:
        clearable_input("Model Number", "jax_form", "model_number", "j_mn")

    col_sn, col_id = st.columns(2)
    with col_sn:
        clearable_input("Serial Number", "jax_form", "serial_number", "j_sn")
    with col_id:
        clearable_input("Install Date",  "jax_form", "install_date",  "j_id")

    st.divider()
    st.markdown("**Initial Test**")
    col_icv1r, col_icv1p = st.columns(2)
    with col_icv1r:
        icv1_opts = ["", "Closed Tight", "Leaked"]
        cur_icv1 = form.get("init_cv1_result", "")
        icv1_idx = icv1_opts.index(cur_icv1) if cur_icv1 in icv1_opts else 0
        form["init_cv1_result"] = st.selectbox("Init CV1", icv1_opts, index=icv1_idx, key="init_cv1_result")
    with col_icv1p:
        clearable_input("Init CV1 (psi)", "jax_form", "init_cv1_psi", "j_icv1p")

    col_icv2r, col_icv2p = st.columns(2)
    with col_icv2r:
        icv2_opts = ["", "Closed Tight", "Leaked"]
        cur_icv2 = form.get("init_cv2_result", "")
        icv2_idx = icv2_opts.index(cur_icv2) if cur_icv2 in icv2_opts else 0
        form["init_cv2_result"] = st.selectbox("Init CV2", icv2_opts, index=icv2_idx, key="init_cv2_result")
    with col_icv2p:
        clearable_input("Init CV2 (psi)", "jax_form", "init_cv2_psi", "j_icv2p")

    col_irvr, col_irvp = st.columns(2)
    with col_irvr:
        irv_opts = ["", "Opened", "Did Not Open"]
        cur_irv = form.get("init_rv_result", "")
        irv_idx = irv_opts.index(cur_irv) if cur_irv in irv_opts else 0
        form["init_rv_result"] = st.selectbox("Init RV", irv_opts, index=irv_idx, key="init_rv_result")
    with col_irvp:
        clearable_input("Init RV (psi)", "jax_form", "init_rv_psi", "j_irvp")

    col_ipvbr, col_ipvbp = st.columns(2)
    with col_ipvbr:
        ipvb_opts = ["", "Air Inlet Opened", "Air Inlet Did Not"]
        cur_ipvb = form.get("init_pvb_result", "")
        ipvb_idx = ipvb_opts.index(cur_ipvb) if cur_ipvb in ipvb_opts else 0
        form["init_pvb_result"] = st.selectbox("Init PVB", ipvb_opts, index=ipvb_idx, key="init_pvb_result")
    with col_ipvbp:
        clearable_input("Init PVB (psi)", "jax_form", "init_pvb_psi", "j_ipvbp")

    st.divider()
    st.markdown("**Final Test**")
    col_fcv1r, col_fcv1p = st.columns(2)
    with col_fcv1r:
        fcv1_opts = ["", "Closed Tight", "Leaked"]
        cur_fcv1 = form.get("final_cv1_result", "")
        fcv1_idx = fcv1_opts.index(cur_fcv1) if cur_fcv1 in fcv1_opts else 0
        form["final_cv1_result"] = st.selectbox("Final CV1", fcv1_opts, index=fcv1_idx, key="final_cv1_result")
    with col_fcv1p:
        clearable_input("Final CV1 (psi)", "jax_form", "final_cv1_psi", "j_fcv1p")

    col_fcv2r, col_fcv2p = st.columns(2)
    with col_fcv2r:
        fcv2_opts = ["", "Closed Tight", "Leaked"]
        cur_fcv2 = form.get("final_cv2_result", "")
        fcv2_idx = fcv2_opts.index(cur_fcv2) if cur_fcv2 in fcv2_opts else 0
        form["final_cv2_result"] = st.selectbox("Final CV2", fcv2_opts, index=fcv2_idx, key="final_cv2_result")
    with col_fcv2p:
        clearable_input("Final CV2 (psi)", "jax_form", "final_cv2_psi", "j_fcv2p")

    col_frvr, col_frvp = st.columns(2)
    with col_frvr:
        frv_opts = ["", "Opened", "Did Not Open"]
        cur_frv = form.get("final_rv_result", "")
        frv_idx = frv_opts.index(cur_frv) if cur_frv in frv_opts else 0
        form["final_rv_result"] = st.selectbox("Final RV", frv_opts, index=frv_idx, key="final_rv_result")
    with col_frvp:
        clearable_input("Final RV (psi)", "jax_form", "final_rv_psi", "j_frvp")

    col_fpvbr, _ = st.columns(2)
    with col_fpvbr:
        fpvb_opts = ["", "Satisfactory", "Unsatisfactory"]
        cur_fpvb = form.get("final_pvb_result", "")
        fpvb_idx = fpvb_opts.index(cur_fpvb) if cur_fpvb in fpvb_opts else 0
        form["final_pvb_result"] = st.selectbox("Final PVB", fpvb_opts, index=fpvb_idx, key="final_pvb_result")

    st.divider()
    ares_opts = ["", "PASSED", "FAILED"]
    cur_ares = form.get("assembly_result", "")
    ares_idx = ares_opts.index(cur_ares) if cur_ares in ares_opts else 0
    form["assembly_result"] = st.selectbox("Assembly Result", ares_opts, index=ares_idx, key="j_ares")

    clearable_input("Repairs / Notes", "jax_form", "repairs", "j_rep")

    # Tester panel at the bottom
    render_tester_panel_jax()

    st.divider()
    col_jgen, col_jsave = st.columns(2)
    with col_jgen:
        if st.button("\U0001f5a8\ufe0f Generate PDF", key="j_gen_pdf"):
            try:
                pdf_bytes = generate_jax_pdf(form)
                fname = f"jax_{form.get('premises_name','report').replace(' ','_')}.pdf"
                st.download_button("\U0001f4e5 Download PDF", pdf_bytes, fname, "application/pdf", key="j_dl_pdf")
            except Exception as e:
                st.error(f"PDF error: {e}")

    with col_jsave:
        if st.button("\U0001f4be Save to Job Folder", key="j_save_job"):
            if "job_folder" not in st.session_state:
                st.session_state["job_folder"] = []
            st.session_state["job_folder"].append({
                "form_type": "jax",
                "form_data": dict(form),
                "label": form.get("premises_name") or form.get("service_address") or "JAX Report",
            })
            st.success(f"Saved! Job folder has {len(st.session_state['job_folder'])} report(s).")

    st.divider()
    col_jnr, col_jnj = st.columns(2)
    with col_jnr:
        if st.button("\U0001f504 Next Report (same job)", key="j_next_report"):
            keep = {k: v for k, v in form.items() if k in JAX_NEXT_REPORT_KEEP}
            st.session_state["jax_form"] = keep
            for wk in JAX_GREEN_WIDGET_KEYS:
                st.session_state.pop(wk, None)
            st.rerun()
    with col_jnj:
        if st.button("\U0001f195 New Job", key="j_new_job"):
            keep = {k: v for k, v in form.items() if k in JAX_NEW_JOB_KEEP}
            st.session_state["jax_form"] = keep
            for wk in JAX_GREEN_WIDGET_KEYS + JAX_BLUE_WIDGET_KEYS:
                st.session_state.pop(wk, None)
            st.rerun()


# ===========================================================================
# Job Folder tab
# ===========================================================================

def render_job_folder_tab():
    st.subheader("\U0001f4c2 Job Folder")

    folder = st.session_state.get("job_folder", [])
    if not folder:
        st.info("No reports saved yet. Use \"Save to Job Folder\" on either form tab.")
        return

    st.write(f"**{len(folder)} report(s) in folder:**")
    for i, item in enumerate(folder):
        label = item.get("label", f"Report {i+1}")
        ft = item.get("form_type", "united")
        tag = "\U0001f534 JAX" if ft == "jax" else "\U0001f535 United"
        st.write(f"{i+1}. {tag} — {label}")

    st.divider()
    col_xl, col_cl = st.columns(2)
    with col_xl:
        if st.button("\U0001f4ca Export to Excel", key="export_excel"):
            try:
                buf = export_jobs_to_excel(folder)
                st.download_button(
                    "\U0001f4e5 Download Excel",
                    buf,
                    "job_folder.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_excel",
                )
            except Exception as e:
                st.error(f"Excel export error: {e}")
    with col_cl:
        if st.button("\U0001f5d1\ufe0f Clear Folder", key="clear_folder"):
            st.session_state["job_folder"] = []
            st.rerun()


# ===========================================================================
# Main app
# ===========================================================================

def main():
    st.set_page_config(
        page_title="Backflow Test Reports",
        page_icon="\U0001f4cb",
        layout="wide",
    )

    render_technician_sidebar()

    tab_united, tab_jax, tab_folder = st.tabs([
        "\U0001f535 United Fire",
        "\U0001f534 Jacksonville (JEA)",
        "\U0001f4c2 Job Folder",
    ])

    with tab_united:
        render_united_tab()

    with tab_jax:
        render_jax_tab()

    with tab_folder:
        render_job_folder_tab()


if __name__ == "__main__":
    main()
