import streamlit as st
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import json, os, base64, requests, zipfile
from datetime import date, datetime
from pypdf import PdfReader, PdfWriter
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

TEMPLATE_UNITED = "backflow_template.pdf"
TEMPLATE_JAX = "jacksonville_template.pdf"
TECHNICIANS_FILE = "technicians.json"
JOBS_FILE = "jobs/jobs.json"
JOBS_FOLDER = "jobs"
PAGE_W, PAGE_H = 612, 792
GITHUB_REPO = "kellyjwinkle/backflow_app"
GITHUB_API_BASE = "https://api.github.com"


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


def _github_token():
    try:
        return st.secrets["GITHUB_TOKEN"]
    except Exception:
        return None


def _github_headers():
    token = _github_token()
    headers = {"Accept": "application/vnd.github+json"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    return headers


# ─────────────────────────────────────────────────────────────
# Technician helpers
# ─────────────────────────────────────────────────────────────

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
    payload = {"message": "Update technician profiles via app", "content": content_b64}
    if current_sha:
        payload["sha"] = current_sha
    try:
        r = requests.put(url, headers=_github_headers(), json=payload, timeout=10)
        if r.status_code in (200, 201):
            return True, "Profile saved to GitHub ✓"
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
    ok, msg = save_technicians_to_github(st.session_state["technicians"], st.session_state.get("technicians_sha"))
    if ok:
        _, new_sha = load_technicians_from_github()
        st.session_state["technicians_sha"] = new_sha
    return ok, msg


def delete_technician_profile(name: str):
    _init_technicians()
    if not name or name not in st.session_state["technicians"]:
        return False, "Profile not found."
    del st.session_state["technicians"][name]
    ok, msg = save_technicians_to_github(st.session_state["technicians"], st.session_state.get("technicians_sha"))
    if ok:
        _, new_sha = load_technicians_from_github()
        st.session_state["technicians_sha"] = new_sha
        return True, "Profile deleted."
    return False, msg


# ─────────────────────────────────────────────────────────────
# Jobs / autosave helpers
# ─────────────────────────────────────────────────────────────

def _load_jobs_index():
    """Return (list_of_job_dicts, sha_or_None) from jobs/jobs.json on GitHub."""
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO}/contents/{JOBS_FILE}"
    try:
        r = requests.get(url, headers=_github_headers(), timeout=8)
        if r.status_code == 200:
            data = r.json()
            content = base64.b64decode(data["content"]).decode("utf-8")
            return json.loads(content), data["sha"]
    except Exception:
        pass
    return [], None


def _save_jobs_index(jobs: list, current_sha):
    token = _github_token()
    if not token:
        return False, "No GITHUB_TOKEN — cannot save jobs index."
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO}/contents/{JOBS_FILE}"
    content_b64 = base64.b64encode(json.dumps(jobs, indent=2).encode()).decode()
    payload = {"message": "Update jobs index via app", "content": content_b64}
    if current_sha:
        payload["sha"] = current_sha
    try:
        r = requests.put(url, headers=_github_headers(), json=payload, timeout=10)
        if r.status_code in (200, 201):
            return True, r.json().get("content", {}).get("sha")
        return False, f"GitHub error {r.status_code}"
    except Exception as e:
        return False, str(e)


def _upload_pdf_to_github(filename: str, pdf_bytes: bytes):
    """Upload a PDF to jobs/<filename> on GitHub. Returns (ok, sha_or_error)."""
    token = _github_token()
    if not token:
        return False, "No GITHUB_TOKEN."
    path = f"{JOBS_FOLDER}/{filename}"
    existing_sha = None
    check_url = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO}/contents/{path}"
    try:
        cr = requests.get(check_url, headers=_github_headers(), timeout=8)
        if cr.status_code == 200:
            existing_sha = cr.json().get("sha")
    except Exception:
        pass
    content_b64 = base64.b64encode(pdf_bytes).decode()
    payload = {"message": f"Autosave job PDF: {filename}", "content": content_b64}
    if existing_sha:
        payload["sha"] = existing_sha
    try:
        r = requests.put(check_url, headers=_github_headers(), json=payload, timeout=15)
        if r.status_code in (200, 201):
            html_url = r.json().get("content", {}).get("html_url", "")
            return True, html_url
        return False, f"GitHub error {r.status_code}: {r.text[:200]}"
    except Exception as e:
        return False, str(e)


def _delete_pdf_from_github(filename: str) -> bool:
    """Delete a PDF file from jobs/<filename> on GitHub. Returns True on success."""
    token = _github_token()
    if not token:
        return False
    path = f"{JOBS_FOLDER}/{filename}"
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO}/contents/{path}"
    try:
        cr = requests.get(url, headers=_github_headers(), timeout=8)
        if cr.status_code != 200:
            return True  # already gone
        file_sha = cr.json().get("sha")
        payload = {"message": f"Clear job PDF: {filename}", "sha": file_sha}
        r = requests.delete(url, headers=_github_headers(), json=payload, timeout=10)
        return r.status_code in (200, 201, 204)
    except Exception:
        return False


def _fetch_pdf_from_github(filename: str) -> bytes | None:
    """Fetch PDF bytes from jobs/<filename> on GitHub."""
    path = f"{JOBS_FOLDER}/{filename}"
    url = f"{GITHUB_API_BASE}/repos/{GITHUB_REPO}/contents/{path}"
    try:
        r = requests.get(url, headers=_github_headers(), timeout=15)
        if r.status_code == 200:
            data = r.json()
            return base64.b64decode(data["content"])
    except Exception:
        pass
    return None


def autosave_job(form_data: dict, pdf_bytes: bytes, form_type: str):
    """
    Save PDF to jobs/ and append a summary row to jobs/jobs.json.
    Returns (ok: bool, message: str, pdf_filename: str).
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    if form_type == "united":
        cust = form_data.get("customer_name", "unknown").replace(" ", "_").replace("/", "-")
        filename = f"united_{cust}_{ts}.pdf"
    else:
        prem = form_data.get("premises_name", "unknown").replace(" ", "_").replace("/", "-")
        filename = f"jax_{prem}_{ts}.pdf"

    pdf_ok, pdf_result = _upload_pdf_to_github(filename, pdf_bytes)
    if not pdf_ok:
        return False, f"PDF upload failed: {pdf_result}", filename

    jobs, jobs_sha = _load_jobs_index()
    job_entry = {
        "filename": filename,
        "form_type": form_type,
        "saved_at": datetime.now().strftime("%m/%d/%Y %H:%M"),
        "saved_date": datetime.now().strftime("%Y-%m-%d"),
        "technician": form_data.get("technician", ""),
        "date": form_data.get("date") or form_data.get("signature_date", ""),
        "customer": form_data.get("customer_name") or form_data.get("premises_name", ""),
        "address": form_data.get("street_address") or form_data.get("service_address", ""),
        "serial_number": form_data.get("serial_number", ""),
        "assembly_type": form_data.get("assembly_type") or form_data.get("device_type", ""),
        "assembly_result": form_data.get("assembly_result", ""),
        "pdf_url": pdf_result,
    }
    jobs.append(job_entry)
    idx_ok, new_sha = _save_jobs_index(jobs, jobs_sha)
    if not idx_ok:
        return True, f"PDF saved but index update failed: {new_sha}", filename

    st.session_state["_jobs_cache"] = jobs
    return True, f"Job saved ✓ ({filename})", filename


def load_jobs_cached():
    if "_jobs_cache" not in st.session_state:
        jobs, _ = _load_jobs_index()
        st.session_state["_jobs_cache"] = jobs
    return st.session_state["_jobs_cache"]


def clear_all_reports() -> tuple[bool, str]:
    """
    Delete all PDFs from jobs/ on GitHub and reset jobs.json to [].
    Returns (ok, message).
    """
    jobs, jobs_sha = _load_jobs_index()
    errors = []
    deleted = 0
    for job in jobs:
        fname = job.get("filename", "")
        if fname:
            ok = _delete_pdf_from_github(fname)
            if ok:
                deleted += 1
            else:
                errors.append(fname)

    # Reset the index to empty list
    ok, result = _save_jobs_index([], jobs_sha)
    if not ok:
        return False, f"PDFs cleared ({deleted}) but failed to reset index: {result}"

    st.session_state["_jobs_cache"] = []
    msg = f"Cleared {deleted} PDF(s) and reset job log."
    if errors:
        msg += f" ({len(errors)} file(s) could not be deleted from GitHub.)"
    return True, msg


def build_jobs_excel(jobs: list) -> bytes:
    """Build an .xlsx summary of all saved jobs."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Job Summary"

    headers = [
        "Saved At", "Form Type", "Date", "Technician", "Customer / Premises",
        "Address", "Serial Number", "Assembly Type", "Result", "PDF Filename",
    ]
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row_idx, job in enumerate(jobs, 2):
        ws.cell(row=row_idx, column=1, value=job.get("saved_at", ""))
        ws.cell(row=row_idx, column=2, value=job.get("form_type", "").upper())
        ws.cell(row=row_idx, column=3, value=job.get("date", ""))
        ws.cell(row=row_idx, column=4, value=job.get("technician", ""))
        ws.cell(row=row_idx, column=5, value=job.get("customer", ""))
        ws.cell(row=row_idx, column=6, value=job.get("address", ""))
        ws.cell(row=row_idx, column=7, value=job.get("serial_number", ""))
        ws.cell(row=row_idx, column=8, value=job.get("assembly_type", ""))
        ws.cell(row=row_idx, column=9, value=job.get("assembly_result", ""))
        ws.cell(row=row_idx, column=10, value=job.get("filename", ""))
        result_cell = ws.cell(row=row_idx, column=9)
        if job.get("assembly_result") == "PASSED":
            result_cell.fill = PatternFill("solid", fgColor="C6EFCE")
            result_cell.font = Font(color="276221")
        elif job.get("assembly_result") == "FAILED":
            result_cell.fill = PatternFill("solid", fgColor="FFC7CE")
            result_cell.font = Font(color="9C0006")

    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_zip_with_pdfs(jobs: list) -> bytes:
    """
    Build a ZIP containing all PDFs fetched from GitHub plus an Excel summary.
    Only includes jobs from today.
    """
    today_str = datetime.now().strftime("%Y-%m-%d")
    today_jobs = [j for j in jobs if j.get("saved_date", "") == today_str]

    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Excel summary (all jobs, not just today)
        excel_bytes = build_jobs_excel(jobs)
        zf.writestr(f"backflow_jobs_{today_str}.xlsx", excel_bytes)

        # PDFs for today
        for job in today_jobs:
            fname = job.get("filename", "")
            if not fname:
                continue
            pdf_bytes = _fetch_pdf_from_github(fname)
            if pdf_bytes:
                zf.writestr(fname, pdf_bytes)

    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────
# PDF field maps
# ─────────────────────────────────────────────────────────────

UNITED_TEXT_FIELDS = {
    "date": (135, 583, 8), "branch": (235, 583, 8), "ahj": (437, 583, 8),
    "customer_name": (200, 567, 8), "street_address": (200, 551, 8), "location": (200, 533, 8),
    "serial_number": (205, 507, 8), "manufacturer": (205, 490, 8), "model": (205, 475, 8), "size": (390, 507, 8),
    "rv_psi": (300, 398, 8), "cv1_dp": (183, 320, 8), "cv2_dp": (395, 312, 8), "pvb_ai_psi": (495, 378, 8),
    "pvb_cv_psi": (495, 320, 8), "test_date": (168, 290, 8), "gauge_mfg": (215, 178, 8),
    "gauge_serial": (313, 178, 8), "date_cal": (455, 178, 8), "technician": (176, 155, 8),
    "cert_no": (407, 165, 8), "recert": (407, 150, 8),
}
UNITED_SIG_X, UNITED_SIG_Y, UNITED_SIG_W, UNITED_SIG_H = 170, 118, 130, 28
UNITED_CHECKBOXES = {
    "RP": (210,460), "DC": (270,460), "PVB": (210,441), "SVB": (270,441), "FIRE": (395,490), "DOMESTIC": (395,475),
    "IRRIGATION": (395,460), "ATTRACTION": (395,442), "BYPASS_YES": (500,470), "BYPASS_NO": (500,450),
    "CV1_CLOSED": (130,390), "CV1_LEAKED": (130,375), "CV2_CLOSED": (330,390), "CV2_LEAKED": (330,375),
    "PVB_AI_CLOSED": (426,398), "PVB_AI_OPENED": (426,378), "PVB_CV_LEAKED": (426,350), "PVB_CV_HELD": (426,323),
    "RV_OPENED": (225,398), "RV_DIDNOTOPEN": (225,378), "RV_OUT_CLOSED": (225,334), "RV_OUT_LEAKED": (272,334),
    "RV_IN_CLOSED": (225,310), "RV_IN_LEAKED": (272,310), "PASSED": (360,292), "FAILED": (415,292),
}
UNITED_REPAIR_BOX = (228, 200, 10, 3, 70)

JAX_TEXT_FIELDS = {
    "premises_name": (102, 696, 9), "owner_name": (333, 696, 9), "service_address": (104, 652, 9),
    "mailing_address": (335, 652, 9), "physical_location": (105, 614, 9), "contact_phone": (333, 612, 9),
    "jea_account": (103, 569, 9), "meter_number": (335, 571, 9), "device_type": (90, 419, 9),
    "manufacturer": (151, 421, 9), "size": (231, 421, 9), "model_number": (283, 420, 9),
    "serial_number": (358, 420, 9), "install_date": (459, 418, 9), "init_cv1_psi": (151, 338, 9),
    "init_cv2_psi": (249, 338, 9), "init_rv_psi": (396, 356, 9), "init_pvb_psi": (471, 335, 9),
    "final_cv1_psi": (165, 282, 9), "final_cv2_psi": (263, 283, 9), "final_rv_psi": (404, 299, 9),
    "repairs": (99, 244, 9), "init_tester_name": (94, 182, 9), "init_company": (236, 183, 9),
    "init_cert": (356, 183, 9), "init_test_date": (464, 183, 9), "repaired_by": (95, 158, 9),
    "repair_company": (239, 156, 9), "repair_cert": (355, 159, 9), "repair_date": (464, 162, 9),
    "final_tester_name": (93, 135, 9), "final_company": (239, 135, 9), "final_cert": (354, 136, 9),
    "final_test_date": (468, 139, 9), "signature_date": (433, 84, 9),
}
JAX_SIG_X, JAX_SIG_Y, JAX_SIG_W, JAX_SIG_H = 161, 68, 160, 22
JAX_CHECKBOXES = {
    "COMM_ANNUAL": (214, 545), "COMM_REPAIR": (286, 544), "COMM_REPLACEMENT": (358, 545), "COMM_NEW_INSTALL": (463, 545),
    "COMM_FIRE": (214, 523), "COMM_IRRIGATION": (294, 522), "COMM_PROCESS": (362, 521), "COMM_POTABLE": (472, 523),
    "COMM_FIRE_BYPASS": (215, 510), "RECLAIM_YES": (421, 511), "RECLAIM_NO": (459, 510), "RES_ANNUAL": (210, 489),
    "RES_REPAIR": (280, 488), "RES_REPLACEMENT": (358, 489), "RES_NEW_INSTALL": (462, 489), "RES_POTABLE": (202, 466),
    "RES_IRRIGATION": (255, 465), "RES_RECLAIM_YES": (434, 466), "RES_RECLAIM_NO": (472, 464), "INIT_CV1_CLOSED": (139, 363),
    "INIT_CV2_CLOSED": (235, 362), "INIT_RV_OPENED": (331, 356), "INIT_RV_DIDNOT": (336, 329), "INIT_PVB_AIOPEN": (445, 359),
    "INIT_PVB_AIDNOT": (451, 323), "FINAL_CV1_CLOSED": (138, 306), "FINAL_CV2_CLOSED": (236, 301),
    "FINAL_RV_OPENED": (331, 296), "FINAL_PVB_SAT": (450, 290), "JAX_PASSED": (300, 106), "JAX_FAILED": (358, 108),
}

# ── All widget keys that must be wiped on profile load ──────────────────────
UNITED_TESTER_DISPLAY_KEYS = [
    "u_gmfg_display", "u_gsn_display", "u_cal_display",
    "u_tech_display", "u_cert_display", "u_recert_display",
]
JAX_TESTER_DISPLAY_KEYS = [
    "j_itn_display", "j_ico_display", "j_ic_display",
    "j_rb_display",  "j_rco_display", "j_rc_display",
    "j_ftn_display", "j_fco_display", "j_fc_display",
]
UNITED_TESTER_WIDGET_KEYS = ["u_gmfg", "u_gsn", "u_cal", "u_tech", "u_cert", "u_recert"]
JAX_TESTER_WIDGET_KEYS    = ["j_itn", "j_ico", "j_ic", "j_rb", "j_rco", "j_rc", "j_ftn", "j_fco", "j_fc"]

TESTER_KEYS = ["gauge_mfg", "gauge_serial", "date_cal", "technician", "cert_no", "recert"]

JAX_TESTER_MAP = {
    "init_tester_name": "technician", "init_company": "company", "init_cert": "cert_no",
    "repaired_by":      "technician", "repair_company": "company", "repair_cert": "cert_no",
    "final_tester_name":"technician", "final_company":  "company", "final_cert":  "cert_no",
}


# ─────────────────────────────────────────────────────────────
# Drawing helpers
# ─────────────────────────────────────────────────────────────

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


# ─────────────────────────────────────────────────────────────
# Form / session helpers
# ─────────────────────────────────────────────────────────────

def _init_form(key, defaults=None):
    if key not in st.session_state:
        st.session_state[key] = defaults or {}


def _clear_widget_keys(keys):
    for key in keys:
        st.session_state.pop(key, None)


def clearable_input(label, form_key, field_key, widget_key, **kwargs):
    form = st.session_state[form_key]
    if widget_key not in st.session_state:
        st.session_state[widget_key] = form.get(field_key, "")
    val = st.text_input(label, key=widget_key, **kwargs)
    form[field_key] = val
    return val


def get_signature_image_reader(form_data=None):
    source_b64 = form_data.get("signature_b64") if form_data and form_data.get("signature_b64") else st.session_state.get("signature_b64")
    if source_b64:
        try:
            buf = BytesIO(base64.b64decode(source_b64))
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


def apply_profile_to_forms(profile: dict):
    _clear_widget_keys(
        UNITED_TESTER_WIDGET_KEYS + JAX_TESTER_WIDGET_KEYS
        + UNITED_TESTER_DISPLAY_KEYS + JAX_TESTER_DISPLAY_KEYS
    )

    if not profile:
        return

    if profile.get("signature_b64"):
        st.session_state["signature_b64"] = profile["signature_b64"]

    _init_form("united_form")
    united = st.session_state["united_form"]
    for tk in TESTER_KEYS:
        united[tk] = profile.get(tk, "")
    united["signature_b64"] = profile.get("signature_b64", "")

    _init_form("jax_form")
    jax = st.session_state["jax_form"]
    for jk, pk in JAX_TESTER_MAP.items():
        jax[jk] = profile.get(pk, "")
    for tk in TESTER_KEYS:
        jax[tk] = profile.get(tk, "")
    jax["signature_b64"] = profile.get("signature_b64", "")


# ─────────────────────────────────────────────────────────────
# PDF generators
# ─────────────────────────────────────────────────────────────

def generate_united_pdf(form_data: dict) -> bytes:
    reader = PdfReader(TEMPLATE_UNITED)
    writer = PdfWriter()
    writer.append(reader)
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
    for field, (x, y, sz) in UNITED_TEXT_FIELDS.items():
        put_text(c, form_data.get(field, ""), x, y, sz)
    atype = form_data.get("assembly_type", "")
    if atype == "RP": draw_x(c, *UNITED_CHECKBOXES["RP"])
    elif atype == "DC": draw_x(c, *UNITED_CHECKBOXES["DC"])
    elif atype == "PVB": draw_x(c, *UNITED_CHECKBOXES["PVB"])
    elif atype == "SVB": draw_x(c, *UNITED_CHECKBOXES["SVB"])
    svc = form_data.get("system_service", "")
    if svc == "Fire": draw_x(c, *UNITED_CHECKBOXES["FIRE"])
    elif svc == "Domestic": draw_x(c, *UNITED_CHECKBOXES["DOMESTIC"])
    elif svc == "Irrigation": draw_x(c, *UNITED_CHECKBOXES["IRRIGATION"])
    elif svc == "Attraction": draw_x(c, *UNITED_CHECKBOXES["ATTRACTION"])
    bypass = form_data.get("bypass", "")
    if bypass == "Yes": draw_x(c, *UNITED_CHECKBOXES["BYPASS_YES"])
    elif bypass == "No": draw_x(c, *UNITED_CHECKBOXES["BYPASS_NO"])
    cv1r = form_data.get("cv1_result", "")
    if cv1r == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["CV1_CLOSED"])
    elif cv1r == "Leaked": draw_x(c, *UNITED_CHECKBOXES["CV1_LEAKED"])
    cv2r = form_data.get("cv2_result", "")
    if cv2r == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["CV2_CLOSED"])
    elif cv2r == "Leaked": draw_x(c, *UNITED_CHECKBOXES["CV2_LEAKED"])
    rvr = form_data.get("rv_result", "")
    if rvr == "Opened": draw_x(c, *UNITED_CHECKBOXES["RV_OPENED"])
    elif rvr == "Did Not Open": draw_x(c, *UNITED_CHECKBOXES["RV_DIDNOTOPEN"])
    rvo = form_data.get("rv_out_result", "")
    if rvo == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["RV_OUT_CLOSED"])
    elif rvo == "Leaked": draw_x(c, *UNITED_CHECKBOXES["RV_OUT_LEAKED"])
    rvi = form_data.get("rv_in_result", "")
    if rvi == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["RV_IN_CLOSED"])
    elif rvi == "Leaked": draw_x(c, *UNITED_CHECKBOXES["RV_IN_LEAKED"])
    pai = form_data.get("pvb_ai_result", "")
    if pai == "Opened": draw_x(c, *UNITED_CHECKBOXES["PVB_AI_OPENED"])
    elif pai == "Did Not Open": draw_x(c, *UNITED_CHECKBOXES["PVB_AI_CLOSED"])
    pcv = form_data.get("pvb_cv_result", "")
    if pcv == "Leaked": draw_x(c, *UNITED_CHECKBOXES["PVB_CV_LEAKED"])
    elif pcv == "Held": draw_x(c, *UNITED_CHECKBOXES["PVB_CV_HELD"])
    ar = form_data.get("assembly_result", "")
    if ar == "PASSED": draw_x(c, *UNITED_CHECKBOXES["PASSED"])
    elif ar == "FAILED": draw_x(c, *UNITED_CHECKBOXES["FAILED"])
    repair = form_data.get("repair_desc", "")
    if repair:
        bx, by, bh, bl, bw = UNITED_REPAIR_BOX
        lines = wrap_text(repair, bw)
        c.setFont("Helvetica-Bold", 7)
        c.setFillColorRGB(1, 0, 0)
        for i, ln in enumerate(lines[:bl]):
            c.drawString(bx, by - i * bh, ln)
    sig_reader = get_signature_image_reader(form_data)
    if sig_reader:
        c.drawImage(sig_reader, UNITED_SIG_X, UNITED_SIG_Y, width=UNITED_SIG_W, height=UNITED_SIG_H, mask="auto")
    c.save()
    packet.seek(0)
    overlay_reader = PdfReader(packet)
    page = writer.pages[0]
    page.merge_page(overlay_reader.pages[0])
    out = BytesIO()
    writer.write(out)
    return out.getvalue()


def generate_jax_pdf(form_data: dict) -> bytes:
    reader = PdfReader(TEMPLATE_JAX)
    writer = PdfWriter()
    writer.append(reader)
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(JAX_PAGE_W, JAX_PAGE_H))
    for field, (x, y, sz) in JAX_TEXT_FIELDS.items():
        put_text(c, form_data.get(field, ""), x, y, sz)
    comm_tp = form_data.get("comm_test_purpose", "")
    if comm_tp == "Annual": draw_x(c, *JAX_CHECKBOXES["COMM_ANNUAL"])
    elif comm_tp == "Repair": draw_x(c, *JAX_CHECKBOXES["COMM_REPAIR"])
    elif comm_tp == "Replacement": draw_x(c, *JAX_CHECKBOXES["COMM_REPLACEMENT"])
    elif comm_tp == "New Install": draw_x(c, *JAX_CHECKBOXES["COMM_NEW_INSTALL"])
    comm_st = form_data.get("comm_service_type", "")
    if comm_st == "Fire": draw_x(c, *JAX_CHECKBOXES["COMM_FIRE"])
    elif comm_st == "Irrigation": draw_x(c, *JAX_CHECKBOXES["COMM_IRRIGATION"])
    elif comm_st == "Process": draw_x(c, *JAX_CHECKBOXES["COMM_PROCESS"])
    elif comm_st == "Potable": draw_x(c, *JAX_CHECKBOXES["COMM_POTABLE"])
    comm_fp = form_data.get("comm_fire_bypass", "")
    if comm_fp: draw_x(c, *JAX_CHECKBOXES["COMM_FIRE_BYPASS"])
    comm_rc = form_data.get("comm_reclaim", "")
    if comm_rc == "Yes": draw_x(c, *JAX_CHECKBOXES["RECLAIM_YES"])
    elif comm_rc == "No": draw_x(c, *JAX_CHECKBOXES["RECLAIM_NO"])
    res_tp = form_data.get("res_test_purpose", "")
    if res_tp == "Annual": draw_x(c, *JAX_CHECKBOXES["RES_ANNUAL"])
    elif res_tp == "Repair": draw_x(c, *JAX_CHECKBOXES["RES_REPAIR"])
    elif res_tp == "Replacement": draw_x(c, *JAX_CHECKBOXES["RES_REPLACEMENT"])
    elif res_tp == "New Install": draw_x(c, *JAX_CHECKBOXES["RES_NEW_INSTALL"])
    res_st = form_data.get("res_service_type", "")
    if res_st == "Potable": draw_x(c, *JAX_CHECKBOXES["RES_POTABLE"])
    elif res_st == "Irrigation": draw_x(c, *JAX_CHECKBOXES["RES_IRRIGATION"])
    res_rc = form_data.get("res_reclaim", "")
    if res_rc == "Yes": draw_x(c, *JAX_CHECKBOXES["RES_RECLAIM_YES"])
    elif res_rc == "No": draw_x(c, *JAX_CHECKBOXES["RES_RECLAIM_NO"])
    icv1 = form_data.get("init_cv1_result", "")
    if icv1 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["INIT_CV1_CLOSED"])
    icv2 = form_data.get("init_cv2_result", "")
    if icv2 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["INIT_CV2_CLOSED"])
    irv = form_data.get("init_rv_result", "")
    if irv == "Opened": draw_x(c, *JAX_CHECKBOXES["INIT_RV_OPENED"])
    elif irv == "Did Not Open": draw_x(c, *JAX_CHECKBOXES["INIT_RV_DIDNOT"])
    ipvb = form_data.get("init_pvb_result", "")
    if ipvb == "Air Inlet Opened": draw_x(c, *JAX_CHECKBOXES["INIT_PVB_AIOPEN"])
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
    if ar == "PASSED": draw_x(c, *JAX_CHECKBOXES["JAX_PASSED"])
    elif ar == "FAILED": draw_x(c, *JAX_CHECKBOXES["JAX_FAILED"])
    sig_reader = get_signature_image_reader(form_data)
    if sig_reader:
        c.drawImage(sig_reader, JAX_SIG_X, JAX_SIG_Y, width=JAX_SIG_W, height=JAX_SIG_H, mask="auto")
    c.save()
    packet.seek(0)
    overlay_reader = PdfReader(packet)
    page = writer.pages[0]
    page.merge_page(overlay_reader.pages[0])
    out = BytesIO()
    writer.write(out)
    return out.getvalue()


# ─────────────────────────────────────────────────────────────
# Sidebar
# ─────────────────────────────────────────────────────────────

def render_technician_sidebar():
    st.sidebar.title("👤 Technician Profile")
    names = get_technician_names()
    prev_sel = st.session_state.get("_sidebar_tech_sel", "")
    if prev_sel not in names:
        prev_sel = ""
    selected = st.sidebar.selectbox("Select technician", names, index=names.index(prev_sel) if prev_sel in names else 0, key="sidebar_tech_select")
    st.session_state["_sidebar_tech_sel"] = selected
    last_loaded = st.session_state.get("_last_loaded_tech", None)
    if selected and selected != last_loaded:
        profile = get_technician_profile(selected)
        apply_profile_to_forms(profile)
        st.session_state["_last_loaded_tech"] = selected
        st.rerun()
    if st.sidebar.button("📥 Reload Profile", key="load_profile_btn", disabled=not bool(selected)):
        profile = get_technician_profile(selected)
        apply_profile_to_forms(profile)
        st.session_state["_last_loaded_tech"] = selected
        st.sidebar.success(f"Loaded: {selected}")
        st.rerun()
    if selected:
        confirm_key = "_confirm_delete_profile"
        if not st.session_state.get(confirm_key, False):
            if st.sidebar.button("🗑️ Delete Profile", key="delete_profile_btn"):
                st.session_state[confirm_key] = True
                st.rerun()
        else:
            st.sidebar.warning(f"Delete **{selected}**? This cannot be undone.")
            col_yes, col_no = st.sidebar.columns(2)
            with col_yes:
                if st.button("✅ Yes, delete", key="confirm_delete_yes"):
                    ok, msg = delete_technician_profile(selected)
                    st.session_state[confirm_key] = False
                    if ok:
                        st.session_state["_sidebar_tech_sel"] = ""
                        st.session_state["_last_loaded_tech"] = None
                        clear_signature()
                        st.sidebar.success("Profile deleted.")
                    else:
                        st.sidebar.warning(msg)
                    st.rerun()
            with col_no:
                if st.button("❌ Cancel", key="confirm_delete_no"):
                    st.session_state[confirm_key] = False
                    st.rerun()
    st.sidebar.divider()
    with st.sidebar.expander("✏️ Edit / Add Profile", expanded=False):
        current = get_technician_profile(selected) if selected else {}
        prof_name = st.text_input("Name (key)", value=selected, key="prof_name")
        prof_co = st.text_input("Company", value=current.get("company", ""), key="prof_co")
        prof_cert = st.text_input("Cert No.", value=current.get("cert_no", ""), key="prof_cert")
        prof_rec = st.text_input("Re-Cert", value=current.get("recert", ""), key="prof_rec")
        prof_gmfg = st.text_input("Gauge Mfg", value=current.get("gauge_mfg", ""), key="prof_gmfg")
        prof_gsn = st.text_input("Gauge SN", value=current.get("gauge_serial", ""), key="prof_gsn")
        prof_cal = st.text_input("Date Cal", value=current.get("date_cal", ""), key="prof_cal")
        if current.get("signature_b64"):
            st.caption("Saved profile signature:")
            st.image(base64.b64decode(current["signature_b64"]), width=180)
        if st.button("💾 Save Profile", key="save_profile_btn"):
            if prof_name.strip():
                new_profile = {
                    "technician": prof_name.strip(),
                    "company": prof_co.strip(),
                    "cert_no": prof_cert.strip(),
                    "recert": prof_rec.strip(),
                    "gauge_mfg": prof_gmfg.strip(),
                    "gauge_serial": prof_gsn.strip(),
                    "date_cal": prof_cal.strip(),
                    "signature_b64": st.session_state.get("signature_b64", current.get("signature_b64", "")),
                }
                ok, msg = upsert_technician_profile(prof_name.strip(), new_profile)
                if ok:
                    st.success(msg)
                else:
                    st.warning(msg)
                st.session_state["_sidebar_tech_sel"] = prof_name.strip()
                st.session_state["_last_loaded_tech"] = None
                st.rerun()
            else:
                st.error("Name cannot be empty.")
    st.sidebar.divider()
    st.sidebar.markdown("**Signature**")
    sig_b64 = st.session_state.get("signature_b64")
    if sig_b64:
        st.sidebar.image(base64.b64decode(sig_b64), caption="Current signature", use_container_width=True)
    upload = st.sidebar.file_uploader("Upload signature (PNG preferred)", type=["png", "jpg", "jpeg"], key="sig_upload")
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
        canvas_result = st_canvas(fill_color="rgba(255,255,255,0)", stroke_width=2, stroke_color="#000000", background_color="#FFFFFF", height=80, width=220, drawing_mode="freedraw", key="sb_sig_canvas")
        if canvas_result.image_data is not None:
            arr = canvas_result.image_data
            if arr.max() > 0 and arr[:, :, 3].max() > 0:
                save_signature(arr)
    except ImportError:
        st.sidebar.info("Install streamlit-drawable-canvas for signature support.")
    col_sig1, col_sig2 = st.sidebar.columns(2)
    with col_sig1:
        if st.button("🗑️ Clear Sig", key="sb_clear_sig"):
            clear_signature()
            st.rerun()
    with col_sig2:
        if selected and st.button("💾 Save Sig to Profile", key="save_sig_to_profile"):
            current = get_technician_profile(selected)
            current["signature_b64"] = st.session_state.get("signature_b64", "")
            ok, msg = upsert_technician_profile(selected, current)
            if ok:
                st.sidebar.success("Signature saved to profile.")
            else:
                st.sidebar.warning(msg)


# ─────────────────────────────────────────────────────────────
# Tester panels
# ─────────────────────────────────────────────────────────────

def render_tester_panel_united():
    _init_form("united_form")
    form = st.session_state["united_form"]
    selected_tech = st.session_state.get("_sidebar_tech_sel", "")
    st.divider()
    st.markdown("**🔧 Tester / Technician Info**")
    if not selected_tech:
        st.warning("Select a technician profile in the sidebar to populate tester information.")
        return
    st.caption(f"Profile loaded: **{selected_tech}**")
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Gauge Mfg",       value=form.get("gauge_mfg",    ""), disabled=True, key="u_gmfg_display")
        st.text_input("Gauge Serial",     value=form.get("gauge_serial", ""), disabled=True, key="u_gsn_display")
        st.text_input("Date Calibrated",  value=form.get("date_cal",     ""), disabled=True, key="u_cal_display")
    with col2:
        st.text_input("Technician",       value=form.get("technician",   ""), disabled=True, key="u_tech_display")
        st.text_input("Cert No.",         value=form.get("cert_no",      ""), disabled=True, key="u_cert_display")
        st.text_input("Re-Cert Date",     value=form.get("recert",       ""), disabled=True, key="u_recert_display")
    sig_b64 = st.session_state.get("signature_b64") or form.get("signature_b64")
    if sig_b64:
        st.image(base64.b64decode(sig_b64), caption="Signature on file", width=200)
    else:
        st.caption("⚠️ No signature loaded. Upload or draw one in the sidebar and save it to the profile.")


def render_tester_panel_jax():
    _init_form("jax_form")
    form = st.session_state["jax_form"]
    selected_tech = st.session_state.get("_sidebar_tech_sel", "")
    st.divider()
    st.markdown("**🔧 Tester / Technician Info**")
    if not selected_tech:
        st.warning("Select a technician profile in the sidebar to populate tester information.")
        return
    st.caption(f"Profile loaded: **{selected_tech}**")
    col_itn, col_ico, col_ic = st.columns(3)
    with col_itn:
        st.text_input("Init Tester Name", value=form.get("init_tester_name", ""), disabled=True, key="j_itn_display")
    with col_ico:
        st.text_input("Init Company",     value=form.get("init_company",    ""), disabled=True, key="j_ico_display")
    with col_ic:
        st.text_input("Init Cert",        value=form.get("init_cert",       ""), disabled=True, key="j_ic_display")
    col_rb, col_rco, col_rc = st.columns(3)
    with col_rb:
        st.text_input("Repaired By",      value=form.get("repaired_by",     ""), disabled=True, key="j_rb_display")
    with col_rco:
        st.text_input("Repair Company",   value=form.get("repair_company",  ""), disabled=True, key="j_rco_display")
    with col_rc:
        st.text_input("Repair Cert",      value=form.get("repair_cert",     ""), disabled=True, key="j_rc_display")
    col_ftn, col_fco, col_fc = st.columns(3)
    with col_ftn:
        st.text_input("Final Tester Name",value=form.get("final_tester_name",""), disabled=True, key="j_ftn_display")
    with col_fco:
        st.text_input("Final Company",    value=form.get("final_company",   ""), disabled=True, key="j_fco_display")
    with col_fc:
        st.text_input("Final Cert",       value=form.get("final_cert",      ""), disabled=True, key="j_fc_display")
    sig_b64 = st.session_state.get("signature_b64") or form.get("signature_b64")
    if sig_b64:
        st.image(base64.b64decode(sig_b64), caption="Signature on file", width=200)
    else:
        st.caption("⚠️ No signature loaded. Upload or draw one in the sidebar and save it to the profile.")


# ─────────────────────────────────────────────────────────────
# Form tabs
# ─────────────────────────────────────────────────────────────

def render_united_tab():
    _init_form("united_form")
    form = st.session_state["united_form"]
    st.subheader("📋 United Fire — Backflow Test Report")
    synced_date_input("Inspection Date", "united_form", "date", "u_date_picker", ["date", "test_date"])
    clearable_input("Branch", "united_form", "branch", "u_branch")
    clearable_input("AHJ", "united_form", "ahj", "u_ahj")
    col4, col5 = st.columns(2)
    with col4:
        clearable_input("Customer Name", "united_form", "customer_name", "u_cust")
    with col5:
        clearable_input("Street Address", "united_form", "street_address", "u_addr")
    clearable_input("Location / Description", "united_form", "location", "u_loc")
    st.divider()
    st.markdown("**Assembly Info**")
    col6, col7 = st.columns(2)
    with col6:
        clearable_input("Serial Number", "united_form", "serial_number", "u_sn")
        clearable_input("Manufacturer", "united_form", "manufacturer", "u_mfg")
    with col7:
        clearable_input("Model", "united_form", "model", "u_mdl")
        clearable_input("Size", "united_form", "size", "u_sz")
    atype_opts = ["", "RP", "DC", "PVB", "SVB"]
    svc_opts = ["", "Fire", "Domestic", "Irrigation", "Attraction"]
    bp_opts = ["", "Yes", "No"]
    col10, col11, col12 = st.columns(3)
    with col10:
        form["assembly_type"] = st.selectbox("Assembly Type", atype_opts, index=atype_opts.index(form.get("assembly_type", "")) if form.get("assembly_type", "") in atype_opts else 0, key="u_assembly_type")
    with col11:
        form["system_service"] = st.selectbox("System Service", svc_opts, index=svc_opts.index(form.get("system_service", "")) if form.get("system_service", "") in svc_opts else 0, key="u_system_service")
    with col12:
        form["bypass"] = st.selectbox("Fire Bypass", bp_opts, index=bp_opts.index(form.get("bypass", "")) if form.get("bypass", "") in bp_opts else 0, key="u_bypass")
    st.divider()
    st.markdown("**Test Results**")
    atype = form.get("assembly_type", "")
    if atype in ("", "RP", "DC"):
        col_a, col_b = st.columns(2)
        with col_a:
            cv_opts = ["", "Closed Tight", "Leaked"]
            form["cv1_result"] = st.selectbox("CV1 Result", cv_opts, index=cv_opts.index(form.get("cv1_result", "")) if form.get("cv1_result", "") in cv_opts else 0, key="u_cv1_result")
        with col_b:
            clearable_input("CV1 DP (psi)", "united_form", "cv1_dp", "u_cv1dp")
        col_a, col_b = st.columns(2)
        with col_a:
            cv_opts = ["", "Closed Tight", "Leaked"]
            form["cv2_result"] = st.selectbox("CV2 Result", cv_opts, index=cv_opts.index(form.get("cv2_result", "")) if form.get("cv2_result", "") in cv_opts else 0, key="u_cv2_result")
        with col_b:
            clearable_input("CV2 DP (psi)", "united_form", "cv2_dp", "u_cv2dp")
    if atype in ("", "RP"):
        col_a, col_b = st.columns(2)
        with col_a:
            rv_opts = ["", "Opened", "Did Not Open"]
            form["rv_result"] = st.selectbox("RV Result", rv_opts, index=rv_opts.index(form.get("rv_result", "")) if form.get("rv_result", "") in rv_opts else 0, key="u_rv_result")
        with col_b:
            clearable_input("RV Opened At (psi)", "united_form", "rv_psi", "u_rvpsi")
        col_a, col_b = st.columns(2)
        with col_a:
            out_opts = ["", "Closed Tight", "Leaked"]
            form["rv_out_result"] = st.selectbox("RV Outlet", out_opts, index=out_opts.index(form.get("rv_out_result", "")) if form.get("rv_out_result", "") in out_opts else 0, key="u_rv_out_result")
        with col_b:
            in_opts = ["", "Closed Tight", "Leaked"]
            form["rv_in_result"] = st.selectbox("RV Inlet", in_opts, index=in_opts.index(form.get("rv_in_result", "")) if form.get("rv_in_result", "") in in_opts else 0, key="u_rv_in_result")
    if atype in ("", "PVB", "SVB"):
        col_a, col_b = st.columns(2)
        with col_a:
            ai_opts = ["", "Opened", "Did Not Open"]
            form["pvb_ai_result"] = st.selectbox("PVB Air Inlet", ai_opts, index=ai_opts.index(form.get("pvb_ai_result", "")) if form.get("pvb_ai_result", "") in ai_opts else 0, key="u_pvb_ai_result")
        with col_b:
            clearable_input("PVB AI (psi)", "united_form", "pvb_ai_psi", "u_aipsi")
        col_a, col_b = st.columns(2)
        with col_a:
            pcv_opts = ["", "Leaked", "Held"]
            form["pvb_cv_result"] = st.selectbox("PVB CV Result", pcv_opts, index=pcv_opts.index(form.get("pvb_cv_result", "")) if form.get("pvb_cv_result", "") in pcv_opts else 0, key="u_pvb_cv_result")
        with col_b:
            clearable_input("PVB CV (psi)", "united_form", "pvb_cv_psi", "u_cvpsi")
    st.divider()
    st.caption(f"Test Date will match Inspection Date: {form.get('test_date', '')}")
    ar_opts = ["", "PASSED", "FAILED"]
    form["assembly_result"] = st.selectbox("Assembly Result", ar_opts, index=ar_opts.index(form.get("assembly_result", "")) if form.get("assembly_result", "") in ar_opts else 0, key="u_assembly_result")
    clearable_input("Repair Description", "united_form", "repair_desc", "u_rep")
    render_tester_panel_united()
    if st.button("🖨️ Generate & Save PDF", key="u_gen_pdf"):
        selected_profile = st.session_state.get("_sidebar_tech_sel", "")
        if not selected_profile:
            st.error("You must select a technician profile before generating a PDF.")
        else:
            try:
                form["signature_b64"] = st.session_state.get("signature_b64", form.get("signature_b64", ""))
                pdf_bytes = generate_united_pdf(form)
                save_ok, save_msg, fname = autosave_job(form, pdf_bytes, "united")
                if save_ok:
                    st.success(save_msg)
                else:
                    st.warning(save_msg)
                # Store for immediate download
                st.session_state["_last_pdf_bytes"] = pdf_bytes
                st.session_state["_last_pdf_name"] = fname
            except Exception as e:
                st.error(f"PDF error: {e}")

    # Always show download button if PDF was just generated this session
    if st.session_state.get("_last_pdf_bytes") and st.session_state.get("_last_pdf_name", "").startswith("united_"):
        st.download_button(
            "📥 Download PDF to Device",
            st.session_state["_last_pdf_bytes"],
            st.session_state["_last_pdf_name"],
            "application/pdf",
            key="u_dl_pdf",
        )


def render_jax_tab():
    _init_form("jax_form")
    form = st.session_state["jax_form"]
    st.subheader("📋 Jacksonville (JEA) — Backflow Test Report")
    synced_date_input("Inspection Date", "jax_form", "signature_date", "j_sig_date_picker", ["signature_date", "init_test_date", "final_test_date", "repair_date"])
    col1, col2 = st.columns(2)
    with col1:
        clearable_input("Premises Name", "jax_form", "premises_name", "j_prem")
        clearable_input("Service Address", "jax_form", "service_address", "j_sa")
        clearable_input("Physical Location", "jax_form", "physical_location", "j_pl")
        clearable_input("JEA Account", "jax_form", "jea_account", "j_acct")
    with col2:
        clearable_input("Owner Name", "jax_form", "owner_name", "j_own")
        clearable_input("Mailing Address", "jax_form", "mailing_address", "j_ma")
        clearable_input("Contact Phone", "jax_form", "contact_phone", "j_ph")
        clearable_input("Meter Number", "jax_form", "meter_number", "j_meter")
    st.divider()
    ctp_opts = ["", "Annual", "Repair", "Replacement", "New Install"]
    cst_opts = ["", "Fire", "Irrigation", "Process", "Potable"]
    yn_opts = ["", "Yes", "No"]
    rst_opts = ["", "Potable", "Irrigation"]
    col1, col2, col3 = st.columns(3)
    with col1:
        form["comm_test_purpose"] = st.selectbox("Comm Test Purpose", ctp_opts, index=ctp_opts.index(form.get("comm_test_purpose", "")) if form.get("comm_test_purpose", "") in ctp_opts else 0, key="j_comm_test_purpose")
    with col2:
        form["comm_service_type"] = st.selectbox("Comm Service Type", cst_opts, index=cst_opts.index(form.get("comm_service_type", "")) if form.get("comm_service_type", "") in cst_opts else 0, key="j_comm_service_type")
    with col3:
        form["comm_reclaim"] = st.selectbox("Comm Reclaim Water", yn_opts, index=yn_opts.index(form.get("comm_reclaim", "")) if form.get("comm_reclaim", "") in yn_opts else 0, key="j_comm_reclaim")
    col1, col2, col3 = st.columns(3)
    with col1:
        form["res_test_purpose"] = st.selectbox("Res Test Purpose", ctp_opts, index=ctp_opts.index(form.get("res_test_purpose", "")) if form.get("res_test_purpose", "") in ctp_opts else 0, key="j_res_test_purpose")
    with col2:
        form["res_service_type"] = st.selectbox("Res Service Type", rst_opts, index=rst_opts.index(form.get("res_service_type", "")) if form.get("res_service_type", "") in rst_opts else 0, key="j_res_service_type")
    with col3:
        form["res_reclaim"] = st.selectbox("Res Reclaim Water", yn_opts, index=yn_opts.index(form.get("res_reclaim", "")) if form.get("res_reclaim", "") in yn_opts else 0, key="j_res_reclaim")
    st.divider()
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        clearable_input("Device Type", "jax_form", "device_type", "j_dt")
    with col2:
        clearable_input("Manufacturer", "jax_form", "manufacturer", "j_mfg")
    with col3:
        clearable_input("Size", "jax_form", "size", "j_sz")
    with col4:
        clearable_input("Model Number", "jax_form", "model_number", "j_mn")
    col1, col2 = st.columns(2)
    with col1:
        clearable_input("Serial Number", "jax_form", "serial_number", "j_sn")
    with col2:
        clearable_input("Install Date", "jax_form", "install_date", "j_id")
    st.divider()
    cv_opts = ["", "Closed Tight", "Leaked"]
    rv_opts = ["", "Opened", "Did Not Open"]
    pvb_opts = ["", "Air Inlet Opened", "Air Inlet Did Not"]
    fpvb_opts = ["", "Satisfactory", "Unsatisfactory"]
    col1, col2 = st.columns(2)
    with col1:
        form["init_cv1_result"] = st.selectbox("Init CV1", cv_opts, index=cv_opts.index(form.get("init_cv1_result", "")) if form.get("init_cv1_result", "") in cv_opts else 0, key="j_init_cv1_result")
    with col2:
        clearable_input("Init CV1 (psi)", "jax_form", "init_cv1_psi", "j_icv1p")
    col1, col2 = st.columns(2)
    with col1:
        form["init_cv2_result"] = st.selectbox("Init CV2", cv_opts, index=cv_opts.index(form.get("init_cv2_result", "")) if form.get("init_cv2_result", "") in cv_opts else 0, key="j_init_cv2_result")
    with col2:
        clearable_input("Init CV2 (psi)", "jax_form", "init_cv2_psi", "j_icv2p")
    col1, col2 = st.columns(2)
    with col1:
        form["init_rv_result"] = st.selectbox("Init RV", rv_opts, index=rv_opts.index(form.get("init_rv_result", "")) if form.get("init_rv_result", "") in rv_opts else 0, key="j_init_rv_result")
    with col2:
        clearable_input("Init RV (psi)", "jax_form", "init_rv_psi", "j_irvp")
    col1, col2 = st.columns(2)
    with col1:
        form["init_pvb_result"] = st.selectbox("Init PVB", pvb_opts, index=pvb_opts.index(form.get("init_pvb_result", "")) if form.get("init_pvb_result", "") in pvb_opts else 0, key="j_init_pvb_result")
    with col2:
        clearable_input("Init PVB (psi)", "jax_form", "init_pvb_psi", "j_ipvbp")
    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        form["final_cv1_result"] = st.selectbox("Final CV1", cv_opts, index=cv_opts.index(form.get("final_cv1_result", "")) if form.get("final_cv1_result", "") in cv_opts else 0, key="j_final_cv1_result")
    with col2:
        clearable_input("Final CV1 (psi)", "jax_form", "final_cv1_psi", "j_fcv1p")
    col1, col2 = st.columns(2)
    with col1:
        form["final_cv2_result"] = st.selectbox("Final CV2", cv_opts, index=cv_opts.index(form.get("final_cv2_result", "")) if form.get("final_cv2_result", "") in cv_opts else 0, key="j_final_cv2_result")
    with col2:
        clearable_input("Final CV2 (psi)", "jax_form", "final_cv2_psi", "j_fcv2p")
    col1, col2 = st.columns(2)
    with col1:
        form["final_rv_result"] = st.selectbox("Final RV", rv_opts, index=rv_opts.index(form.get("final_rv_result", "")) if form.get("final_rv_result", "") in rv_opts else 0, key="j_final_rv_result")
    with col2:
        clearable_input("Final RV (psi)", "jax_form", "final_rv_psi", "j_frvp")
    form["final_pvb_result"] = st.selectbox("Final PVB", fpvb_opts, index=fpvb_opts.index(form.get("final_pvb_result", "")) if form.get("final_pvb_result", "") in fpvb_opts else 0, key="j_final_pvb_result")
    ar_opts = ["", "PASSED", "FAILED"]
    form["assembly_result"] = st.selectbox("Assembly Result", ar_opts, index=ar_opts.index(form.get("assembly_result", "")) if form.get("assembly_result", "") in ar_opts else 0, key="j_assembly_result")
    clearable_input("Repairs / Notes", "jax_form", "repairs", "j_rep")
    render_tester_panel_jax()
    if st.button("🖨️ Generate & Save PDF", key="j_gen_pdf"):
        selected_profile = st.session_state.get("_sidebar_tech_sel", "")
        if not selected_profile:
            st.error("You must select a technician profile before generating a PDF.")
        else:
            try:
                form["signature_b64"] = st.session_state.get("signature_b64", form.get("signature_b64", ""))
                pdf_bytes = generate_jax_pdf(form)
                save_ok, save_msg, fname = autosave_job(form, pdf_bytes, "jax")
                if save_ok:
                    st.success(save_msg)
                else:
                    st.warning(save_msg)
                st.session_state["_last_pdf_bytes"] = pdf_bytes
                st.session_state["_last_pdf_name"] = fname
            except Exception as e:
                st.error(f"PDF error: {e}")

    if st.session_state.get("_last_pdf_bytes") and st.session_state.get("_last_pdf_name", "").startswith("jax_"):
        st.download_button(
            "📥 Download PDF to Device",
            st.session_state["_last_pdf_bytes"],
            st.session_state["_last_pdf_name"],
            "application/pdf",
            key="j_dl_pdf",
        )


# ─────────────────────────────────────────────────────────────
# Jobs tab
# ─────────────────────────────────────────────────────────────

def render_jobs_tab():
    st.subheader("📁 Saved Jobs")

    today_str = datetime.now().strftime("%Y-%m-%d")
    today_display = datetime.now().strftime("%B %d, %Y")

    # ── Top action bar ────────────────────────────────────────
    col_refresh, col_zip, col_excel, col_clear = st.columns([1, 2, 2, 2])

    with col_refresh:
        if st.button("🔄 Refresh", key="jobs_refresh"):
            st.session_state.pop("_jobs_cache", None)
            st.rerun()

    jobs = load_jobs_cached()
    today_jobs = [j for j in jobs if j.get("saved_date", "") == today_str]

    with col_zip:
        if today_jobs:
            with st.spinner("Building ZIP..."):
                zip_bytes = build_zip_with_pdfs(jobs)
            st.download_button(
                f"📦 Download Today's ZIP ({len(today_jobs)} PDF{'s' if len(today_jobs) != 1 else ''} + Excel)",
                zip_bytes,
                f"backflow_jobs_{today_str}.zip",
                "application/zip",
                key="jobs_zip_dl",
            )
        else:
            st.button("📦 Download Today's ZIP", disabled=True, key="jobs_zip_disabled",
                      help="No jobs generated today yet.")

    with col_excel:
        if jobs:
            excel_bytes = build_jobs_excel(jobs)
            st.download_button(
                "📊 Download Excel Summary",
                excel_bytes,
                f"backflow_jobs_{today_str}.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="jobs_excel_dl",
            )

    with col_clear:
        confirm_clear = st.session_state.get("_confirm_clear_reports", False)
        if not confirm_clear:
            if st.button("🗑️ Clear All Reports", key="clear_reports_btn", type="secondary"):
                st.session_state["_confirm_clear_reports"] = True
                st.rerun()
        else:
            st.warning("⚠️ This will delete **all PDFs and the job log** from GitHub permanently.")
            cc1, cc2 = st.columns(2)
            with cc1:
                if st.button("✅ Yes, clear all", key="confirm_clear_yes"):
                    with st.spinner("Clearing reports..."):
                        ok, msg = clear_all_reports()
                    st.session_state["_confirm_clear_reports"] = False
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)
                    st.rerun()
            with cc2:
                if st.button("❌ Cancel", key="confirm_clear_no"):
                    st.session_state["_confirm_clear_reports"] = False
                    st.rerun()

    st.divider()

    if not jobs:
        st.info("No jobs saved yet. Generate a PDF from the United Fire or Jacksonville tab to autosave a job here.")
        return

    # ── Today's jobs ──────────────────────────────────────────
    if today_jobs:
        st.markdown(f"#### 📅 Today — {today_display} ({len(today_jobs)} job{'s' if len(today_jobs) != 1 else ''})")
        for job in reversed(today_jobs):
            result = job.get("assembly_result", "")
            result_icon = "✅" if result == "PASSED" else ("❌" if result == "FAILED" else "❔")
            with st.expander(f"{result_icon} {job.get('customer', 'Unknown')} — {job.get('date', '')} [{job.get('form_type','').upper()}]", expanded=False):
                c1, c2, c3 = st.columns(3)
                c1.metric("Result", result or "—")
                c2.metric("Technician", job.get("technician", "—"))
                c3.metric("Saved", job.get("saved_at", "—"))
                st.write(f"**Address:** {job.get('address', '—')}")
                st.write(f"**Serial #:** {job.get('serial_number', '—')} | **Assembly:** {job.get('assembly_type', '—')}")
                st.write(f"**File:** `{job.get('filename', '—')}`")
                pdf_url = job.get("pdf_url", "")
                if pdf_url:
                    st.markdown(f"[🔗 View PDF on GitHub]({pdf_url})")
    else:
        st.info(f"No jobs generated today ({today_display}). Previous jobs are in the history below.")

    # ── Older jobs (collapsed) ────────────────────────────────
    older_jobs = [j for j in jobs if j.get("saved_date", "") != today_str]
    if older_jobs:
        with st.expander(f"📂 Previous Jobs ({len(older_jobs)} record{'s' if len(older_jobs) != 1 else ''})", expanded=False):
            for job in reversed(older_jobs):
                result = job.get("assembly_result", "")
                result_icon = "✅" if result == "PASSED" else ("❌" if result == "FAILED" else "❔")
                st.markdown(
                    f"{result_icon} **{job.get('customer', 'Unknown')}** — "
                    f"{job.get('date', '')} [{job.get('form_type','').upper()}] — "
                    f"_{job.get('saved_at', '')}_"
                )
            st.caption("Use 'Download Excel Summary' for a full export of all records.")


# ─────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────

def main():
    st.set_page_config(page_title="Backflow Test Reports", page_icon="📋", layout="wide")
    render_technician_sidebar()
    tab_united, tab_jax, tab_jobs = st.tabs(["🔵 United Fire", "🔴 Jacksonville (JEA)", "📁 Jobs"])
    with tab_united:
        render_united_tab()
    with tab_jax:
        render_jax_tab()
    with tab_jobs:
        render_jobs_tab()


if __name__ == "__main__":
    main()
