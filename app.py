import streamlit as st
import streamlit.components.v1 as components
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import json, os, re, base64, tempfile, zipfile
from datetime import date
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ContentStream
from PIL import Image
import numpy as np

# ---------------------------------------------------------------------------
# Paths & constants
# ---------------------------------------------------------------------------
TEMPLATE_UNITED   = "backflow_template.pdf"
TEMPLATE_JAX      = "jacksonville_template.pdf"
SIG_FILE          = "signature_b64.txt"
PAGE_W, PAGE_H     = 612, 792   # United Fire (US Letter)

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

UNITED_NEXT_REPORT_KEEP = {
    "branch", "ahj", "customer_name", "street_address",
    "manufacturer", "model", "size", "assembly_type", "system_service",
    "gauge_mfg", "gauge_serial", "date_cal", "technician", "cert_no", "recert",
}
UNITED_NEW_JOB_KEEP = {
    "gauge_mfg", "gauge_serial", "date_cal", "technician", "cert_no", "recert",
}

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

JAX_SIG_X, JAX_SIG_Y, JAX_SIG_W, JAX_SIG_H = 161, 82, 160, 22

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
    # INIT_CV1_LEAKED removed
    "INIT_CV2_CLOSED":       (235, 362),
    # INIT_CV2_LEAKED removed
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
    "manufacturer", "model_number", "size", "device_type",
    "init_tester_name", "init_company", "init_cert",
    "final_tester_name", "final_company", "final_cert",
}
JAX_NEW_JOB_KEEP = {
    "init_tester_name", "init_company", "init_cert",
    "final_tester_name", "final_company", "final_cert",
}

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
# Signature helpers
# ---------------------------------------------------------------------------

def _load_sig_from_disk():
    if "sig_loaded" not in st.session_state:
        st.session_state["sig_loaded"] = True
        if os.path.exists(SIG_FILE):
            with open(SIG_FILE, "r") as fh:
                data = fh.read().strip()
            if data:
                st.session_state["signature_b64"] = data


def save_signature(img_array):
    buf = BytesIO()
    img = Image.fromarray(img_array.astype("uint8"), "RGBA")
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    st.session_state["signature_b64"] = b64
    with open(SIG_FILE, "w") as fh:
        fh.write(b64)


def clear_signature():
    st.session_state.pop("signature_b64", None)
    if os.path.exists(SIG_FILE):
        os.remove(SIG_FILE)


def get_signature_image_reader():
    b64 = st.session_state.get("signature_b64")
    if b64:
        buf = BytesIO(base64.b64decode(b64))
        buf.seek(0)
        return ImageReader(buf)
    return None


# ---------------------------------------------------------------------------
# Job folder (in-memory list of generated PDFs for ZIP download)
# ---------------------------------------------------------------------------

def _init_job_folder():
    if "job_folder" not in st.session_state:
        st.session_state.job_folder = []   # list of {"name": str, "bytes": bytes}


def add_to_job_folder(pdf_bytes: bytes, filename: str):
    _init_job_folder()
    # Avoid exact duplicates (same filename); overwrite if re-generated
    st.session_state.job_folder = [
        f for f in st.session_state.job_folder if f["name"] != filename
    ]
    st.session_state.job_folder.append({"name": filename, "bytes": pdf_bytes})


def build_zip() -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in st.session_state.job_folder:
            zf.writestr(item["name"], item["bytes"])
    buf.seek(0)
    return buf.read()


def render_job_folder_sidebar():
    """Render the job folder panel in the sidebar."""
    _init_job_folder()
    folder = st.session_state.job_folder
    st.sidebar.markdown("---")
    st.sidebar.subheader("📁 Job Folder")

    # Job name input for ZIP filename — persists across reports
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

    # Build ZIP — use entered job name if available, else derive from first PDF name
    zip_bytes = build_zip()
    entered_name = st.session_state.job_folder_name.strip()
    if entered_name:
        zip_name = re.sub(r"[^\w\s\-]", "", entered_name).strip()
    else:
        first_name = folder[0]["name"].replace(".pdf", "")
        zip_name = re.sub(r" - [^-]+$", "", first_name)
        zip_name = re.sub(r"[^\w\s\-]", "", zip_name).strip() or "Job"
    zip_filename = f"{zip_name} - Reports.zip"

    # Use a dynamic key tied to folder length so iOS re-renders the button
    # when new PDFs are added (avoids stale cached download)
    st.sidebar.download_button(
        label=f"📦 Download ZIP ({len(folder)} report{'s' if len(folder)!=1 else ''})",
        data=zip_bytes,
        file_name=zip_filename,
        mime="application/zip",
        use_container_width=True,
        key=f"zip_dl_{len(folder)}",
    )
    st.sidebar.caption("📱 iPhone/iPad: tap ZIP button → Share → Save to Files")

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
    overlay_page  = overlay_reader.pages[0]
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
    if bp == "YES":  draw_x(c, *UNITED_CHECKBOXES["BYPASS_YES"])
    elif bp == "NO": draw_x(c, *UNITED_CHECKBOXES["BYPASS_NO"])

    cv1 = form.get("cv1_result", "")
    if cv1 == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["CV1_CLOSED"])
    elif cv1 == "Leaked":     draw_x(c, *UNITED_CHECKBOXES["CV1_LEAKED"])

    cv2 = form.get("cv2_result", "")
    if cv2 == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["CV2_CLOSED"])
    elif cv2 == "Leaked":     draw_x(c, *UNITED_CHECKBOXES["CV2_LEAKED"])

    rv = form.get("rv_result", "")
    if rv == "Opened At":      draw_x(c, *UNITED_CHECKBOXES["RV_OPENED"])
    elif rv == "Did Not Open": draw_x(c, *UNITED_CHECKBOXES["RV_DIDNOTOPEN"])

    rvo = form.get("rv_out_result", "")
    if rvo == "Closed":   draw_x(c, *UNITED_CHECKBOXES["RV_OUT_CLOSED"])
    elif rvo == "Leaked": draw_x(c, *UNITED_CHECKBOXES["RV_OUT_LEAKED"])

    rvi = form.get("rv_in_result", "")
    if rvi == "Closed":   draw_x(c, *UNITED_CHECKBOXES["RV_IN_CLOSED"])
    elif rvi == "Leaked": draw_x(c, *UNITED_CHECKBOXES["RV_IN_LEAKED"])

    pvb_ai = form.get("pvb_ai_result", "")
    if pvb_ai == "Closed Tight": draw_x(c, *UNITED_CHECKBOXES["PVB_AI_CLOSED"])
    elif pvb_ai == "Opened At":  draw_x(c, *UNITED_CHECKBOXES["PVB_AI_OPENED"])

    pvb_cv = form.get("pvb_cv_result", "")
    if pvb_cv == "Leaked":    draw_x(c, *UNITED_CHECKBOXES["PVB_CV_LEAKED"])
    elif pvb_cv == "Held At": draw_x(c, *UNITED_CHECKBOXES["PVB_CV_HELD"])

    result = form.get("assembly_result", "")
    if result == "PASSED":   draw_x(c, *UNITED_CHECKBOXES["PASSED"])
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
    if rcl == "Yes":  draw_x(c, *JAX_CHECKBOXES["RECLAIM_YES"])
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
    if res_rcl == "Yes":  draw_x(c, *JAX_CHECKBOXES["RES_RECLAIM_YES"])
    elif res_rcl == "No": draw_x(c, *JAX_CHECKBOXES["RES_RECLAIM_NO"])

    # CV1 Initial — Closed Tight only (Leaked removed)
    icv1 = form.get("init_cv1_result", "")
    if icv1 == "Closed Tight":
        draw_x(c, *JAX_CHECKBOXES["INIT_CV1_CLOSED"])

    # CV2 Initial — Closed Tight only (Leaked removed)
    icv2 = form.get("init_cv2_result", "")
    if icv2 == "Closed Tight":
        draw_x(c, *JAX_CHECKBOXES["INIT_CV2_CLOSED"])

    # DP RV Initial: Opened At checkbox + psi text field
    irv = form.get("init_rv_result", "")
    if irv == "Opened At":      draw_x(c, *JAX_CHECKBOXES["INIT_RV_OPENED"])
    elif irv == "Did Not Open": draw_x(c, *JAX_CHECKBOXES["INIT_RV_DIDNOT"])
    # init_rv_psi is rendered via JAX_TEXT_FIELDS

    # Air Inlet Opened At checkbox + psi text field
    ipvb = form.get("init_pvb_result", "")
    if ipvb == "Air inlet opened at": draw_x(c, *JAX_CHECKBOXES["INIT_PVB_AIOPEN"])
    elif ipvb == "Did not open":      draw_x(c, *JAX_CHECKBOXES["INIT_PVB_AIDNOT"])
    # init_pvb_psi is rendered via JAX_TEXT_FIELDS

    fcv1 = form.get("final_cv1_result", "")
    if fcv1 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["FINAL_CV1_CLOSED"])

    fcv2 = form.get("final_cv2_result", "")
    if fcv2 == "Closed Tight": draw_x(c, *JAX_CHECKBOXES["FINAL_CV2_CLOSED"])

    frv = form.get("final_rv_result", "")
    if frv == "Opened At": draw_x(c, *JAX_CHECKBOXES["FINAL_RV_OPENED"])

    fpvb = form.get("final_pvb_result", "")
    if fpvb == "Satisfactory": draw_x(c, *JAX_CHECKBOXES["FINAL_PVB_SAT"])

    result = form.get("assembly_result", "")
    if result == "PASSED":   draw_x(c, *JAX_CHECKBOXES["JAX_PASSED"])
    elif result == "FAILED": draw_x(c, *JAX_CHECKBOXES["JAX_FAILED"])

    sig_ir = get_signature_image_reader()
    if sig_ir:
        c.drawImage(sig_ir, JAX_SIG_X, JAX_SIG_Y,
                    width=JAX_SIG_W, height=JAX_SIG_H, mask="auto")

    c.save()
    return _merge_overlay(TEMPLATE_JAX, overlay_buf)


# ---------------------------------------------------------------------------
# Delivery  — iOS-safe: single native Streamlit download_button only
# ---------------------------------------------------------------------------

def safe_filename(customer, street, location, prefix=""):
    def clean(s): return re.sub(r"[^\w\s\-]", "", s or "").strip()
    parts = [clean(customer) or "Customer",
             clean(street)   or "Address",
             clean(location) or "Location"]
    name = " - ".join(parts) + ".pdf"
    return (prefix + " " + name).strip() if prefix else name


def deliver_pdf(pdf_bytes: bytes, filename: str):
    """iOS/iPad-safe PDF delivery via native Streamlit download_button."""
    st.download_button(
        label="📥 Download PDF",
        data=pdf_bytes,
        file_name=filename,
        mime="application/pdf",
        use_container_width=True,
        key=f"pdf_dl_{filename}_{len(st.session_state.get('job_folder', []))}",
    )


# ---------------------------------------------------------------------------
# Tester defaults
# ---------------------------------------------------------------------------

TESTER_KEYS = ["gauge_mfg", "gauge_serial", "date_cal", "technician", "cert_no", "recert"]

def get_tester_defaults():
    return st.session_state.get("tester_defaults", {k: "" for k in TESTER_KEYS})

def save_tester_defaults(form):
    st.session_state["tester_defaults"] = {k: form.get(k, "") for k in TESTER_KEYS}


def _radio(label, options, key, form, **kwargs):
    opts = [""] + list(options)
    current = form.get(key, "")
    idx = opts.index(current) if current in opts else 0
    chosen = st.radio(label, opts, index=idx, key=key,
                      format_func=lambda x: "—" if x == "" else x, **kwargs)
    form[key] = chosen
    return chosen


# ===========================================================================
# App layout
# ===========================================================================

st.set_page_config(page_title="United Fire — Backflow Report", page_icon="🔧", layout="wide")
_load_sig_from_disk()
_init_job_folder()

# Sidebar: job folder always visible
render_job_folder_sidebar()

st.title("🔧 United Fire — Backflow Preventer Test Report")

# ---------------------------------------------------------------------------
# Form selector
# ---------------------------------------------------------------------------
form_choice = st.radio(
    "Select Form",
    ["United Fire (Standard)", "Jacksonville / JEA"],
    horizontal=True,
    key="form_choice",
)
st.divider()

# ---------------------------------------------------------------------------
# United Fire form
# ---------------------------------------------------------------------------
if form_choice == "United Fire (Standard)":

    if "united_form" not in st.session_state:
        defs = get_tester_defaults()
        f0 = {k: defs.get(k, "") for k in TESTER_KEYS}
        f0["date"]      = date.today().strftime("%m/%d/%Y")
        f0["test_date"] = date.today().strftime("%m/%d/%Y")
        st.session_state.united_form = f0

    f = st.session_state.united_form

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("➡️ Next Report (same job)"):
            kept = {k: f.get(k, "") for k in UNITED_NEXT_REPORT_KEEP}
            kept["date"]      = date.today().strftime("%m/%d/%Y")
            kept["test_date"] = date.today().strftime("%m/%d/%Y")
            save_tester_defaults(f)
            st.session_state.united_form = kept
            st.rerun()
    with col2:
        if st.button("🏢 New Job"):
            save_tester_defaults(f)
            defs = get_tester_defaults()
            f0 = {k: defs.get(k, "") for k in TESTER_KEYS}
            f0["date"]      = date.today().strftime("%m/%d/%Y")
            f0["test_date"] = date.today().strftime("%m/%d/%Y")
            st.session_state.united_form = f0
            st.session_state.job_folder = []
            st.rerun()
    with col3:
        if st.button("🗑️ Clear Form"):
            save_tester_defaults(f)
            defs = get_tester_defaults()
            f0 = {k: defs.get(k, "") for k in TESTER_KEYS}
            f0["date"]      = date.today().strftime("%m/%d/%Y")
            f0["test_date"] = date.today().strftime("%m/%d/%Y")
            st.session_state.united_form = f0
            st.rerun()

    st.divider()
    st.subheader("📋 Job Information")
    r1c1, r1c2, r1c3 = st.columns([1, 1, 2])
    f["date"]   = r1c1.text_input("Date",   f.get("date",   date.today().strftime("%m/%d/%Y")), key="u_date")
    f["branch"] = r1c2.text_input("Branch", f.get("branch", ""), key="u_branch")
    f["ahj"]    = r1c3.text_input("Authority Having Jurisdiction", f.get("ahj", ""), key="u_ahj")
    f["customer_name"]  = st.text_input("Customer / Site Name",  f.get("customer_name",  ""), key="u_cust")
    f["street_address"] = st.text_input("Street Address",         f.get("street_address", ""), key="u_addr")
    f["location"]       = st.text_input("Location of Assembly",   f.get("location",       ""), key="u_loc")

    st.divider()
    st.subheader("🔩 Backflow Assembly")
    c1, c2, c3, c4 = st.columns(4)
    f["serial_number"] = c1.text_input("Serial Number",  f.get("serial_number", ""), key="u_sn")
    f["manufacturer"]  = c2.text_input("Manufacturer ↺", f.get("manufacturer",  ""), key="u_mfg")
    f["model"]         = c3.text_input("Model ↺",         f.get("model",         ""), key="u_mdl")
    f["size"]          = c4.text_input("Size ↺",           f.get("size",          ""), key="u_sz")

    c1, c2, c3 = st.columns(3)
    asm_opts = ["", "RP", "DC", "PVB", "SVB"]
    f["assembly_type"] = c1.selectbox("Type of Assembly ↺", asm_opts,
        index=asm_opts.index(f.get("assembly_type","")) if f.get("assembly_type","") in asm_opts else 0,
        key="u_atype")
    ss_opts = ["", "FIRE", "DOMESTIC", "IRRIGATION", "ATTRACTION"]
    f["system_service"] = c2.selectbox("System Service ↺", ss_opts,
        index=ss_opts.index(f.get("system_service","")) if f.get("system_service","") in ss_opts else 0,
        key="u_ss")
    bp_opts = ["", "YES", "NO"]
    f["bypass"] = c3.selectbox("Bypass", bp_opts,
        index=bp_opts.index(f.get("bypass","")) if f.get("bypass","") in bp_opts else 0,
        key="u_bp")

    st.divider()
    st.subheader("🧪 Test Results")

    show_rp  = f.get("assembly_type") in ("", "RP", "DC")
    show_pvb = f.get("assembly_type") in ("", "PVB", "SVB")

    if show_rp:
        st.markdown("**Check Valves**")
        cv_c1, cv_c2 = st.columns(2)
        with cv_c1:
            st.markdown("*CV #1*")
            _radio("CV #1 Result", ["Closed Tight","Leaked"], "cv1_result", f, horizontal=True)
            f["cv1_dp"] = st.text_input("CV #1 Differential Pressure (psi)", f.get("cv1_dp",""), key="u_cv1dp")
        with cv_c2:
            st.markdown("*CV #2*")
            _radio("CV #2 Result", ["Closed Tight","Leaked"], "cv2_result", f, horizontal=True)
            f["cv2_dp"] = st.text_input("CV #2 Differential Pressure (psi)", f.get("cv2_dp",""), key="u_cv2dp")

        st.markdown("**Relief Valve**")
        rv_c1, rv_c2, rv_c3 = st.columns(3)
        with rv_c1:
            _radio("RV Result", ["Opened At","Did Not Open"], "rv_result", f, horizontal=True)
            f["rv_psi"] = st.text_input("RV Opened At (psi)", f.get("rv_psi",""), key="u_rvpsi")
        with rv_c2:
            _radio("RV Outlet", ["Closed","Leaked"], "rv_out_result", f, horizontal=True)
        with rv_c3:
            _radio("RV Inlet", ["Closed","Leaked"], "rv_in_result", f, horizontal=True)

    if show_pvb:
        st.markdown("**PVB**")
        pvb_c1, pvb_c2 = st.columns(2)
        with pvb_c1:
            _radio("Air Inlet", ["Closed Tight","Opened At"], "pvb_ai_result", f, horizontal=True)
            f["pvb_ai_psi"] = st.text_input("Air Inlet Opened At (psi)", f.get("pvb_ai_psi",""), key="u_aipsi")
        with pvb_c2:
            _radio("Check Valve", ["Leaked","Held At"], "pvb_cv_result", f, horizontal=True)
            f["pvb_cv_psi"] = st.text_input("CV Held/Leaked At (psi)", f.get("pvb_cv_psi",""), key="u_cvpsi")

    st.divider()
    f["test_date"] = st.text_input("Test Date", f.get("test_date", date.today().strftime("%m/%d/%Y")), key="u_tdate")

    res_opts = ["", "PASSED", "FAILED"]
    f["assembly_result"] = st.radio("Pass / Fail", res_opts,
        index=res_opts.index(f.get("assembly_result","")) if f.get("assembly_result","") in res_opts else 0,
        horizontal=True, format_func=lambda x: "—" if x=="" else x, key="u_res")

    st.divider()
    st.subheader("🔧 Repairs / Comments")
    f["repair_desc"] = st.text_area("Repair description", f.get("repair_desc",""), height=80, key="u_rep")

    st.divider()
    st.subheader("✍️ Signature")
    if st.session_state.get("signature_b64"):
        st.success("Signature on file ✓")
        if st.button("🗑️ Clear Signature", key="u_clrsig"):
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
                height=80, width=300, drawing_mode="freedraw", key="sig_canvas")
            if st.button("💾 Save Signature", key="u_savesig"):
                if sig_canvas.image_data is not None:
                    arr = sig_canvas.image_data
                    if arr.max() > 0:
                        save_signature(arr)
                        st.success("Signature saved!")
                        st.rerun()
                    else:
                        st.warning("Canvas is empty — draw your signature first.")
        except ImportError:
            st.warning("`pip install streamlit-drawable-canvas` to enable signature pad.")

    st.divider()

    with st.expander("🧰 Tester Info / Defaults", expanded=False):
        st.caption("Fill once — carries forward on Next Report and New Job.")
        t1, t2, t3 = st.columns(3)
        f["gauge_mfg"]    = t1.text_input("Gauge Manufacturer", f.get("gauge_mfg",""),    key="u_gmfg")
        f["gauge_serial"] = t2.text_input("Gauge Serial #",     f.get("gauge_serial",""), key="u_gsn")
        f["date_cal"]     = t3.text_input("Date Calibrated",    f.get("date_cal",""),     key="u_cal")
        t1b, t2b, t3b = st.columns(3)
        f["technician"] = t1b.text_input("Technician",        f.get("technician",""), key="u_tech")
        f["cert_no"]    = t2b.text_input("Certification No.", f.get("cert_no",""),    key="u_cert")
        f["recert"]     = t3b.text_input("Re-Cert Due Date",  f.get("recert",""),     key="u_recert")

    st.divider()

    if st.button("📄 Generate PDF", type="primary", use_container_width=True, key="u_gen"):
        with st.spinner("Building PDF..."):
            try:
                save_tester_defaults(f)
                pdf_bytes = generate_united_pdf(f)
                fname = safe_filename(f.get("customer_name",""), f.get("street_address",""), f.get("location",""))
                add_to_job_folder(pdf_bytes, fname)
                deliver_pdf(pdf_bytes, fname)
                st.success(f"✅ PDF ready: {fname}  |  Added to Job Folder ({len(st.session_state.job_folder)} total)")
            except Exception as e:
                st.error(f"Error generating PDF: {e}")

# ===========================================================================
# Jacksonville / JEA form
# ===========================================================================
else:

    if "jax_form" not in st.session_state:
        st.session_state.jax_form = {
            "init_test_date":  date.today().strftime("%m/%d/%Y"),
            "final_test_date": date.today().strftime("%m/%d/%Y"),
        }

    f = st.session_state.jax_form

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("➡️ Next Report (same job)", key="j_next"):
            kept = {k: f.get(k, "") for k in JAX_NEXT_REPORT_KEEP}
            kept["init_test_date"]  = date.today().strftime("%m/%d/%Y")
            kept["final_test_date"] = date.today().strftime("%m/%d/%Y")
            st.session_state.jax_form = kept
            st.rerun()
    with col2:
        if st.button("🏢 New Job", key="j_newjob"):
            kept = {k: f.get(k, "") for k in JAX_NEW_JOB_KEEP}
            kept["init_test_date"]  = date.today().strftime("%m/%d/%Y")
            kept["final_test_date"] = date.today().strftime("%m/%d/%Y")
            st.session_state.jax_form = kept
            st.session_state.job_folder = []
            st.rerun()
    with col3:
        if st.button("🗑️ Clear Form", key="j_clear"):
            st.session_state.jax_form = {
                "init_test_date":  date.today().strftime("%m/%d/%Y"),
                "final_test_date": date.today().strftime("%m/%d/%Y"),
            }
            st.rerun()

    st.divider()

    st.subheader("📋 Property & Contact Information")
    c1, c2 = st.columns(2)
    f["premises_name"]    = c1.text_input("Name of premises (company / person)", f.get("premises_name",""),    key="j_prem")
    f["owner_name"]       = c2.text_input("Owner or agent's name",               f.get("owner_name",""),       key="j_own")
    f["service_address"]  = c1.text_input("Service address",                      f.get("service_address",""),  key="j_sa")
    f["mailing_address"]  = c2.text_input("Mailing address",                      f.get("mailing_address",""),  key="j_ma")
    f["physical_location"]= c1.text_input("Physical location of device",          f.get("physical_location",""),key="j_pl")
    f["contact_phone"]    = c2.text_input("Contact phone number",                 f.get("contact_phone",""),    key="j_ph")
    f["jea_account"]      = c1.text_input("JEA account number",                   f.get("jea_account",""),      key="j_acct")
    f["meter_number"]     = c2.text_input("Meter number",                         f.get("meter_number",""),     key="j_meter")

    st.divider()

    st.subheader("📝 Test Purpose & Service Type")
    tp_opts_comm = ["", "Annual", "Repair", "Replacement", "New Installation"]
    st_opts_comm = ["", "Fire", "Irrigation", "Process/Isolation", "Potable", "Fire bypass"]
    rc_opts      = ["", "Yes", "No"]
    tp_opts_res  = ["", "Annual", "Repair", "Replacement", "New Installation"]
    st_opts_res  = ["", "Potable", "Irrigation / Is reclaimed"]

    cc1, cc2, cc3 = st.columns(3)
    f["comm_test_purpose"] = cc1.selectbox("Commercial test purpose", tp_opts_comm,
        index=tp_opts_comm.index(f.get("comm_test_purpose","")) if f.get("comm_test_purpose","") in tp_opts_comm else 0, key="j_ctp")
    f["comm_service_type"] = cc2.selectbox("Commercial service type", st_opts_comm,
        index=st_opts_comm.index(f.get("comm_service_type","")) if f.get("comm_service_type","") in st_opts_comm else 0, key="j_cst")
    f["comm_reclaim"] = cc3.selectbox("Reclaimed water? (Commercial)", rc_opts,
        index=rc_opts.index(f.get("comm_reclaim","")) if f.get("comm_reclaim","") in rc_opts else 0, key="j_crc")

    rc1, rc2, rc3 = st.columns(3)
    f["res_test_purpose"] = rc1.selectbox("Residential test purpose", tp_opts_res,
        index=tp_opts_res.index(f.get("res_test_purpose","")) if f.get("res_test_purpose","") in tp_opts_res else 0, key="j_rtp")
    f["res_service_type"] = rc2.selectbox("Residential service type", st_opts_res,
        index=st_opts_res.index(f.get("res_service_type","")) if f.get("res_service_type","") in st_opts_res else 0, key="j_rst")
    f["res_reclaim"] = rc3.selectbox("Reclaimed water? (Residential)", rc_opts,
        index=rc_opts.index(f.get("res_reclaim","")) if f.get("res_reclaim","") in rc_opts else 0, key="j_rrc")

    st.divider()

    st.subheader("🔩 Device Information")
    d1, d2, d3, d4, d5, d6 = st.columns(6)
    f["device_type"]    = d1.text_input("Device type",        f.get("device_type",""),    key="j_dt")
    f["manufacturer"]   = d2.text_input("Manufacturer",       f.get("manufacturer",""),   key="j_mfg")
    f["size"]           = d3.text_input("Size",               f.get("size",""),           key="j_sz")
    f["model_number"]   = d4.text_input("Model Number",       f.get("model_number",""),   key="j_mn")
    f["serial_number"]  = d5.text_input("Serial Number",      f.get("serial_number",""),  key="j_sn")
    f["install_date"]   = d6.text_input("Installation Date",  f.get("install_date",""),   key="j_id")

    st.divider()

    st.subheader("🧪 Initial Test")
    it1, it2, it3, it4 = st.columns(4)

    with it1:
        st.markdown("**Check Valve #1**")
        _radio("Result", ["Closed Tight"], "init_cv1_result", f, horizontal=True)
        f["init_cv1_psi"] = st.text_input("at ___ psi", f.get("init_cv1_psi",""), key="j_icv1p")

    with it2:
        st.markdown("**Check Valve #2**")
        _radio("Result", ["Closed Tight"], "init_cv2_result", f, horizontal=True)
        f["init_cv2_psi"] = st.text_input("at ___ psi", f.get("init_cv2_psi",""), key="j_icv2p")

    with it3:
        st.markdown("**DP RV Initial: Opened at**")
        _radio("Result", ["Opened At","Did Not Open"], "init_rv_result", f, horizontal=True)
        f["init_rv_psi"] = st.text_input("lbs reduced pressure", f.get("init_rv_psi",""), key="j_irvp")

    with it4:
        st.markdown("**Air Inlet Opened At**")
        _radio("Result", ["Air inlet opened at","Did not open"], "init_pvb_result", f, horizontal=True)
        f["init_pvb_psi"] = st.text_input("psi", f.get("init_pvb_psi",""), key="j_ipvbp")

    f["init_test_date"] = st.text_input("Initial Test Date", f.get("init_test_date", date.today().strftime("%m/%d/%Y")), key="j_itd")

    st.divider()

    st.subheader("✅ Final Test")
    ft1, ft2, ft3, ft4 = st.columns(4)

    with ft1:
        st.markdown("**Check Valve #1**")
        _radio("Result", ["Closed Tight"], "final_cv1_result", f, horizontal=True)
        f["final_cv1_psi"] = st.text_input("at ___ psi", f.get("final_cv1_psi",""), key="j_fcv1p")

    with ft2:
        st.markdown("**Check Valve #2**")
        _radio("Result", ["Closed Tight"], "final_cv2_result", f, horizontal=True)
        f["final_cv2_psi"] = st.text_input("at ___ psi", f.get("final_cv2_psi",""), key="j_fcv2p")

    with ft3:
        st.markdown("**Relief Valve**")
        _radio("Result", ["Opened At"], "final_rv_result", f, horizontal=True)
        f["final_rv_psi"] = st.text_input("lbs reduced pressure", f.get("final_rv_psi",""), key="j_frvp")

    with ft4:
        st.markdown("**PVB**")
        _radio("Result", ["Satisfactory"], "final_pvb_result", f, horizontal=True)

    f["final_test_date"] = st.text_input("Final Test Date", f.get("final_test_date", date.today().strftime("%m/%d/%Y")), key="j_ftd")

    res_opts = ["", "PASSED", "FAILED"]
    f["assembly_result"] = st.radio("Pass / Fail Certification", res_opts,
        index=res_opts.index(f.get("assembly_result","")) if f.get("assembly_result","") in res_opts else 0,
        horizontal=True, format_func=lambda x: "—" if x=="" else x, key="j_ares")

    st.divider()

    st.subheader("🔧 Repairs / Unusual Conditions")
    f["repairs"] = st.text_area("Repairs/unusual conditions", f.get("repairs",""), height=80, key="j_rep")

    st.divider()

    st.subheader("🖊️ Tester Information")
    ti1, ti2, ti3, ti4 = st.columns(4)
    f["init_tester_name"] = ti1.text_input("Initial tester name ↺", f.get("init_tester_name",""), key="j_itn")
    f["init_company"]     = ti2.text_input("Company ↺",              f.get("init_company",""),     key="j_ico")
    f["init_cert"]        = ti3.text_input("Cert # ↺",               f.get("init_cert",""),        key="j_ic")
    f["init_test_date"]   = ti4.text_input("Test date",              f.get("init_test_date", date.today().strftime("%m/%d/%Y")), key="j_itd2")

    ri1, ri2, ri3, ri4 = st.columns(4)
    f["repaired_by"]    = ri1.text_input("Repaired by",     f.get("repaired_by",""),    key="j_rb")
    f["repair_company"] = ri2.text_input("Repair company",  f.get("repair_company",""), key="j_rco")
    f["repair_cert"]    = ri3.text_input("Repair cert #",   f.get("repair_cert",""),    key="j_rc")
    f["repair_date"]    = ri4.text_input("Repair date",     f.get("repair_date",""),    key="j_rd")

    fi1, fi2, fi3, fi4 = st.columns(4)
    f["final_tester_name"] = fi1.text_input("Final tester name ↺", f.get("final_tester_name",""), key="j_ftn")
    f["final_company"]     = fi2.text_input("Company ↺",             f.get("final_company",""),     key="j_fco")
    f["final_cert"]        = fi3.text_input("Cert # ↺",              f.get("final_cert",""),        key="j_fc")
    f["final_test_date"]   = fi4.text_input("Test date",             f.get("final_test_date", date.today().strftime("%m/%d/%Y")), key="j_ftd2")
    f["signature_date"]    = st.text_input("Signature date",         f.get("signature_date",""),    key="j_sd")

    st.divider()
    st.subheader("✍️ Signature")
    if st.session_state.get("signature_b64"):
        st.success("Signature on file ✓")
        if st.button("🗑️ Clear Signature", key="j_clrsig"):
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
                height=80, width=300, drawing_mode="freedraw", key="j_sig_canvas")
            if st.button("💾 Save Signature", key="j_savesig"):
                if sig_canvas.image_data is not None:
                    arr = sig_canvas.image_data
                    if arr.max() > 0:
                        save_signature(arr)
                        st.success("Signature saved!")
                        st.rerun()
                    else:
                        st.warning("Canvas is empty — draw your signature first.")
        except ImportError:
            st.warning("`pip install streamlit-drawable-canvas` to enable signature pad.")

    st.divider()

    if st.button("📄 Generate Jacksonville PDF", type="primary", use_container_width=True, key="j_gen"):
        with st.spinner("Building PDF..."):
            try:
                pdf_bytes = generate_jax_pdf(f)
                fname = safe_filename(
                    f.get("premises_name",""),
                    f.get("service_address",""),
                    f.get("physical_location",""),
                    prefix="JAX"
                )
                add_to_job_folder(pdf_bytes, fname)
                deliver_pdf(pdf_bytes, fname)
                st.success(f"✅ PDF ready: {fname}  |  Added to Job Folder ({len(st.session_state.job_folder)} total)")
            except Exception as e:
                st.error(f"Error generating PDF: {e}")
