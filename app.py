import streamlit as st
import streamlit.components.v1 as components
from io import BytesIO
from reportlab.pdfgen import canvas
import json, os, re, base64, tempfile
from datetime import date
from pdfrw import PdfReader, PdfWriter, PageMerge
from PIL import Image
import numpy as np

TEMPLATE_PATH = "backflow_template.pdf"
SESSION_FILE  = "session_data.json"
SIG_FILE      = "signature.png"
PAGE_W, PAGE_H = 612, 792

NEXT_REPORT_KEEP = {
    'branch', 'ahj', 'customer_name', 'street_address',
    'manufacturer', 'model', 'size', 'assembly_type', 'system_service',
    'gauge_mfg', 'gauge_serial', 'date_cal', 'technician', 'cert_no', 'recert',
}
NEW_JOB_KEEP = {
    'gauge_mfg', 'gauge_serial', 'date_cal', 'technician', 'cert_no', 'recert',
}
STATIC_TESTER_DEFAULTS = {
    'gauge_mfg': '', 'gauge_serial': '', 'date_cal': '',
    'technician': '', 'cert_no': '', 'recert': '',
}

TEXT_FIELDS = {
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
    "rv_psi":         (310, 398, 8),
    "cv1_dp":         (200, 321, 8),
    "cv2_dp":         (400, 312, 8),
    "pvb_ai_psi":     (510, 375, 8),
    "pvb_cv_psi":     (510, 320, 8),
    "test_date":      (165, 290, 8),
    "gauge_mfg":      (215, 178, 8),
    "gauge_serial":   (313, 178, 8),
    "date_cal":       (455, 178, 8),
    "technician":     (176, 150, 8),
    "cert_no":        (407, 165, 8),
    "recert":         (407, 150, 8),
}

# PSI field keys and their (x, y) text positions — used to draw red boxes
PSI_FIELDS = {
    "rv_psi":     (310, 398),
    "cv1_dp":     (200, 321),
    "cv2_dp":     (400, 312),
    "pvb_ai_psi": (510, 375),
    "pvb_cv_psi": (510, 320),
}

CHECKBOXES = {
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

# Repairs/remarks text box: (x, y, line_height, max_lines, wrap_width)
REPAIR_BOX = (75, 212, 10, 4, 65)


def draw_x(c, bx, by, size=3.8):
    c.setStrokeColorRGB(1, 0, 0)
    c.setLineWidth(2.0)
    c.line(bx-size, by-size, bx+size, by+size)
    c.line(bx+size, by-size, bx-size, by+size)


def draw_psi_box(c, x, y, width=30, height=12):
    """Draw a red rectangle around a PSI value field."""
    c.setStrokeColorRGB(1, 0, 0)
    c.setLineWidth(1.2)
    c.rect(x - 2, y - 3, width, height, stroke=1, fill=0)


def put_text(c, val, x, y, sz=8):
    if val:
        c.setFont("Helvetica-Bold", sz)
        c.setFillColorRGB(1, 0, 0)
        c.drawString(x, y, str(val))


def wrap_text(text, w=65):
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


def load_signature():
    if os.path.exists(SIG_FILE):
        with open(SIG_FILE, "rb") as fh:
            return BytesIO(fh.read())
    return None


def save_signature(img_array):
    img = Image.fromarray(img_array.astype("uint8"), "RGBA")
    img.save(SIG_FILE)


def generate_pdf(form):
    """Build overlay with ReportLab, merge onto template with pdfrw."""
    if not os.path.exists(TEMPLATE_PATH):
        st.error("⚠️ Place backflow_template.pdf in the same folder as app.py")
        st.stop()

    # 1. Build overlay
    overlay_buf = BytesIO()
    c = canvas.Canvas(overlay_buf, pagesize=(PAGE_W, PAGE_H))

    for field, (x, y, sz) in TEXT_FIELDS.items():
        put_text(c, form.get(field, ""), x, y, sz)

    # Draw red boxes around PSI fields (always draw boxes, text inside if value present)
    for field, (x, y) in PSI_FIELDS.items():
        draw_psi_box(c, x, y)

    for key in ["RP", "DC", "PVB", "SVB"]:
        if form.get("assembly_type") == key:
            draw_x(c, *CHECKBOXES[key])

    for key in ["FIRE", "DOMESTIC", "IRRIGATION", "ATTRACTION"]:
        if form.get("system_service") == key:
            draw_x(c, *CHECKBOXES[key])

    bp = form.get("bypass", "")
    if bp == "YES":  draw_x(c, *CHECKBOXES["BYPASS_YES"])
    elif bp == "NO": draw_x(c, *CHECKBOXES["BYPASS_NO"])

    cv1 = form.get("cv1_result", "")
    if cv1 == "Closed Tight": draw_x(c, *CHECKBOXES["CV1_CLOSED"])
    elif cv1 == "Leaked":     draw_x(c, *CHECKBOXES["CV1_LEAKED"])

    cv2 = form.get("cv2_result", "")
    if cv2 == "Closed Tight": draw_x(c, *CHECKBOXES["CV2_CLOSED"])
    elif cv2 == "Leaked":     draw_x(c, *CHECKBOXES["CV2_LEAKED"])

    rv = form.get("rv_result", "")
    if rv == "Opened At":      draw_x(c, *CHECKBOXES["RV_OPENED"])
    elif rv == "Did Not Open": draw_x(c, *CHECKBOXES["RV_DIDNOTOPEN"])

    rvo = form.get("rv_out_result", "")
    if rvo == "Closed":   draw_x(c, *CHECKBOXES["RV_OUT_CLOSED"])
    elif rvo == "Leaked": draw_x(c, *CHECKBOXES["RV_OUT_LEAKED"])

    rvi = form.get("rv_in_result", "")
    if rvi == "Closed":   draw_x(c, *CHECKBOXES["RV_IN_CLOSED"])
    elif rvi == "Leaked": draw_x(c, *CHECKBOXES["RV_IN_LEAKED"])

    pvb_ai = form.get("pvb_ai_result", "")
    if pvb_ai == "Closed Tight": draw_x(c, *CHECKBOXES["PVB_AI_CLOSED"])
    elif pvb_ai == "Opened At":  draw_x(c, *CHECKBOXES["PVB_AI_OPENED"])

    pvb_cv = form.get("pvb_cv_result", "")
    if pvb_cv == "Leaked":    draw_x(c, *CHECKBOXES["PVB_CV_LEAKED"])
    elif pvb_cv == "Held At": draw_x(c, *CHECKBOXES["PVB_CV_HELD"])

    result = form.get("assembly_result", "")
    if result == "PASSED":   draw_x(c, *CHECKBOXES["PASSED"])
    elif result == "FAILED": draw_x(c, *CHECKBOXES["FAILED"])

    # Repairs/remarks text — positioned at REPAIR_BOX
    rx, ry, rh, rmax, rw = REPAIR_BOX
    for i, ln in enumerate(wrap_text(form.get("repair_desc", ""), rw)[:rmax]):
        put_text(c, ln, rx, ry - i * rh, 7)

    sig_buf = load_signature()
    if sig_buf:
        c.drawImage(sig_buf, 170, 138, width=130, height=28, mask="auto")

    c.save()
    overlay_buf.seek(0)

    # 2. Merge onto template via pdfrw (must use temp file path)
    tp = PdfReader(TEMPLATE_PATH)
    op = PdfReader(overlay_buf)
    pg = tp.pages[0]
    PageMerge(pg).add(op.pages[0]).render()
    if pg.Annots:
        pg.Annots = []

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        PdfWriter().write(tmp_path, tp)
        with open(tmp_path, "rb") as fh:
            return fh.read()
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


def safe_filename(customer, street, location):
    def clean(s): return re.sub(r"[^\w\s\-]", "", s or "").strip()
    parts = [clean(customer) or "Customer",
             clean(street) or "Address",
             clean(location) or "Location"]
    return " - ".join(parts) + ".pdf"


def deliver_pdf(pdf_bytes: bytes, filename: str):
    """Cross-platform PDF delivery."""
    st.download_button(
        label="📥 Download PDF",
        data=pdf_bytes,
        file_name=filename,
        mime="application/pdf",
        use_container_width=True,
    )

    b64 = base64.b64encode(pdf_bytes).decode()
    html = f"""
    <html><body style="margin:0;padding:4px 0 0 0;font-family:sans-serif;">
      <a href="data:application/pdf;base64,{b64}"
         target="_blank"
         style="display:block;text-align:center;padding:12px;
                background:#5a5a5a;color:white;font-size:0.95rem;
                font-weight:bold;border-radius:8px;text-decoration:none;">
        &#127760; iOS / iPad: Tap here to open PDF
      </a>
      <p style="text-align:center;font-size:0.75rem;color:#666;margin-top:6px;">
        After PDF opens: tap Share &#128228; &rarr; Save to Files or Print.
      </p>
    </body></html>
    """
    components.html(html, height=90)


def load_session():
    data = {}
    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE) as fh:
                data = json.load(fh)
        except Exception:
            pass
    for k, v in STATIC_TESTER_DEFAULTS.items():
        data.setdefault(k, v)
    return data


def save_session(data):
    with open(SESSION_FILE, "w") as fh:
        json.dump(data, fh)


# ── Radio helper ─────────────────────────────────────────────────────────────
def _radio(label, options, key, form, **kwargs):
    opts = [""] + list(options)
    current = form.get(key, "")
    idx = opts.index(current) if current in opts else 0
    chosen = st.radio(label, opts, index=idx, key=key,
                      format_func=lambda x: "—" if x == "" else x, **kwargs)
    form[key] = chosen
    return chosen


# ── App ───────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="United Fire — Backflow Report", page_icon="🔧", layout="wide")
st.title("🔧 United Fire — Backflow Preventer Test Report")

if "form" not in st.session_state:
    st.session_state.form = load_session()

f = st.session_state.form

# ── Action bar ────────────────────────────────────────────────────────────────
col1, col2, col3, col4 = st.columns(4)
with col1:
    if st.button("💾 Save Session"):
        save_session(f)
        st.success("Saved!")
with col2:
    if st.button("📂 Load Session"):
        st.session_state.form = load_session()
        st.rerun()
with col3:
    if st.button("➡️ Next Report (same job)"):
        kept = {k: v for k, v in f.items() if k in NEXT_REPORT_KEEP}
        for k, v in STATIC_TESTER_DEFAULTS.items():
            kept.setdefault(k, f.get(k, v))
        kept["date"] = date.today().strftime("%m/%d/%Y")
        kept["test_date"] = date.today().strftime("%m/%d/%Y")
        st.session_state.form = kept
        st.rerun()
with col4:
    if st.button("🏢 New Job"):
        kept = {k: f.get(k, v) for k, v in STATIC_TESTER_DEFAULTS.items()}
        kept["date"] = date.today().strftime("%m/%d/%Y")
        kept["test_date"] = date.today().strftime("%m/%d/%Y")
        st.session_state.form = kept
        st.rerun()

st.divider()

# ── Job Information ───────────────────────────────────────────────────────────
st.subheader("📋 Job Information")
r1c1, r1c2, r1c3 = st.columns([1, 1, 2])
f["date"]   = r1c1.text_input("Date",   f.get("date",   date.today().strftime("%m/%d/%Y")))
f["branch"] = r1c2.text_input("Branch", f.get("branch", ""))
f["ahj"]    = r1c3.text_input("Authority Having Jurisdiction", f.get("ahj", ""))
f["customer_name"]  = st.text_input("Customer / Site Name",  f.get("customer_name",  ""))
f["street_address"] = st.text_input("Street Address",         f.get("street_address", ""))
f["location"]       = st.text_input("Location of Assembly",   f.get("location",       ""))

st.divider()

# ── Backflow Assembly ─────────────────────────────────────────────────────────
st.subheader("🔩 Backflow Assembly")
c1, c2, c3, c4 = st.columns(4)
f["serial_number"] = c1.text_input("Serial Number",      f.get("serial_number", ""))
f["manufacturer"]  = c2.text_input("Manufacturer ↺", f.get("manufacturer", ""))
f["model"]         = c3.text_input("Model ↺",         f.get("model", ""))
f["size"]          = c4.text_input("Size ↺",           f.get("size", ""))

c1, c2, c3 = st.columns(3)
asm_opts = ["", "RP", "DC", "PVB", "SVB"]
f["assembly_type"] = c1.selectbox(
    "Type of Assembly ↺", asm_opts,
    index=asm_opts.index(f.get("assembly_type", "")) if f.get("assembly_type", "") in asm_opts else 0,
)
ss_opts = ["", "FIRE", "DOMESTIC", "IRRIGATION", "ATTRACTION"]
f["system_service"] = c2.selectbox(
    "System Service ↺", ss_opts,
    index=ss_opts.index(f.get("system_service", "")) if f.get("system_service", "") in ss_opts else 0,
)
bp_opts = ["", "YES", "NO"]
f["bypass"] = c3.selectbox(
    "Bypass?", bp_opts,
    index=bp_opts.index(f.get("bypass", "")) if f.get("bypass", "") in bp_opts else 0,
)

st.divider()

# ── Testing Information ───────────────────────────────────────────────────────
st.subheader("🧪 Testing Information")
tc1, tc2, tc3, tc4 = st.columns(4)

with tc1:
    st.markdown("**Check Valve #1**")
    _radio("CV1 Result", ["Closed Tight", "Leaked"], "cv1_result", f, horizontal=True)
    f["cv1_dp"] = st.text_input("DP Across CV1 (PSI)", f.get("cv1_dp", ""), key="cv1dp")

with tc2:
    st.markdown("**Relief Valve**")
    _radio("RV Result", ["Opened At", "Did Not Open"], "rv_result", f, horizontal=True)
    f["rv_psi"] = st.text_input("RV PSI", f.get("rv_psi", ""), key="rvpsi")
    st.caption("Outlet Shut-Off")
    _radio("Outlet", ["Closed", "Leaked"], "rv_out_result", f, horizontal=True)
    st.caption("Inlet Shut-Off")
    _radio("Inlet", ["Closed", "Leaked"], "rv_in_result", f, horizontal=True)

with tc3:
    st.markdown("**Check Valve #2**")
    _radio("CV2 Result", ["Closed Tight", "Leaked"], "cv2_result", f, horizontal=True)
    f["cv2_dp"] = st.text_input("DP Across CV2 (PSI)", f.get("cv2_dp", ""), key="cv2dp")

with tc4:
    st.markdown("**PVB / SVB**")
    st.caption("Air Inlet")
    _radio("Air Inlet", ["Closed Tight", "Opened At"], "pvb_ai_result", f, horizontal=True)
    f["pvb_ai_psi"] = st.text_input("Air Inlet PSI", f.get("pvb_ai_psi", ""), key="pvbaipsi")
    st.caption("Check Valve")
    _radio("CV", ["Leaked", "Held At"], "pvb_cv_result", f, horizontal=True)
    f["pvb_cv_psi"] = st.text_input("CV PSI", f.get("pvb_cv_psi", ""), key="pvbcvpsi")

tc_l, tc_r = st.columns(2)
f["test_date"] = tc_l.text_input("Test Date", f.get("test_date", date.today().strftime("%m/%d/%Y")))
res_opts = ["", "PASSED", "FAILED"]
f["assembly_result"] = tc_r.radio(
    "This Assembly", res_opts,
    index=res_opts.index(f.get("assembly_result", "")) if f.get("assembly_result", "") in res_opts else 0,
    horizontal=True,
    format_func=lambda x: "—" if x == "" else x,
)

st.divider()

# ── Repairs & Remarks ─────────────────────────────────────────────────────────
st.subheader("🔧 Repairs & Remarks")
f["repair_desc"] = st.text_area(
    "Description of Repairs / Remarks (including Part #)",
    f.get("repair_desc", ""),
    height=100,
)

st.divider()

# ── Tester Info / Defaults ────────────────────────────────────────────────────
with st.expander("🧰 Tester Info / Defaults", expanded=False):
    st.caption("These values stay saved and auto-populate each new form.")
    t1, t2, t3 = st.columns(3)
    f["gauge_mfg"]    = t1.text_input("Gauge Manufacturer", f.get("gauge_mfg", ""))
    f["gauge_serial"] = t2.text_input("Gauge Serial #",     f.get("gauge_serial", ""))
    f["date_cal"]     = t3.text_input("Date Calibrated",    f.get("date_cal", ""))
    t1b, t2b, t3b = st.columns(3)
    f["technician"] = t1b.text_input("Technician",        f.get("technician", ""))
    f["cert_no"]    = t2b.text_input("Certification No.", f.get("cert_no", ""))
    f["recert"]     = t3b.text_input("Re-Cert Due Date",  f.get("recert", ""))

    st.markdown("---")
    st.markdown("**✍️ Digital Signature**")
    sig_exists = os.path.exists(SIG_FILE)
    if sig_exists:
        st.success("✅ Signature saved — stamps every PDF automatically.")
        with open(SIG_FILE, "rb") as sf:
            st.image(sf.read(), width=200)
        if st.button("🗑️ Clear Saved Signature"):
            os.remove(SIG_FILE)
            st.rerun()
    else:
        st.info("Draw your signature below then click Save Signature.")

    try:
        from streamlit_drawable_canvas import st_canvas
        sig_canvas = st_canvas(
            fill_color="rgba(0,0,0,0)",
            stroke_width=2,
            stroke_color="#cc0000",
            background_color="#ffffff",
            height=80,
            width=300,
            drawing_mode="freedraw",
            key="sig_canvas",
        )
        if st.button("💾 Save Signature"):
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

# ── Generate PDF ──────────────────────────────────────────────────────────────
if st.button("📄 Generate PDF", type="primary", use_container_width=True):
    with st.spinner("Building PDF..."):
        try:
            pdf_bytes = generate_pdf(f)
            fname = safe_filename(
                f.get("customer_name", ""),
                f.get("street_address", ""),
                f.get("location", ""),
            )
            save_session(f)
            deliver_pdf(pdf_bytes, fname)
            st.success(f"✅ PDF ready: {fname}")
        except Exception as e:
            st.error(f"Error generating PDF: {e}")
