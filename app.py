import streamlit as st
from io import BytesIO
from reportlab.pdfgen import canvas
import json, os, re
from datetime import date
from pdfrw import PdfReader, PdfWriter, PageMerge

TEMPLATE_PATH = "backflow_template.pdf"
SESSION_FILE  = "session_data.json"
PAGE_W, PAGE_H = 612, 792

STICKY = {
    'branch', 'ahj', 'customer_name', 'street_address', 'manufacturer', 'model', 'size',
    'gauge_mfg', 'gauge_serial', 'date_cal', 'technician', 'cert_no', 'recert', 'technician_signature'
}

STATIC_TESTER_DEFAULTS = {
    'gauge_mfg': '',
    'gauge_serial': '',
    'date_cal': '',
    'technician': '',
    'cert_no': '',
    'recert': '',
    'technician_signature': ''
}

# Calibrated from PPTX text box positions mapped to PDF 612x792 coordinate space
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
    "rv_psi":         (290, 398, 8),
    "cv1_dp":         (180, 321, 8),
    "cv2_dp":         (380, 312, 8),
    "pvb_ai_psi":     (490, 375, 8),
    "pvb_cv_psi":     (490, 320, 8),
    "test_date":      (165, 290, 8),
    "gauge_mfg":      (215, 178, 8),
    "gauge_serial":   (313, 178, 8),
    "date_cal":       (455, 178, 8),
    "technician":     (176, 150, 8),
    "cert_no":        (407, 165, 8),
    "recert":         (407, 150, 8),
}

CHECKBOXES = {
    "RP": (210,460), "DC": (270,460), "PVB": (210,441), "SVB": (270,441),
    "FIRE": (395,490), "DOMESTIC": (395,475), "IRRIGATION": (395,460), "ATTRACTION": (395,442),
    "BYPASS_YES": (500,470), "BYPASS_NO": (500,450),
    "CV1_CLOSED": (130,390), "CV1_LEAKED": (130,375),
    "CV2_CLOSED": (330,390), "CV2_LEAKED": (330,375),
    "PVB_AI_CLOSED": (425,405), "PVB_AI_OPENED": (425,370),
    "PVB_CV_LEAKED": (425,355), "PVB_CV_HELD": (425,325),
    "RV_OPENED": (225,398), "RV_DIDNOTOPEN": (225,378),
    "RV_OUT_CLOSED": (225,334), "RV_OUT_LEAKED": (272,334),
    "RV_IN_CLOSED": (225,310), "RV_IN_LEAKED": (272,310),
    "PASSED": (400,290), "FAILED": (462,290),
}

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

def wrap_text(text, w=38):
    words = text.split(); lines, line = [], ""
    for word in words:
        t = (line + " " + word).strip()
        if len(t) > w: lines.append(line.strip()); line = word
        else: line = t
    if line: lines.append(line)
    return lines

def draw_signature_line(c, x1, x2, y):
    c.setStrokeColorRGB(1, 0, 0)
    c.setLineWidth(1.3)
    c.line(x1, y, x2, y)


def generate_pdf(form):
    if not os.path.exists(TEMPLATE_PATH):
        st.error(f"⚠️ Place **backflow_template.pdf** in the same folder as app.py")
        st.stop()

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))

    for field, (x, y, sz) in TEXT_FIELDS.items():
        put_text(c, form.get(field, ""), x, y, sz)

    asm = form.get("assembly_type", "")
    for key in ["RP", "DC", "PVB", "SVB"]:
        if asm == key: draw_x(c, *CHECKBOXES[key])

    ss = form.get("system_service", "")
    for key in ["FIRE", "DOMESTIC", "IRRIGATION", "ATTRACTION"]:
        if ss == key: draw_x(c, *CHECKBOXES[key])

    bp = form.get("bypass", "")
    if bp == "YES": draw_x(c, *CHECKBOXES["BYPASS_YES"])
    elif bp == "NO": draw_x(c, *CHECKBOXES["BYPASS_NO"])

    for k in ["CV1_CLOSED", "CV1_LEAKED"]:
        if form.get(k.lower()): draw_x(c, *CHECKBOXES[k])

    for k in ["CV2_CLOSED", "CV2_LEAKED"]:
        if form.get(k.lower()): draw_x(c, *CHECKBOXES[k])

    for k in ["PVB_AI_CLOSED", "PVB_AI_OPENED", "PVB_CV_LEAKED", "PVB_CV_HELD"]:
        if form.get(k.lower()): draw_x(c, *CHECKBOXES[k])

    for k in ["RV_OPENED", "RV_DIDNOTOPEN", "RV_OUT_CLOSED", "RV_OUT_LEAKED", "RV_IN_CLOSED", "RV_IN_LEAKED"]:
        if form.get(k.lower()): draw_x(c, *CHECKBOXES[k])

    result = form.get("assembly_result", "")
    if result == "PASSED": draw_x(c, *CHECKBOXES["PASSED"])
    elif result == "FAILED": draw_x(c, *CHECKBOXES["FAILED"])

    for i, ln in enumerate(wrap_text(form.get("repair_desc", ""), 38)[:4]):
        put_text(c, ln, 96, 200 - i * 10, 7)

    draw_signature_line(c, 170, 340, 151)
    put_text(c, form.get("technician_signature", ""), 176, 153, 8)

    c.save(); buf.seek(0)

    tp = PdfReader(TEMPLATE_PATH)
    op = PdfReader(buf)
    pg = tp.pages[0]
    PageMerge(pg).add(op.pages[0]).render()
    if pg.Annots: pg.Annots = []
    out = BytesIO()
    PdfWriter().write(out, tp)
    out.seek(0)
    return out

def safe_filename(customer, location):
    def clean(s): return re.sub(r"[^\w\s\-]", "", s).strip()
    return f"{clean(customer) or 'Customer'} - {clean(location) or 'location'}.pdf"

def load_session():
    data = {}
    if os.path.exists(SESSION_FILE):
        try:
            with open(SESSION_FILE) as f:
                data = json.load(f)
        except:
            pass
    for k, v in STATIC_TESTER_DEFAULTS.items():
        data.setdefault(k, v)
    return data

def save_session(data):
    with open(SESSION_FILE, "w") as f: json.dump(data, f)

st.set_page_config(page_title="United Fire — Backflow Report", page_icon="🔧", layout="wide")
st.title("🔧 United Fire — Backflow Preventer Test Report")

if "form" not in st.session_state:
    st.session_state.form = load_session()

f = st.session_state.form

col1, col2, col3 = st.columns(3)
with col1:
    if st.button("💾 Save Session"):
        save_session(f); st.success("Saved!")
with col2:
    if st.button("📂 Load Session"):
        st.session_state.form = load_session(); st.rerun()
with col3:
    if st.button("🗑️ Clear (keep ⭐ sticky fields)"):
        st.session_state.form = {k: v for k, v in f.items() if k in STICKY}
        for k, v in STATIC_TESTER_DEFAULTS.items():
            st.session_state.form.setdefault(k, f.get(k, v))
        st.rerun()

st.divider()

st.subheader("📋 Job Information")
r1c1, r1c2, r1c3 = st.columns([1, 1, 2])
f["date"]   = r1c1.text_input("Date",  f.get("date", date.today().strftime("%m/%d/%Y")))
f["branch"] = r1c2.text_input("Branch ⭐", f.get("branch", ""))
f["ahj"]    = r1c3.text_input("Authority Having Jurisdiction ⭐", f.get("ahj", ""))
f["customer_name"]  = st.text_input("Customer / Site Name",  f.get("customer_name", ""))
f["street_address"] = st.text_input("Street Address",         f.get("street_address", ""))
f["location"]       = st.text_input("Location of Assembly",   f.get("location", ""))

st.divider()

st.subheader("🔩 Backflow Assembly")
c1, c2, c3, c4 = st.columns(4)
f["serial_number"] = c1.text_input("Serial Number",    f.get("serial_number", ""))
f["manufacturer"]  = c2.text_input("Manufacturer ⭐", f.get("manufacturer", ""))
f["model"]         = c3.text_input("Model ⭐",          f.get("model", ""))
f["size"]          = c4.text_input("Size ⭐",           f.get("size", ""))

c1, c2, c3 = st.columns(3)
asm_opts = ["", "RP", "DC", "PVB", "SVB"]
f["assembly_type"] = c1.selectbox("Type of Assembly",
    asm_opts, index=asm_opts.index(f.get("assembly_type","")) if f.get("assembly_type","") in asm_opts else 0)

ss_opts = ["", "FIRE", "DOMESTIC", "IRRIGATION", "ATTRACTION"]
f["system_service"] = c2.selectbox("System Service",
    ss_opts, index=ss_opts.index(f.get("system_service","")) if f.get("system_service","") in ss_opts else 0)

bp_opts = ["", "YES", "NO"]
f["bypass"] = c3.selectbox("Bypass?",
    bp_opts, index=bp_opts.index(f.get("bypass","")) if f.get("bypass","") in bp_opts else 0)

st.divider()

st.subheader("🧪 Testing Information")
tc1, tc2, tc3, tc4 = st.columns(4)

with tc1:
    st.markdown("**Check Valve #1**")
    f["cv1_closed"] = st.checkbox("Closed Tight", f.get("cv1_closed", False), key="cv1c")
    f["cv1_leaked"] = st.checkbox("Leaked",       f.get("cv1_leaked", False), key="cv1l")
    f["cv1_dp"]     = st.text_input("DP Across CV1 (PSI)", f.get("cv1_dp", ""), key="cv1dp")

with tc2:
    st.markdown("**Relief Valve**")
    f["rv_opened"]     = st.checkbox("Opened At",    f.get("rv_opened",     False), key="rvo")
    f["rv_psi"]        = st.text_input("RV PSI",      f.get("rv_psi", ""),          key="rvpsi")
    f["rv_didnotopen"] = st.checkbox("Did Not Open", f.get("rv_didnotopen", False), key="rvdno")
    st.markdown("*Outlet Shut-Off*")
    f["rv_out_closed"] = st.checkbox("Closed",  f.get("rv_out_closed", False), key="rvoc")
    f["rv_out_leaked"] = st.checkbox("Leaked",  f.get("rv_out_leaked", False), key="rvol")
    st.markdown("*Inlet Shut-Off*")
    f["rv_in_closed"]  = st.checkbox("Closed",  f.get("rv_in_closed",  False), key="rvic")
    f["rv_in_leaked"]  = st.checkbox("Leaked",  f.get("rv_in_leaked",  False), key="rvil")

with tc3:
    st.markdown("**Check Valve #2**")
    f["cv2_closed"] = st.checkbox("Closed Tight", f.get("cv2_closed", False), key="cv2c")
    f["cv2_leaked"] = st.checkbox("Leaked",       f.get("cv2_leaked", False), key="cv2l")
    f["cv2_dp"]     = st.text_input("DP Across CV2 (PSI)", f.get("cv2_dp", ""), key="cv2dp")

with tc4:
    st.markdown("**PVB / SVB**")
    st.caption("Air Inlet")
    f["pvb_ai_closed"] = st.checkbox("Closed Tight", f.get("pvb_ai_closed", False), key="pvbaic")
    f["pvb_ai_opened"] = st.checkbox("Opened At",    f.get("pvb_ai_opened", False), key="pvbaio")
    f["pvb_ai_psi"]    = st.text_input("Air Inlet PSI", f.get("pvb_ai_psi", ""),   key="pvbaipsi")
    st.caption("Check Valve")
    f["pvb_cv_leaked"] = st.checkbox("Leaked",   f.get("pvb_cv_leaked", False), key="pvbcvl")
    f["pvb_cv_held"]   = st.checkbox("Held At",  f.get("pvb_cv_held",   False), key="pvbcvh")
    f["pvb_cv_psi"]    = st.text_input("CV PSI",  f.get("pvb_cv_psi",   ""),   key="pvbcvpsi")

tc_l, tc_r = st.columns(2)
f["test_date"] = tc_l.text_input("Test Date", f.get("test_date", date.today().strftime("%m/%d/%Y")))
res_opts = ["", "PASSED", "FAILED"]
f["assembly_result"] = tc_r.radio("This Assembly",
    res_opts,
    index=res_opts.index(f.get("assembly_result","")) if f.get("assembly_result","") in res_opts else 0,
    horizontal=True)

st.divider()

st.subheader("🔧 Repairs & Remarks")
f["repair_desc"] = st.text_area("Description of Repairs / Remarks (including Part #)",
    f.get("repair_desc", ""), height=100)

st.divider()

with st.expander("🧰 Tester Info / Defaults", expanded=False):
    st.caption("These values stay saved and auto-populate each new form. Open only when you need to update them.")
    t1, t2, t3 = st.columns(3)
    f["gauge_mfg"]    = t1.text_input("Gauge Manufacturer ⭐", f.get("gauge_mfg", ""))
    f["gauge_serial"] = t2.text_input("Gauge Serial # ⭐",     f.get("gauge_serial", ""))
    f["date_cal"]     = t3.text_input("Date Calibrated ⭐",    f.get("date_cal", ""))
    t1b, t2b, t3b = st.columns(3)
    f["technician"] = t1b.text_input("Technician ⭐",         f.get("technician", ""))
    f["cert_no"]    = t2b.text_input("Certification No. ⭐",  f.get("cert_no", ""))
    f["recert"]     = t3b.text_input("Re-Cert Due Date ⭐",   f.get("recert", ""))
    f["technician_signature"] = st.text_input("Technician Signature Text ⭐", f.get("technician_signature", ""))

st.divider()

if st.button("📄 Generate PDF", type="primary", use_container_width=True):
    with st.spinner("Building PDF..."):
        try:
            pdf_bytes = generate_pdf(f)
            fname = safe_filename(f.get("customer_name","Customer"), f.get("street_address","Address"))
            st.download_button(f"⬇️ Download: {fname}", data=pdf_bytes,
                file_name=fname, mime="application/pdf", use_container_width=True)
            save_session(f)
            st.success(f"✅ {fname}")
        except Exception as e:
            st.error(f"Error: {e}")
