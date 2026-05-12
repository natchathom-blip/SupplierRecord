import streamlit as st
import io, os, json, smtplib, copy
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import openpyxl
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter
from reportlab.lib.utils import simpleSplit

# ─── page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="แบบสอบถาม CPRAM – ผักสลัด", layout="wide")

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pdf")
EXCEL_PATH = os.path.join(os.path.dirname(__file__), "data_cpram.xlsx")

# ─── Thai font setup ────────────────────────────────────────────────────────────
THAI_FONT = None
_local = os.path.join(os.path.dirname(__file__), "fonts", "THSarabunNew.ttf")

if os.path.exists(_local):
    pdfmetrics.registerFont(TTFont("ThaiFont", _local))
    THAI_FONT = "ThaiFont"
else:
    import glob
    for pattern in ["C:/Windows/Fonts/**/*.ttf", "/usr/share/fonts/**/*.ttf"]:
        for h in glob.glob(pattern, recursive=True):
            if any(k in h.lower() for k in ["sarabun","thai","garuda","cordia"]):
                try:
                    pdfmetrics.registerFont(TTFont("ThaiFont", h))
                    THAI_FONT = "ThaiFont"
                    break
                except: continue
        if THAI_FONT: break

FONT_NAME = THAI_FONT if THAI_FONT else "Helvetica"


# ─── CSS ────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { font-family: 'Sarabun', sans-serif; }
    .section-header {
        color: #2e7d32;
        font-size: 1.15rem;
        font-weight: 700;
        border-bottom: 2px solid #2e7d32;
        padding-bottom: 6px;
        margin-bottom: 16px;
        margin-top: 8px;
    }
    .item-card {
        background: #f1f8e9;
        border: 1px solid #a5d6a7;
        border-radius: 10px;
        padding: 16px 20px 8px;
        margin-bottom: 18px;
    }
    .item-title {
        font-weight: 700;
        font-size: 1rem;
        color: #1b5e20;
        margin-bottom: 12px;
    }
    .stButton>button {
        background: #2e7d32;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
    }
    .stButton>button:hover { background: #1b5e20; }
    .delete-btn>button {
        background: #c62828 !important;
        font-size: 0.85rem;
        padding: 0.3rem 0.9rem;
    }
    .add-btn>button {
        background: white !important;
        color: #2e7d32 !important;
        border: 2px dashed #2e7d32 !important;
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# ─── session state ──────────────────────────────────────────────────────────────
if "item_list" not in st.session_state:
    st.session_state.item_list = [{}]

def add_item():
    st.session_state.item_list.append({})

def remove_item(idx):
    if len(st.session_state.item_list) > 1:
        st.session_state.item_list.pop(idx)


# ─── Thai Data with Zipcode from JSON ──────────────────────────────────────────
@st.cache_data
def load_geo_data():
    file_path = os.path.join(os.path.dirname(__file__), "thailand_geo_with_zip.json")
    with open(file_path, encoding="utf-8") as f:
        return json.load(f)

GEO_DATA = load_geo_data()

def get_provinces(country):
    if country == "ไทย" and GEO_DATA:
        return sorted(list(GEO_DATA.keys()))
    return []

def get_districts(country, province):
    if country == "ไทย" and GEO_DATA and province in GEO_DATA:
        return sorted(list(GEO_DATA[province].keys()))
    return []

def get_subdistricts(country, province, district):
    if country == "ไทย" and GEO_DATA and province in GEO_DATA and district in GEO_DATA[province]:
        return sorted(list(GEO_DATA[province][district].keys()))
    return []

def get_zipcode(province, district, subdistrict):
    try:
        return GEO_DATA[province][district][subdistrict]
    except KeyError:
        return ""


# ─── EXCEL helper ───────────────────────────────────────────────────────────────
def save_to_excel(form_data: dict):
    headers = [
        "วันที่บันทึก", "ผู้ส่งมอบ", "วันที่ส่งวัตถุดิบ", "เวลาส่ง",
        "อีเมล", "ลงชื่อผู้กรอก", "ประเทศแหล่งปลูก(default)",
        "ชนิดวัตถุดิบ", "Code", "จำนวน(KG)",
        "วันที่เก็บเกี่ยว", "เวลาเก็บเกี่ยว", "วันที่ล้างความสะอาด",
        "เวลาล้างความสะอาด", "ชื่อผู้ปลูก", "เลขที่GAP",
        "รหัสไร่", "ที่อยู่เลขที่", "หมู่ที่",
        "ประเทศ", "จังหวัด", "อำเภอ/เมือง", "ตำบล/เขต",
        "รหัสไปรษณีย์", "สายพันธุ์", "ลักษณะการปลูก", "ลักษณะสถานที่ปลูก",
    ]
    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "CPRAM Data"
        ws.append(headers)
        wb.save(EXCEL_PATH)

    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active

    for item in form_data.get("items", []):
        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            form_data.get("supplier", ""),
            form_data.get("delivery_date", ""),
            form_data.get("delivery_time", ""),
            form_data.get("email", ""),
            form_data.get("signer", ""),
            form_data.get("default_country", ""),
            item.get("material_type", ""),
            item.get("code", ""),
            item.get("quantity", ""),
            item.get("harvest_date", ""),
            item.get("harvest_time", ""),
            item.get("clean_date", ""),
            item.get("clean_time", ""),
            item.get("grower_name", ""),
            item.get("gap_number", ""),
            item.get("farm_code", ""),
            item.get("address_no", ""),
            item.get("moo", ""),
            item.get("country", ""),
            item.get("province", ""),
            item.get("district", ""),
            item.get("subdistrict", ""),
            item.get("postal_code", ""),
            item.get("variety", ""),
            item.get("growing_type", ""),
            item.get("growing_condition", ""),
        ]
        ws.append(row)
    wb.save(EXCEL_PATH)


# ─── PDF generation ─────────────────────────────────────────────────────────────
def draw_text(c, text, x, y, font_size=9, max_width=None):
    if not text:
        return
    c.setFont(FONT_NAME, font_size)
    if max_width:
        lines = simpleSplit(str(text), FONT_NAME, font_size, max_width)
        for i, line in enumerate(lines):
            c.drawString(x, y - i * (font_size + 2), line)
    else:
        c.drawString(x, y, str(text))


def generate_pdf(form_data: dict) -> bytes:
    reader = PdfReader(TEMPLATE_PATH)
    writer = PdfWriter()

    supplier      = form_data.get("supplier", "")
    delivery_date = form_data.get("delivery_date", "")
    delivery_time = form_data.get("delivery_time", "")
    signer        = form_data.get("signer", "")
    items         = form_data.get("items", [])

    W, H = 595.32, 841.92

    overlay_buf = io.BytesIO()
    c = canvas.Canvas(overlay_buf, pagesize=(W, H))
    c.setFillColorRGB(0, 0, 0)

    fs = 14

    draw_text(c, supplier,      220, 728, fs, max_width=230)
    draw_text(c, delivery_date, 465, 728, fs, max_width=100)

    item = items[0] if items else {}

    draw_text(c, item.get("material_type",""), 260, 696, fs, max_width=120)
    draw_text(c, item.get("code",""),          390, 696, fs, max_width=80)
    draw_text(c, item.get("quantity",""),      490, 696, fs, max_width=60)

    draw_text(c, item.get("variety",""),           145, 667, fs, max_width=130)
    draw_text(c, item.get("growing_type",""),      310, 667, fs, max_width=130)
    draw_text(c, item.get("growing_condition",""), 470, 667, fs, max_width=90)

    draw_text(c, item.get("harvest_date",""), 210, 638, fs, max_width=150)
    draw_text(c, item.get("harvest_time",""), 420, 638, fs, max_width=130)

    draw_text(c, item.get("clean_date",""), 190, 609, fs, max_width=150)
    draw_text(c, item.get("clean_time",""), 420, 609, fs, max_width=130)

    draw_text(c, item.get("grower_name",""), 160, 580, fs, max_width=330)

    draw_text(c, item.get("gap_number",""), 155, 551, fs, max_width=200)
    draw_text(c, item.get("farm_code",""),  370, 551, fs, max_width=120)

    addr = " ".join(filter(None, [
        item.get("address_no",""),
        f"หมู่ {item.get('moo','')}" if item.get("moo","") else "",
        item.get("subdistrict",""),
        item.get("district",""),
        item.get("province",""),
        item.get("postal_code",""),
    ]))
    draw_text(c, addr, 115, 522, fs, max_width=440)

    draw_text(c, signer,        420, 62, fs)
    draw_text(c, delivery_date, 420, 42, fs)

    c.save()
    overlay_buf.seek(0)

    overlay_reader = PdfReader(overlay_buf)
    page1 = copy.copy(reader.pages[0])
    page1.merge_page(overlay_reader.pages[0])
    writer.add_page(page1)

    for page in reader.pages[1:]:
        extra_buf = io.BytesIO()
        ec = canvas.Canvas(extra_buf, pagesize=(W, H))
        ec.setFillColorRGB(0, 0, 0)
        draw_text(ec, signer,        420, 62, fs)
        draw_text(ec, delivery_date, 420, 42, fs)
        ec.save()
        extra_buf.seek(0)
        extra_reader = PdfReader(extra_buf)
        p = copy.copy(page)
        p.merge_page(extra_reader.pages[0])
        writer.add_page(p)

    out_buf = io.BytesIO()
    writer.write(out_buf)
    out_buf.seek(0)
    return out_buf.read()


# ─── Email send ─────────────────────────────────────────────────────────────────
def send_email(to_addr: str, pdf_bytes: bytes, smtp_cfg: dict) -> bool:
    try:
        msg = MIMEMultipart()
        msg["From"]    = smtp_cfg["user"]
        msg["To"]      = to_addr
        msg["Subject"] = "แบบสอบถามประจำวัน CPRAM – ผักสลัด"
        msg.attach(MIMEText("กรุณาดูเอกสารแบบสอบถามที่แนบมาด้วยครับ/ค่ะ", "plain", "utf-8"))
        part = MIMEBase("application", "octet-stream")
        part.set_payload(pdf_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", 'attachment; filename="cpram_questionnaire.pdf"')
        msg.attach(part)
        with smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["port"]) as server:
            server.login(smtp_cfg["user"], smtp_cfg["password"])
            server.sendmail(smtp_cfg["user"], to_addr, msg.as_string())
        return True
    except Exception as e:
        st.warning(f"ส่งอีเมลไม่สำเร็จ: {e}")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("<h2 style='color:#1b5e20;text-align:center;'>📋 แบบสอบถามประจำวัน CPRAM – ผักสลัด</h2>", unsafe_allow_html=True)

# ── ส่วนที่ 1 ──────────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">ส่วนที่ 1 — ข้อมูลผู้ส่งมอบและการส่งมอบ</div>', unsafe_allow_html=True)

col1, col2, col3 = st.columns([2, 2, 2])
with col1:
    supplier = st.text_input("ผู้ส่งมอบ (Supplier) *", key="supplier")
with col2:
    delivery_date = st.date_input("วันที่ส่งวัตถุดิบ *", key="delivery_date", value=None)
with col3:
    delivery_time = st.time_input("เวลาส่ง", key="delivery_time", value=None, step=60)

col4, col5, col6 = st.columns([2, 2, 2])
with col4:
    email = st.text_input("อีเมลของผู้ส่งมอบ (สำหรับรับ PDF) *", key="email")
with col5:
    signer = st.text_input("ลงชื่อผู้กรอก", key="signer")
with col6:
    default_country = st.selectbox("ประเทศแหล่งปลูก (default)", ["ประเทศไทย", "จีน", "ญี่ปุ่น", "เกาหลี", "อื่นๆ"], key="default_country")

st.markdown("---")

# ── ส่วนที่ 2 ──────────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">ส่วนที่ 2 — รายการวัตถุดิบ <span style="font-size:0.85rem;font-weight:400;">(เพิ่มได้ไม่จำกัด)</span></div>', unsafe_allow_html=True)

GROWING_TYPES = ["- เลือก -", "ปลูกดินยกพื้น", "ปลูกดินไม่ยกพื้น", "ปลูกไฮโดรโปนิกส์"]
GROWING_CONDITIONS = ["- เลือก -", "ระบบเปิด", "ระบบปิด"]

items_data = []
for i, _ in enumerate(st.session_state.item_list):
    st.markdown(f'<div class="item-card"><div class="item-title">รายการที่ {i+1}</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns([3, 2, 2])
    with c1:
        mat = st.text_input("ชนิดวัตถุดิบที่ส่งให้ทาง CPRAM *", key=f"mat_{i}")
    with c2:
        code = st.text_input("Code (เช่น 71000277)", key=f"code_{i}")
    with c3:
        qty = st.text_input("จำนวน (KG) *", key=f"qty_{i}")

    c4, c5, c6 = st.columns([2, 2, 2])
    with c4:
        h_date = st.date_input("วันที่เก็บเกี่ยว", key=f"hdate_{i}", value=None)
    with c5:
        h_time = st.time_input("เวลาเก็บเกี่ยว", key=f"htime_{i}", value=None, step=60)
    with c6:
        c_date = st.date_input("วันที่ล้างทำความสะอาด", key=f"cdate_{i}", value=None)

    c7, c8, c9 = st.columns([2, 2, 2])
    with c7:
        c_time = st.time_input("เวลาล้างทำความสะอาด", key=f"ctime_{i}", value=None, step=60)
    with c8:
        grower = st.text_input("ชื่อผู้ปลูก", key=f"grower_{i}")
    with c9:
        gap = st.text_input("เลขที่ GAP", key=f"gap_{i}")

    c10, c11, c12 = st.columns([2, 2, 2])
    with c10:
        farm_code = st.text_input("รหัสไร่", key=f"fcode_{i}")
    with c11:
        addr_no = st.text_input("ที่อยู่เลขที่", key=f"addr_{i}")
    with c12:
        moo = st.text_input("หมู่ที่", key=f"moo_{i}")

    st.markdown("📍 **ที่อยู่แหล่งปลูก**")
    cc1, cc2, cc3, cc4 = st.columns(4)
    with cc1:
        country_opts = ["ไทย"]
        country = st.selectbox("ประเทศ", country_opts, key=f"cntry_{i}")
    with cc2:
        prov_opts = ["- เลือก -"] + get_provinces(country)
        province = st.selectbox("จังหวัด/มณฑล", prov_opts, key=f"prov_{i}")
    with cc3:
        dist_opts = ["- เลือกจังหวัดก่อน -"]
        if province != "- เลือก -":
            dist_opts = ["- เลือก -"] + get_districts(country, province)
        district = st.selectbox("อำเภอ/เมือง", dist_opts, key=f"dist_{i}")
    with cc4:
        sub_opts = ["- เลือกอำเภอก่อน -"]
        if district not in ("- เลือก -", "- เลือกจังหวัดก่อน -"):
            subs = get_subdistricts(country, province, district)
            sub_opts = ["- เลือก -"] + subs
        subdistrict = st.selectbox("ตำบล/เขต", sub_opts, key=f"sub_{i}")

    auto_zip = ""
    if country == "ไทย" and subdistrict not in ("- เลือก -", "- เลือกอำเภอก่อน -"):
        auto_zip = get_zipcode(province, district, subdistrict)

    cc5, cc6, cc7, cc8 = st.columns(4)
    with cc5:
        st.text_input("รหัสไปรษณีย์", value=auto_zip, key=f"postal_display_{i}", disabled=True)
        postal = auto_zip
    with cc6:
        variety = st.text_input("สายพันธุ์", key=f"variety_{i}")
    with cc7:
        g_type = st.selectbox("ลักษณะการปลูก", GROWING_TYPES, key=f"gtype_{i}")
    with cc8:
        g_cond = st.selectbox("ลักษณะสถานที่ปลูก", GROWING_CONDITIONS, key=f"gcond_{i}")

    if len(st.session_state.item_list) > 1:
        _, del_col = st.columns([6, 1])
        with del_col:
            st.markdown('<div class="delete-btn">', unsafe_allow_html=True)
            if st.button("✕ ลบรายการนี้", key=f"del_{i}"):
                remove_item(i)
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

    items_data.append({
        "material_type": mat,
        "code": code,
        "quantity": qty,
        "harvest_date": h_date.strftime("%d/%m/%Y") if h_date else "",
        "harvest_time": h_time,
        "clean_date": c_date.strftime("%d/%m/%Y") if c_date else "",
        "clean_time": c_time,
        "grower_name": grower,
        "gap_number": gap,
        "farm_code": farm_code,
        "address_no": addr_no,
        "moo": moo,
        "country": country,
        "province": province if province != "- เลือก -" else "",
        "district": district if district not in ("- เลือก -","- เลือกจังหวัดก่อน -") else "",
        "subdistrict": subdistrict if subdistrict not in ("- เลือก -","- เลือกอำเภอก่อน -") else "",
        "postal_code": postal,
        "variety": variety,
        "growing_type": g_type if g_type != "- เลือก -" else "",
        "growing_condition": g_cond if g_cond != "- เลือก -" else "",
    })

st.markdown('<div class="add-btn">', unsafe_allow_html=True)
if st.button("＋ เพิ่มรายการวัตถุดิบ"):
    add_item()
    st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")

with st.expander("⚙️ ตั้งค่า SMTP (สำหรับส่งอีเมล – ไม่บังคับ)"):
    s1, s2, s3 = st.columns(3)
    with s1:
        smtp_host = st.text_input("SMTP Host", value="smtp.gmail.com", key="smtp_host")
        smtp_port = st.number_input("Port", value=465, key="smtp_port")
    with s2:
        smtp_user = st.text_input("Gmail/Email ผู้ส่ง", key="smtp_user")
        smtp_pass = st.text_input("Password / App Password", type="password", key="smtp_pass")
    with s3:
        st.info("💡 Gmail: ใช้ App Password (2FA ต้องเปิด)\nไปที่ myaccount.google.com → Security → App passwords")

st.markdown("")
if st.button("💾 บันทึกและส่ง PDF", use_container_width=True):
    errors = []
    if not supplier:       errors.append("กรุณากรอก ผู้ส่งมอบ")
    if not delivery_date:  errors.append("กรุณาเลือก วันที่ส่งวัตถุดิบ")
    if not delivery_time:  errors.append("กรุณากรอก เวลาส่ง")
    if not email:          errors.append("กรุณากรอก อีเมล")
    if not signer:         errors.append("กรุณากรอก ลงชื่อผู้กรอก")
    for idx, it in enumerate(items_data):
        n = idx + 1
        if not it["material_type"]:     errors.append(f"รายการที่ {n}: กรุณากรอก ชนิดวัตถุดิบ")
        if not it["code"]:              errors.append(f"รายการที่ {n}: กรุณากรอก Code")
        if not it["quantity"]:          errors.append(f"รายการที่ {n}: กรุณากรอก จำนวน")
        if not it["harvest_date"]:      errors.append(f"รายการที่ {n}: กรุณาเลือก วันที่เก็บเกี่ยว")
        if not it["harvest_time"]:      errors.append(f"รายการที่ {n}: กรุณากรอก เวลาเก็บเกี่ยว")
        if not it["clean_date"]:        errors.append(f"รายการที่ {n}: กรุณาเลือก วันที่ล้าง")
        if not it["clean_time"]:        errors.append(f"รายการที่ {n}: กรุณากรอก เวลาล้าง")
        if not it["grower_name"]:       errors.append(f"รายการที่ {n}: กรุณากรอก ชื่อผู้ปลูก")
        if not it["gap_number"]:        errors.append(f"รายการที่ {n}: กรุณากรอก เลขที่ GAP")
        if not it["farm_code"]:         errors.append(f"รายการที่ {n}: กรุณากรอก รหัสไร่")
        if not it["address_no"]:        errors.append(f"รายการที่ {n}: กรุณากรอก ที่อยู่เลขที่")
        if not it["province"]:          errors.append(f"รายการที่ {n}: กรุณาเลือก จังหวัด")
        if not it["district"]:          errors.append(f"รายการที่ {n}: กรุณาเลือก อำเภอ")
        if not it["subdistrict"]:       errors.append(f"รายการที่ {n}: กรุณาเลือก ตำบล")
        if not it["variety"]:           errors.append(f"รายการที่ {n}: กรุณากรอก สายพันธุ์")
        if not it["growing_type"]:      errors.append(f"รายการที่ {n}: กรุณาเลือก ลักษณะการปลูก")
        if not it["growing_condition"]: errors.append(f"รายการที่ {n}: กรุณาเลือก ลักษณะสถานที่ปลูก")

    if errors:
        for e in errors:
            st.error(e)
    else:
        form_data = {
            "supplier":        supplier,
            "delivery_date":   delivery_date.strftime("%d/%m/%Y") if delivery_date else "",
            "delivery_time":   delivery_time,
            "email":           email,
            "signer":          signer,
            "default_country": default_country,
            "items":           items_data,
        }

        with st.spinner("กำลังสร้าง PDF และบันทึกข้อมูล..."):
            try:
                save_to_excel(form_data)
                st.success("✅ บันทึกข้อมูลลง Excel เรียบร้อยแล้ว")
            except Exception as ex:
                st.warning(f"บันทึก Excel ไม่สำเร็จ: {ex}")

            try:
                pdf_bytes = generate_pdf(form_data)
                st.success("✅ สร้าง PDF เรียบร้อยแล้ว")
            except Exception as ex:
                st.error(f"สร้าง PDF ไม่สำเร็จ: {ex}")
                pdf_bytes = None

            if pdf_bytes:
                smtp_cfg = {
                    "host":     st.session_state.get("smtp_host", "smtp.gmail.com"),
                    "port":     int(st.session_state.get("smtp_port", 465)),
                    "user":     st.session_state.get("smtp_user", ""),
                    "password": st.session_state.get("smtp_pass", ""),
                }
                sent = False
                if smtp_cfg["user"] and smtp_cfg["password"]:
                    sent = send_email(email, pdf_bytes, smtp_cfg)
                    if sent:
                        st.success(f"📧 ส่ง PDF ไปยัง {email} เรียบร้อยแล้ว!")

                if not sent:
                    st.info("📥 ดาวน์โหลด PDF ด้านล่าง (ยังไม่ได้ตั้งค่า SMTP หรือส่งอีเมลไม่สำเร็จ)")
                    st.download_button(
                        label="⬇️ ดาวน์โหลด PDF",
                        data=pdf_bytes,
                        file_name=f"cpram_{supplier}_{delivery_date}.pdf",
                        mime="application/pdf",
                    )

if os.path.exists(EXCEL_PATH):
    with open(EXCEL_PATH, "rb") as f:
        st.download_button(
            label="📊 ดาวน์โหลดไฟล์ Excel (ข้อมูลทั้งหมด)",
            data=f.read(),
            file_name="data_cpram.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
