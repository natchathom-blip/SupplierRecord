# CPRAM ผักสลัด – แบบสอบถามประจำวัน

## วิธีติดตั้งและรัน

### 1. ติดตั้ง dependencies
```bash
pip install -r requirements.txt
```

### 2. ติดตั้ง Thai font (Linux)
```bash
# Ubuntu/Debian
sudo apt-get install fonts-thai-tlwg

# หรือ
sudo apt-get install fonts-tlwg-garuda
```

### 3. รัน Streamlit
```bash
streamlit run app.py
```

เปิดเบราว์เซอร์ที่ `http://localhost:8501`

---

## ฟีเจอร์

- ✅ **ส่วนที่ 1** – ข้อมูลผู้ส่งมอบ (Supplier, วันที่, เวลา, อีเมล)
- ✅ **ส่วนที่ 2** – รายการวัตถุดิบ เพิ่ม/ลบได้ไม่จำกัด
- ✅ **Cascading Dropdown** – ประเทศ → จังหวัด → อำเภอ → ตำบล
- ✅ **PDF output** – overlay ข้อมูลลงบน template `ผักxxl.pdf`
- ✅ **Excel** – บันทึกทุก submission ลงไฟล์ `data_cpram.xlsx` เดิม
- ✅ **ส่งอีเมล** – ตั้งค่า SMTP ใน expander แล้วส่ง PDF อัตโนมัติ
- ✅ **ดาวน์โหลด PDF** – ถ้าไม่ได้ตั้ง SMTP ก็ดาวน์โหลดลงเครื่องได้

---

## ตั้งค่าอีเมล (Gmail)

1. เปิด 2-Step Verification ที่ myaccount.google.com
2. ไปที่ Security → App passwords → สร้าง password ใหม่
3. กรอก App Password (16 ตัวอักษร) ในช่อง Password/App Password

---

## ไฟล์ในโปรเจกต์

| ไฟล์ | คำอธิบาย |
|------|----------|
| `app.py` | Streamlit application หลัก |
| `template.pdf` | PDF template ของ CPRAM |
| `requirements.txt` | Python dependencies |
| `data_cpram.xlsx` | ไฟล์ Excel (สร้างอัตโนมัติเมื่อ submit ครั้งแรก) |
