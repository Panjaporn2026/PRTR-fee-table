[README.md](https://github.com/user-attachments/files/28663120/README.md)
# PRTR Fee Table Filler — Web App

## วิธีรันในเครื่อง (Local)

```bash
# 1. ติดตั้ง dependencies
pip install -r requirements.txt

# 2. รัน app
streamlit run app.py
```
เปิด browser ที่ http://localhost:8501

---

## วิธี Deploy บน Streamlit Community Cloud (ฟรี)

1. สร้าง GitHub repository และอัปโหลดไฟล์ใน folder นี้ทั้งหมด
2. ไปที่ https://share.streamlit.io
3. Sign in ด้วย GitHub
4. คลิก **"New app"** → เลือก repo → เลือก `app.py` → คลิก **Deploy**
5. ได้ลิงค์ `https://your-app-name.streamlit.app` ทันที

**ข้อจำกัด Streamlit Cloud ฟรี:**
- ใช้งานได้ไม่จำกัด user
- PDF ที่เป็นรูปภาพ (scan) อาจอ่านไม่ได้ครบ (ไม่มี OCR)
- ไฟล์ไม่ถูกเก็บบน server — ดาวน์โหลดทันทีหลัง generate

---

## โครงสร้างไฟล์

```
prtr-fee-app/
├── app.py           ← หลัก Streamlit app
├── requirements.txt ← Python dependencies
└── README.md        ← คู่มือนี้
```

---

## หมายเหตุ

- PDF ที่เป็น text (ไม่ใช่ scan) จะอ่าน fee rate ได้อัตโนมัติ
- PDF ที่เป็น scan → ระบบอาจอ่านค่าไม่ได้ → ใช้ manual override ก่อน generate
- รองรับทั้ง Flat Rate และ Tiered by HC
