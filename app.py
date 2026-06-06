import streamlit as st
import pdfplumber
import openpyxl
import tempfile
import shutil
import os
import re
from pathlib import Path

st.set_page_config(page_title="PRTR Fee Table Filler", page_icon="📊", layout="centered")

# ─── Account code → fee category ───────────────────────────────────────────
ACCOUNT_MAP = {
    'E51110101': 'fixed',  'E51110106': 'fixed',  'E51110123': 'fixed',
    'E51110102': 'sso',
    'E51110103': 'pvd',
    'E51110104': 'health_ins',
    'E51110105': 'variable','E51110108': 'variable','E51110109': 'variable',
    'E51110110': 'variable','E51110111': 'variable','E51110112': 'variable',
    'E51110127': 'variable',
    'E51110116': 'variable',   # Expense Refund = variable rate
    'E51110107': 'health_check',
    'E51110124': 'expense',    'E21710101': 'expense',
    'E51110125': 'uniform',
    'E51110126': 'compensation',
    'D51110101': 'fixed',  'D51110106': 'fixed',  'D51110123': 'fixed',
    'D51110104': 'health_ins',
    'D51110105': 'variable','D51110108': 'variable','D51110109': 'variable',
    'D51110110': 'variable','D51110111': 'variable','D51110112': 'variable',
    'D51110127': 'variable',
    'D51110124': 'expense',    'D21710101': 'expense',
    'D51110125': 'uniform',
}


def pad4(lst):
    if not lst:
        return [0.0, 0.0, 0.0, 0.0]
    lst = list(lst)
    while len(lst) < 4:
        lst.append(lst[-1])
    return lst[:4]


def extract_percentages(text):
    """Extract all % values from text (supports % and ร้อยละ formats)."""
    pcts = []
    for m in re.finditer(r'(\d+(?:\.\d+)?)\s*%', text):
        v = float(m.group(1))
        if 0 < v <= 100:
            pcts.append(v / 100)
    for m in re.finditer(r'ร้อยละ\s+(\d+(?:\.\d+)?)', text):
        v = float(m.group(1))
        if 0 < v <= 100:
            pcts.append(v / 100)
    return pcts


def extract_all_from_pdf(pdf_bytes):
    """Extract text + tables from all PDF pages."""
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
        f.write(pdf_bytes)
        tmp_path = f.name
    try:
        text_pages, all_tables = [], []
        with pdfplumber.open(tmp_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text_pages.append(t)
                tbls = page.extract_tables()
                if tbls:
                    all_tables.extend(tbls)
        return '\n'.join(text_pages), all_tables
    finally:
        os.unlink(tmp_path)


# ─── Company name extraction ────────────────────────────────────────────────
def extract_company_name(text, filename=''):
    patterns = [
        # Thai: บริษัท X จำกัด after "และ" line
        r'(?:และ|And)\s*\n\s*(บริษัท\s+.+?(?:จำกัด|มหาชน)(?:\s*\(มหาชน\))?)',
        # Thai inline after ลูกค้า
        r'(?:ลูกค้า|"ลูกค้า")\s*["\s]\s*(บริษัท\s+.+?(?:จำกัด|มหาชน))',
        # Thai: second party after PRTR block
        r'บริษัท\s+พีอาร์ทีอาร์.{5,200}?\n\s*(บริษัท\s+.+?(?:จำกัด|มหาชน))',
        # English standard
        r'hereinafter referred (?:as|to as) the ["\']Client["\']\.?\s*(?:\n.*?)?(?:and|And)\s+(.+?)\s+whose company',
        r'(?:and|And)\s+([\w\s\(\)\.&\-]+(?:Limited|Co\.,?\s*Ltd\.?|Public Company|PCL|Plc\.?))\s+whose company',
        r'Between\s+PRTR(?:\s+Group)?[^\n]+\s+(?:And|and)\s+(.+?)(?:\n|$)',
        r'(?:And|and)\s*\n\s*(.+?(?:Limited|Ltd\.?|PCL|Co\.))',
        r'"Client"\s+means\s+(.+?(?:Limited|Ltd\.?|PCL))',
    ]

    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE | re.DOTALL)
        if m:
            name = re.sub(r'\s+', ' ', m.group(1).strip()).rstrip('.,')
            if 5 < len(name) < 120 and 'PRTR' not in name.upper():
                return name

    # Fallback: parse filename
    if filename:
        clean = re.sub(r'_?OUT[-_]\d+[-_]\d+[-_]\d+[-_R\d]*', '', filename, flags=re.IGNORECASE)
        clean = re.sub(r'_?Completely[\s_]?Signed.*', '', clean, flags=re.IGNORECASE)
        clean = re.sub(r'\.(pdf)$', '', clean, flags=re.IGNORECASE)
        clean = clean.strip('_- ')
        if len(clean) > 3:
            return clean

    return "Unknown Company"


# ─── Rate helpers ────────────────────────────────────────────────────────────
def find_rate_near(text, *keywords, default=None):
    """Return first % found within ~120 chars after any keyword."""
    for kw in keywords:
        m = re.search(kw + r'[^\n%]{0,120}?(\d+(?:\.\d+)?)\s*%',
                      text, re.IGNORECASE | re.DOTALL)
        if m:
            return float(m.group(1)) / 100
        m = re.search(kw + r'[^\n]{0,120}?ร้อยละ\s+(\d+(?:\.\d+)?)',
                      text, re.IGNORECASE | re.DOTALL)
        if m:
            return float(m.group(1)) / 100
    return default


def extract_rates_from_tables(tables):
    """Extract fee rates from PDF table cells (bilingual)."""
    rates = {}
    keywords = {
        'fixed':        [r'fixed income', r'รายได้คงที่', r'เงินเดือน.*คงที่'],
        'variable':     [r'variable income', r'รายได้แปรผัน'],
        'pvd':          [r'provident fund', r'กองทุนสำรอง'],
        'health_ins':   [r'health insurance', r'ประกันสุขภาพ', r'ประกันชีวิต'],
        'health_check': [r'health check', r'ตรวจสุขภาพ'],
        'uniform':      [r'uniform', r'เครื่องแบบ'],
        'expense':      [r'reimbursement', r'expense.*employee', r'ค่าใช้จ่าย.*ประสาน'],
        'compensation': [r'compensation', r'severance', r'ค่าชดเชย', r'เงินชดเชย'],
    }
    for table in tables:
        if not table:
            continue
        for row in table:
            if not row:
                continue
            row_text = ' '.join(str(c or '') for c in row)
            for cat, kws in keywords.items():
                if cat in rates:
                    continue
                if any(re.search(kw, row_text, re.IGNORECASE) for kw in kws):
                    pcts = extract_percentages(row_text)
                    if pcts:
                        rates[cat] = pad4(pcts)
    return rates


# ─── Fee structure detection ────────────────────────────────────────────────
def detect_fee_structure(text, tables):
    """Detect structure type and extract all rates."""
    m = re.search(
        r'(?:ADDENDUM OF THE CONTRACT 1|บันทึกแนบท้ายสัญญา\s*1).*',
        text, re.DOTALL | re.IGNORECASE)
    add = m.group(0) if m else text

    # 1. Try table extraction
    rates = extract_rates_from_tables(tables)

    # 2. Detect structure type
    has_prtr_recruit  = bool(re.search(
        r'(?:สรรหา.{0,25}?พีอาร์ทีอาร์|Recruit by PRTR|โดย\s*พีอาร์ทีอาร์)', add, re.IGNORECASE))
    has_client_recruit = bool(re.search(
        r'(?:สรรหา.{0,25}?ลูกค้า|Recruit by (?:Client|Customer)|โดย\s*ลูกค้า)', add, re.IGNORECASE))
    has_hc_range = bool(re.search(
        r'(?:\d+\s*[-–]\s*\d+\s*(?:คน|Persons?)|ตั้งแต่\s*\d+\s*คน)', add, re.IGNORECASE))
    prtr_client_hc = has_prtr_recruit and has_client_recruit and has_hc_range

    tier_kws = ['number of active contract employee', 'persons', ' hc ',
                'จำนวนพนักงานที่ปฏิบัติงาน', 'จำนวน contract', 'เป็นต้นไป']
    is_tiered = any(kw in add.lower() for kw in tier_kws)

    # 3. Extract fixed/variable by structure
    if prtr_client_hc and 'fixed' not in rates:
        rates.update(_extract_prtr_client_hc(add))
        struct = 'tiered'
    elif is_tiered and 'fixed' not in rates:
        rates.update(_extract_tiered(add))
        struct = 'tiered'
    else:
        if 'fixed' not in rates:
            rates.update(_extract_flat(add))
        n = len(rates.get('fixed', []))
        struct = 'tiered' if n > 2 else 'flat'

    rates = _fill_common(rates, add)
    return struct, rates


def _extract_prtr_client_hc(text):
    """Iron Mountain style: PRTR/Client × HC range → 4 cols."""
    prtr_rates = [float(x) / 100 for x in re.findall(
        r'(?:สรรหา.{0,25}?พีอาร์ทีอาร์|Recruit by PRTR|โดย\s*พีอาร์ทีอาร์)[^\n%]{0,150}?(\d+(?:\.\d+)?)\s*%',
        text, re.IGNORECASE)]
    client_rates = [float(x) / 100 for x in re.findall(
        r'(?:สรรหา.{0,25}?ลูกค้า|Recruit by (?:Client|Customer)|โดย\s*ลูกค้า)[^\n%]{0,150}?(\d+(?:\.\d+)?)\s*%',
        text, re.IGNORECASE)]

    if prtr_rates and client_rates:
        p, c = prtr_rates[:2], client_rates[:2]
        fixed = [p[0], c[0], p[-1], c[-1]]
    else:
        fixed = [0.18, 0.08, 0.15, 0.07]
    return {'fixed': fixed}


def _extract_tiered(text):
    """DLK style: same rates split by HC tier → up to 4 cols."""
    fixed_sec = re.search(
        r'(?:Fixed income|รายได้คงที่).{0,600}?(?=Variable income|รายได้แปรผัน|$)',
        text, re.DOTALL | re.IGNORECASE)
    var_sec = re.search(
        r'(?:Variable income|รายได้แปรผัน).{0,600}?(?=Expense|Reimbursement|Social Security|กองทุน|$)',
        text, re.DOTALL | re.IGNORECASE)

    fp = extract_percentages(fixed_sec.group(0)) if fixed_sec else []
    vp = extract_percentages(var_sec.group(0)) if var_sec else []
    return {
        'fixed':    pad4(fp[:4]) if fp else [0.11, 0.10, 0.09, 0.085],
        'variable': pad4(vp[:4]) if vp else (pad4(fp[:4]) if fp else [0.11, 0.10, 0.09, 0.085]),
    }


def _extract_flat(text):
    """Flat rate: Recruit by PRTR vs Client → 2 cols."""
    fixed_sec = re.search(
        r'(?:Fixed income|รายได้คงที่).{0,400}?(?=Variable income|รายได้แปรผัน|$)',
        text, re.DOTALL | re.IGNORECASE)
    var_sec = re.search(
        r'(?:Variable income|รายได้แปรผัน).{0,400}?(?=Expense|Reimbursement|กองทุน|Social Security|$)',
        text, re.DOTALL | re.IGNORECASE)

    fp = extract_percentages(fixed_sec.group(0)) if fixed_sec else []
    vp = extract_percentages(var_sec.group(0)) if var_sec else []
    return {
        'fixed':    fp[:2] if len(fp) >= 2 else [0.20, 0.15],
        'variable': vp[:2] if vp else [0.15, 0.15],
    }


def _fill_common(rates, text):
    """Fill PVD, health, uniform, expense, compensation, SSO."""
    def get(keys, default):
        return find_rate_near(text, *keys, default=default)

    if 'pvd' not in rates:
        rates['pvd'] = pad4([get([r'Provident Fund', r'กองทุนสำรองเลี้ยงชีพ', r'กองทุนส\s*ำรอง'], 0.05)])
    if 'health_ins' not in rates:
        rates['health_ins'] = pad4([get([r'Health Insurance', r'ประกันสุขภาพ', r'ประกันชีวิต', r'การประกัน.*ซื่อสัตย์'], 0.05)])
    if 'health_check' not in rates:
        rates['health_check'] = pad4([get([r'Health Check', r'ตรวจสุขภาพ'], 0.05)])
    if 'uniform' not in rates:
        rates['uniform'] = pad4([get([r'Uniform', r'เครื่องแบบ'], 0.05)])
    if 'expense' not in rates:
        rates['expense'] = pad4([get([r'Reimbursement', r'Expense.*Employee', r'ค่าใช้จ่าย.*ประสาน'], 0.10)])
    if 'sso' not in rates:
        rates['sso'] = [0.0, 0.0, 0.0, 0.0]
    if 'compensation' not in rates:
        cm = re.search(
            r'(?:compensation|ค่าชดเชย|เงินชดเชย).{0,250}?(?:plus|บวก)\s+(\d+(?:\.\d+)?)\s*%',
            text, re.IGNORECASE | re.DOTALL)
        comp = float(cm.group(1)) / 100 if cm else 0.03
        rates['compensation'] = pad4([comp])

    for k in rates:
        if isinstance(rates[k], list):
            rates[k] = pad4(rates[k])
    return rates


# ─── Excel filler ───────────────────────────────────────────────────────────
def fill_excel(template_bytes, rates, is_tiered, company_name):
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        f.write(template_bytes)
        tmp_in = f.name
    tmp_out = tmp_in + '_out.xlsx'
    shutil.copy2(tmp_in, tmp_out)

    wb = openpyxl.load_workbook(tmp_out)
    ws = wb['3. Form Fee tabel 22.05.26']
    pct_fmt = '0.0%'

    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        cost    = row[8].value
        account = str(row[4].value).strip() if row[4].value else ''
        paycode = str(row[6].value).strip() if row[6].value else ''
        header  = str(row[3].value).strip() if row[3].value else ''
        j, k, l, m = row[9], row[10], row[11], row[12]

        if cost == 'PRTR':
            for cell in [j, k, l, m]:
                cell.value = 'PRTR'
        elif cost == 'CLNT':
            if paycode == '1 M NOTICE' or header == 'Severance Pay':
                category = 'compensation'
            else:
                category = ACCOUNT_MAP.get(account)
            if category:
                r = pad4(rates.get(category, [0.0]))
                for cell, val in zip([j, k, l, m], r):
                    cell.value = val
                    cell.number_format = pct_fmt
                if not is_tiered:
                    l.value = None
                    m.value = None

    wb.save(tmp_out)
    with open(tmp_out, 'rb') as f:
        out_bytes = f.read()
    os.unlink(tmp_in)
    os.unlink(tmp_out)
    return out_bytes


# ─── UI ─────────────────────────────────────────────────────────────────────
st.title("📊 PRTR Fee Table Filler")
st.caption("อัปโหลดสัญญา PDF + Master Fee Table แล้วระบบจะกรอก fee rate ให้อัตโนมัติ")
st.divider()

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📄 สัญญา PDF", type=['pdf'],
                                 help="Labor Supplier Contract (OUT-X-XXXX-XXXXX)")
with col2:
    xlsx_file = st.file_uploader("📋 Master Fee Table", type=['xlsx'],
                                  help="Master - Fee tebla.xlsx")

if pdf_file and xlsx_file:
    with st.spinner("กำลังอ่านสัญญา..."):
        pdf_bytes = pdf_file.read()
        full_text, tables = extract_all_from_pdf(pdf_bytes)
        company_name = extract_company_name(full_text, pdf_file.name)
        struct_type, rates = detect_fee_structure(full_text, tables)

    st.success("✅ อ่านสัญญาเสร็จแล้ว")
    st.subheader("📋 ข้อมูลที่อ่านได้จากสัญญา")
    st.write(f"**บริษัท Client:** {company_name}")
    st.write(f"**ประเภท Fee:** {'Tiered (4 อัตรา)' if struct_type == 'tiered' else 'Flat Rate (Recruit by PRTR / Client)'}")

    col_labels = ['J', 'K', 'L', 'M']
    import pandas as pd
    rate_display = {
        'Fixed income':           [f"{v*100:.1f}%" for v in rates.get('fixed', [0]*4)],
        'Variable income':        [f"{v*100:.1f}%" for v in rates.get('variable', [0]*4)],
        'Expense Refund':         [f"{v*100:.1f}%" for v in rates.get('variable', [0]*4)],
        'Reimbursement':          [f"{v*100:.1f}%" for v in rates.get('expense', [0]*4)],
        'PVD':                    [f"{v*100:.1f}%" for v in rates.get('pvd', [0]*4)],
        'Health Insurance':       [f"{v*100:.1f}%" for v in rates.get('health_ins', [0]*4)],
        'Severance/Compensation': [f"{v*100:.1f}%" for v in rates.get('compensation', [0]*4)],
    }
    df = pd.DataFrame(rate_display, index=col_labels).T
    st.dataframe(df, use_container_width=True)

    with st.expander("🔍 ดูข้อความที่อ่านจาก PDF (ใช้ตรวจสอบเมื่อค่าไม่ถูกต้อง)"):
        addendum_match = re.search(
            r'(?:ADDENDUM OF THE CONTRACT 1|บันทึกแนบท้ายสัญญา\s*1).*',
            full_text, re.DOTALL | re.IGNORECASE)
        snippet = addendum_match.group(0)[:3000] if addendum_match else full_text[:3000]
        st.text(snippet)

    st.info("⚠️ กรุณาตรวจสอบค่าด้านบนก่อนกด Generate — ถ้าไม่ถูกต้องสามารถแก้ไขได้ด้านล่าง")

    with st.expander("✏️ แก้ไข fee rates (ถ้าจำเป็น)"):
        st.caption("ใส่เป็นตัวเลข % เช่น 20 หมายถึง 20%")
        for key, label in [('fixed','Fixed income'), ('variable','Variable income'),
                            ('expense','Expense/Reimbursement'), ('pvd','Provident Fund'),
                            ('health_ins','Health Insurance'), ('compensation','Severance/Compensation')]:
            current = pad4(rates.get(key, [0.0]))
            cols = st.columns(4)
            new_vals = []
            for i, (c, lbl) in enumerate(zip(cols, col_labels)):
                v = current[i] if i < len(current) else current[-1]
                nv = c.number_input(f"{label} - {lbl}", value=round(v*100, 2),
                                    min_value=0.0, max_value=100.0, step=0.5,
                                    key=f"{key}_{i}") / 100
                new_vals.append(nv)
            rates[key] = new_vals

    st.divider()
    if st.button("🚀 Generate Fee Table", type="primary", use_container_width=True):
        with st.spinner("กำลังสร้างไฟล์..."):
            xlsx_file.seek(0)
            out_bytes = fill_excel(xlsx_file.read(), rates, struct_type == 'tiered', company_name)
            output_filename = f"Master - Fee tebla_{company_name}.xlsx"
        st.success("✅ สร้างไฟล์เสร็จแล้ว!")
        st.download_button(
            label=f"⬇️ ดาวน์โหลด {output_filename}",
            data=out_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    st.info("👆 กรุณาอัปโหลดทั้ง 2 ไฟล์ด้านบนเพื่อเริ่มต้น")

st.divider()
st.caption("PRTR Group Public Company Limited · Fee Table Automation Tool")
