import streamlit as st
import pdfplumber
import openpyxl
import tempfile
import shutil
import os
import re
import json
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
    'E51110116': 'variable',   # Expense Refund = variable
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
    """Ensure list has 4 elements by repeating last value."""
    if not lst:
        return [0.0, 0.0, 0.0, 0.0]
    while len(lst) < 4:
        lst = lst + [lst[-1]]
    return lst[:4]

def extract_text_from_pdf(pdf_bytes):
    """Extract text from all pages of a PDF."""
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
        f.write(pdf_bytes)
        tmp_path = f.name
    try:
        text_pages = []
        with pdfplumber.open(tmp_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    text_pages.append(t)
        return '\n'.join(text_pages)
    finally:
        os.unlink(tmp_path)

def extract_company_name(text):
    """Extract client company name from contract text."""
    patterns = [
        r'hereinafter referred as the "Client"\.\s*\n.*?and\s+(.+?)\s+whose company',
        r'and\s+([\w\s\(\)\.]+(?:Limited|Co\.,?Ltd\.?|Public Company|PCL))\s+whose company',
        r'Between\s+PRTR Group Public Company Limited\s+And\s+(.+?)(?:\n|$)',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE | re.DOTALL)
        if m:
            name = m.group(1).strip()
            name = re.sub(r'\s+', ' ', name)
            if len(name) > 5:
                return name
    # Fallback: look for "And\n<Company Name>"
    m = re.search(r'And\s*\n\s*(.+(?:Limited|Ltd|PCL|Co\.))', text, re.IGNORECASE)
    if m:
        return m.group(1).strip()
    return "Unknown Company"

def parse_percentage(text):
    """Parse percentage string like '20%' or '8.5%' to float."""
    m = re.search(r'(\d+(?:\.\d+)?)\s*%', text)
    if m:
        return float(m.group(1)) / 100
    return None

def detect_fee_structure(text):
    """
    Detect if contract has flat rate or tiered-by-HC rate.
    Returns: ('flat', rates_dict) or ('tiered', rates_dict)
    """
    addendum_match = re.search(r'ADDENDUM OF THE CONTRACT 1.*', text, re.DOTALL | re.IGNORECASE)
    addendum_text = addendum_match.group(0) if addendum_match else text

    # Check for HC tier indicators
    tier_keywords = ['number of active contract employee', 'persons', 'hc']
    is_tiered = any(kw in addendum_text.lower() for kw in tier_keywords)

    rates = {}

    if is_tiered:
        # Extract tiered rates — find all % values near tier descriptions
        # Pattern: look for lines with "persons" and a % value
        tier_lines = re.findall(r'(?:Employee(?:s)?\s+(?:from\s+)?\d[\d\s,\-–]+(?:Persons?|onward|above)[^\n]*|'
                                r'\d[\d\s,\-–]+\s*(?:Persons?|HC)[^\n]*).*?(\d+(?:\.\d+)?)\s*%',
                                addendum_text, re.IGNORECASE | re.DOTALL)
        all_pcts = [float(x) / 100 for x in re.findall(r'(\d+(?:\.\d+)?)\s*%', addendum_text)]

        # Try to find fixed and variable income sections
        fixed_section = re.search(r'Fixed income.*?Variable income', addendum_text, re.DOTALL | re.IGNORECASE)
        variable_section = re.search(r'Variable income.*?(?:Expense|Social Security|$)', addendum_text, re.DOTALL | re.IGNORECASE)

        fixed_pcts = []
        if fixed_section:
            fixed_pcts = [float(x) / 100 for x in re.findall(r'(\d+(?:\.\d+)?)\s*%', fixed_section.group(0))]

        var_pcts = []
        if variable_section:
            var_pcts = [float(x) / 100 for x in re.findall(r'(\d+(?:\.\d+)?)\s*%', variable_section.group(0))]

        rates['fixed'] = pad4(fixed_pcts[:4]) if fixed_pcts else pad4([0.11, 0.10, 0.09, 0.085])
        rates['variable'] = pad4(var_pcts[:4]) if var_pcts else rates['fixed']

    else:
        # Flat rate — Recruit by PRTR vs by Client
        fixed_section = re.search(r'Fixed income.*?(?=Variable income|$)', addendum_text, re.DOTALL | re.IGNORECASE)
        variable_section = re.search(r'Variable income.*?(?=Expense|Social Security|$)', addendum_text, re.DOTALL | re.IGNORECASE)

        fixed_pcts = []
        if fixed_section:
            fixed_pcts = [float(x) / 100 for x in re.findall(r'(\d+(?:\.\d+)?)\s*%', fixed_section.group(0))]

        var_pcts = []
        if variable_section:
            var_pcts = [float(x) / 100 for x in re.findall(r'(\d+(?:\.\d+)?)\s*%', variable_section.group(0))]

        rates['fixed'] = fixed_pcts[:2] if len(fixed_pcts) >= 2 else [0.20, 0.15]
        rates['variable'] = var_pcts[:2] if var_pcts else [0.15, 0.15]

    # Common flat rates (same regardless of tier)
    def find_rate(pattern, default):
        m = re.search(pattern + r'.*?(\d+(?:\.\d+)?)\s*%', addendum_text, re.DOTALL | re.IGNORECASE)
        return float(m.group(1)) / 100 if m else default

    rates['sso'] = [0.0, 0.0, 0.0, 0.0]
    rates['pvd'] = pad4([find_rate(r'Provident Fund', 0.05)])
    rates['health_ins'] = pad4([find_rate(r'Health Insurance', 0.05)])
    rates['health_check'] = pad4([find_rate(r'Health Check', 0.05)])
    rates['uniform'] = pad4([find_rate(r'Uniform', 0.05)])
    rates['expense'] = pad4([find_rate(r'Expense.*?Employee', 0.10)])

    # Compensation rate from "Other Fees"
    comp_m = re.search(r'compensation.*?plus\s+(\d+(?:\.\d+)?)\s*%', addendum_text, re.IGNORECASE)
    comp_rate = float(comp_m.group(1)) / 100 if comp_m else 0.03
    rates['compensation'] = pad4([comp_rate])

    return ('tiered' if is_tiered else 'flat'), rates


def fill_excel(template_bytes, rates, is_tiered, company_name):
    """Fill fee table and return output bytes."""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        f.write(template_bytes)
        tmp_in = f.name

    tmp_out = tmp_in + '_out.xlsx'
    shutil.copy2(tmp_in, tmp_out)

    wb = openpyxl.load_workbook(tmp_out)
    ws = wb['3. Form Fee tabel 22.05.26']
    pct_fmt = '0.0%'

    for row in ws.iter_rows(min_row=4, max_row=ws.max_row):
        cost = row[8].value
        account = str(row[4].value).strip() if row[4].value else ''
        paycode = str(row[6].value).strip() if row[6].value else ''
        header = str(row[3].value).strip() if row[3].value else ''
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


# ─── UI ────────────────────────────────────────────────────────────────────
st.title("📊 PRTR Fee Table Filler")
st.caption("อัปโหลดสัญญา PDF + Master Fee Table แล้วระบบจะกรอก fee rate ให้อัตโนมัติ")

st.divider()

col1, col2 = st.columns(2)
with col1:
    pdf_file = st.file_uploader("📄 สัญญา PDF", type=['pdf'], help="Labor Supplier Contract (OUT-X-XXXX-XXXXX)")
with col2:
    xlsx_file = st.file_uploader("📋 Master Fee Table", type=['xlsx'], help="Master - Fee tebla.xlsx")

if pdf_file and xlsx_file:
    with st.spinner("กำลังอ่านสัญญา..."):
        pdf_text = extract_text_from_pdf(pdf_file.read())
        company_name = extract_company_name(pdf_text)
        struct_type, rates = detect_fee_structure(pdf_text)

    st.success(f"✅ อ่านสัญญาเสร็จแล้ว")

    st.subheader("📋 ข้อมูลที่อ่านได้จากสัญญา")
    st.write(f"**บริษัท Client:** {company_name}")
    st.write(f"**ประเภท Fee:** {'Tiered by HC จำนวน 4 ระดับ' if struct_type == 'tiered' else 'Flat Rate (Recruit by PRTR / Client)'}")

    # Show rate summary
    col_labels = ['J (Tier 1)', 'K (Tier 2)', 'L (Tier 3)', 'M (Tier 4)'] if struct_type == 'tiered' \
                 else ['J (Recruit by PRTR)', 'K (Recruit by Client)', 'L', 'M']

    rate_display = {
        'Fixed income': [f"{v*100:.1f}%" for v in pad4(rates.get('fixed', []))],
        'Variable income': [f"{v*100:.1f}%" for v in pad4(rates.get('variable', []))],
        'Expense Refund': [f"{v*100:.1f}%" for v in pad4(rates.get('variable', []))],
        'Reimbursement': [f"{v*100:.1f}%" for v in pad4(rates.get('expense', []))],
        'PVD': [f"{v*100:.1f}%" for v in pad4(rates.get('pvd', []))],
        'Health Insurance': [f"{v*100:.1f}%" for v in pad4(rates.get('health_ins', []))],
        'Severance/Compensation': [f"{v*100:.1f}%" for v in pad4(rates.get('compensation', []))],
    }

    import pandas as pd
    df = pd.DataFrame(rate_display, index=col_labels).T
    st.dataframe(df, use_container_width=True)

    st.info("⚠️ กรุณาตรวจสอบค่าด้านบนก่อนกด Generate — ถ้าไม่ถูกต้องสามารถแก้ไขได้ด้านล่าง")

    # Manual override
    with st.expander("✏️ แก้ไข fee rates (ถ้าจำเป็น)"):
        st.caption("ใส่เป็นตัวเลข % เช่น 20 หมายถึง 20%")
        for key, label in [('fixed','Fixed income'), ('variable','Variable income'),
                            ('expense','Expense/Reimbursement'), ('pvd','Provident Fund'),
                            ('health_ins','Health Insurance'), ('compensation','Severance/Compensation')]:
            current = rates.get(key, [0.0])
            cols = st.columns(4)
            new_vals = []
            for i, (c, lbl) in enumerate(zip(cols, col_labels)):
                v = current[i] if i < len(current) else current[-1]
                new_v = c.number_input(f"{label} - {lbl}", value=round(v*100, 2),
                                        min_value=0.0, max_value=100.0, step=0.5,
                                        key=f"{key}_{i}") / 100
                new_vals.append(new_v)
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
