import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import re

pdf_path = "/Users/pratikkharat/Downloads/project/docID-387619268-9036ffd802f2fc1a12ee7f114b68fa0a1dc0b2529d69589083a5e475ae366b2d_Shivani Mahore Bank Statement 23-24.pdf"
excel_path = "/Users/pratikkharat/Downloads/project/Shivani_Mahore_Bank_Statement_23-24.xlsx"

print("Opening PDF...")
pdf = pdfplumber.open(pdf_path)
total_pages = len(pdf.pages)
print(f"Total pages: {total_pages}")

all_rows = []
date_pat = re.compile(r'\d{2}/\w{3}/\d{2,4}')
money_pat = re.compile(r'[\d,]+\.\d{2}')

skip_keywords = [
    'Detailed', 'Statement', 'Name:', 'Address:', 'A/C No',
    'Jt. Holder', 'Transaction Date', 'Transaction Period',
    'Statement Request', 'Advanced Search', 'Amount from',
    'Cheque number', 'Transaction remarks', 'Transaction type',
    'Sl Tran Value', 'No Id Date', 'Page No', 'Closing Balance',
    'Branch:', 'Branch Address', 'A/C Type', 'Cust ID',
    'IFSC Code', 'Account Currency', 'Download', 'from:',
    'WARD', 'PRADESH', 'GANDHI', 'CAA', 'ICICI', 'DR'
]

for i in range(total_pages):
    page = pdf.pages[i]
    text = page.extract_text()
    if not text:
        continue
    lines = text.split('\n')
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        if any(kw in stripped for kw in skip_keywords):
            continue
        match = re.match(r'^(\d+)\s+', stripped)
        if match and date_pat.search(stripped):
            all_rows.append(stripped)
    if (i + 1) % 10 == 0:
        print(f"  Processed {i+1}/{total_pages} pages...")

pdf.close()
print(f"\nTotal transaction lines: {len(all_rows)}")

# Parse rows
parsed = []
for row_text in all_rows:
    amounts = money_pat.findall(row_text)
    sl_match = re.match(r'^(\d+)\s+', row_text)
    sl_no = sl_match.group(1) if sl_match else ""

    tran_match = re.search(r'(S\d{4,})\s', row_text)
    tran_id = tran_match.group(1) if tran_match else ""

    dates = date_pat.findall(row_text)
    value_date = dates[0] if len(dates) > 0 else ""
    txn_date = dates[1] if len(dates) > 1 else ""

    posted_match = re.search(r'(\d{2}/\d{2}/\d{4})', row_text)
    posted_date = posted_match.group(1) if posted_match else ""

    # Extract remarks
    remainder = row_text
    remainder = re.sub(r'^\d+\s+', '', remainder)
    if tran_id:
        remainder = remainder.replace(tran_id, '', 1)
    for d in dates:
        remainder = remainder.replace(d, '', 1)
    if posted_date:
        remainder = remainder.replace(posted_date, '', 1)
    for a in amounts:
        remainder = remainder.replace(a, '', 1)
    remainder = re.sub(r'\d{2}:\d{2}:\d{2}\s*(AM|PM)', '', remainder)
    remainder = re.sub(r'\s+', ' ', remainder).strip().strip('- ')
    remarks = remainder

    withdrawal = ""
    deposit = ""
    balance = ""

    if len(amounts) >= 3:
        withdrawal = amounts[-3]
        deposit = amounts[-2]
        balance = amounts[-1]
    elif len(amounts) == 2:
        balance = amounts[-1]
        withdrawal = amounts[0]
    elif len(amounts) == 1:
        balance = amounts[0]

    parsed.append({
        'Sl No': sl_no,
        'Tran Id': tran_id,
        'Value Date': value_date,
        'Transaction Date': txn_date,
        'Posted Date': posted_date,
        'Transaction Remarks': remarks,
        'Withdrawal (Dr)': withdrawal,
        'Deposit (Cr)': deposit,
        'Balance': balance
    })

df = pd.DataFrame(parsed)
print(f"Parsed {len(df)} transactions")

# Convert money columns
for col in ['Withdrawal (Dr)', 'Deposit (Cr)', 'Balance']:
    df[col] = df[col].apply(lambda x: float(str(x).replace(',', '')) if x else None)

# Save
df.to_excel(excel_path, index=False, sheet_name="Bank Statement")
print("Applying formatting...")

# === FORMATTING ===
wb = load_workbook(excel_path)
ws = wb["Bank Statement"]

header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
data_font = Font(size=10, name="Calibri")
thin_border = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)
alt_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
debit_font = Font(size=10, name="Calibri", color="CC0000")
credit_font = Font(size=10, name="Calibri", color="006600")

# Format headers
for col in range(1, ws.max_column + 1):
    cell = ws.cell(row=1, column=col)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

col_map = {}
for col in range(1, ws.max_column + 1):
    val = ws.cell(row=1, column=col).value
    if val:
        col_map[val] = col

wd_col = col_map.get('Withdrawal (Dr)')
dep_col = col_map.get('Deposit (Cr)')
bal_col = col_map.get('Balance')

for row in range(2, ws.max_row + 1):
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=row, column=col)
        cell.border = thin_border
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        if row % 2 == 0:
            cell.fill = alt_fill
        if col in (wd_col, dep_col, bal_col):
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = '#,##0.00'
            if col == wd_col and cell.value:
                cell.font = debit_font
            elif col == dep_col and cell.value:
                cell.font = credit_font
            else:
                cell.font = data_font
        else:
            cell.font = data_font

# Column widths
widths = {'Sl No': 8, 'Tran Id': 12, 'Value Date': 14, 'Transaction Date': 16,
          'Posted Date': 14, 'Transaction Remarks': 45, 'Withdrawal (Dr)': 18,
          'Deposit (Cr)': 18, 'Balance': 18}
for name, w in widths.items():
    if name in col_map:
        letter = ws.cell(row=1, column=col_map[name]).column_letter
        ws.column_dimensions[letter].width = w

# Summary
sr = ws.max_row + 2
ws.cell(row=sr, column=1, value="SUMMARY").font = Font(bold=True, size=12, name="Calibri")
if wd_col:
    cl = ws.cell(row=1, column=wd_col).column_letter
    ws.cell(row=sr+1, column=wd_col-1, value="Total Withdrawals:").font = Font(bold=True, size=10)
    c = ws.cell(row=sr+1, column=wd_col)
    c.value = f"=SUM({cl}2:{cl}{sr-1})"
    c.number_format = '#,##0.00'
    c.font = Font(bold=True, size=10, color="CC0000")
if dep_col:
    cl = ws.cell(row=1, column=dep_col).column_letter
    ws.cell(row=sr+2, column=dep_col-1, value="Total Deposits:").font = Font(bold=True, size=10)
    c = ws.cell(row=sr+2, column=dep_col)
    c.value = f"=SUM({cl}2:{cl}{sr-1})"
    c.number_format = '#,##0.00'
    c.font = Font(bold=True, size=10, color="006600")

ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

wb.save(excel_path)
print(f"\n✅ Excel saved: {excel_path}")
print("Done!")