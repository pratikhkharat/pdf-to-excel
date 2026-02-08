from flask import Flask, request, send_file, render_template_string, jsonify
import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import re
import io
import os
import tempfile
import traceback

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>PDF to Excel Converter</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            background: linear-gradient(135deg, #0f0c29, #302b63, #24243e);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .container {
            background: rgba(255,255,255,0.95);
            border-radius: 24px;
            padding: 48px;
            max-width: 680px;
            width: 100%;
            box-shadow: 0 25px 60px rgba(0,0,0,0.3);
        }
        .header { text-align: center; margin-bottom: 36px; }
        .header h1 {
            font-size: 2rem;
            background: linear-gradient(135deg, #1F4E79, #2E86C1);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            margin-bottom: 8px;
        }
        .header p { color: #666; font-size: 1rem; }
        .upload-area {
            border: 3px dashed #c0d6e8;
            border-radius: 16px;
            padding: 48px 24px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s ease;
            background: #f8fafc;
            margin-bottom: 24px;
            position: relative;
        }
        .upload-area:hover, .upload-area.dragover {
            border-color: #2E86C1;
            background: #eef5fb;
            transform: translateY(-2px);
        }
        .upload-area .icon { font-size: 3rem; margin-bottom: 12px; }
        .upload-area p { color: #555; font-size: 1rem; }
        .upload-area .filename {
            margin-top: 12px;
            font-weight: 600;
            color: #1F4E79;
            font-size: 0.95rem;
        }
        .upload-area input[type="file"] {
            position: absolute;
            top: 0; left: 0; width: 100%; height: 100%;
            opacity: 0; cursor: pointer;
        }
        .mode-select { display: flex; gap: 12px; margin-bottom: 24px; }
        .mode-option {
            flex: 1; padding: 14px;
            border: 2px solid #e0e0e0;
            border-radius: 12px;
            text-align: center;
            cursor: pointer;
            transition: all 0.2s;
            font-size: 0.9rem;
        }
        .mode-option:hover { border-color: #2E86C1; }
        .mode-option.active {
            border-color: #1F4E79;
            background: #eef5fb;
            font-weight: 600;
        }
        .mode-option input { display: none; }
        .btn {
            width: 100%; padding: 16px;
            background: linear-gradient(135deg, #1F4E79, #2E86C1);
            color: white; border: none; border-radius: 12px;
            font-size: 1.1rem; font-weight: 600;
            cursor: pointer; transition: all 0.3s;
        }
        .btn:hover { transform: translateY(-2px); box-shadow: 0 8px 25px rgba(31,78,121,0.35); }
        .btn:disabled { background: #ccc; cursor: not-allowed; transform: none; box-shadow: none; }
        .progress-container { display: none; margin-top: 24px; text-align: center; }
        .spinner {
            width: 48px; height: 48px;
            border: 4px solid #e0e0e0;
            border-top: 4px solid #1F4E79;
            border-radius: 50%;
            animation: spin 0.8s linear infinite;
            margin: 0 auto 16px;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .progress-text { color: #555; font-size: 0.95rem; }
        .result {
            display: none; margin-top: 24px; padding: 20px;
            background: #d4edda; border: 1px solid #c3e6cb;
            border-radius: 12px; text-align: center;
        }
        .result h3 { color: #155724; margin-bottom: 8px; }
        .result p { color: #155724; font-size: 0.9rem; }
        .error {
            display: none; margin-top: 24px; padding: 20px;
            background: #f8d7da; border: 1px solid #f5c6cb;
            border-radius: 12px; text-align: center; color: #721c24;
        }
        .features {
            display: grid; grid-template-columns: 1fr 1fr;
            gap: 10px; margin-top: 32px;
            padding-top: 24px; border-top: 1px solid #eee;
        }
        .feature { font-size: 0.82rem; color: #666; padding: 6px 0; }
        .feature span { margin-right: 6px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>&#x1F4CA; PDF to Excel Converter</h1>
            <p>Upload any bank statement PDF &mdash; auto-detects format &amp; creates formatted Excel</p>
        </div>
        <form id="uploadForm" action="/convert" method="POST" enctype="multipart/form-data">
            <div class="upload-area" id="uploadArea">
                <div class="icon">&#x1F4C2;</div>
                <p>Drag &amp; drop your PDF here or click to browse</p>
                <div class="filename" id="fileName"></div>
                <input type="file" name="pdf_file" id="fileInput" accept=".pdf" required>
            </div>
            <div class="mode-select">
                <label class="mode-option active" id="mode1">
                    <input type="radio" name="mode" value="text" checked>
                    &#x1F524; Auto (Text-based)
                </label>
                <label class="mode-option" id="mode2">
                    <input type="radio" name="mode" value="table">
                    &#x1F4CB; Table-based
                </label>
            </div>
            <button type="submit" class="btn" id="convertBtn" disabled>
                &#x1F680; Convert to Excel
            </button>
        </form>
        <div class="progress-container" id="progress">
            <div class="spinner"></div>
            <p class="progress-text">Processing your PDF... Large files may take up to 2 minutes.</p>
        </div>
        <div class="result" id="result">
            <h3>&#x2705; Conversion Complete!</h3>
            <p id="resultText"></p>
        </div>
        <div class="error" id="error"></div>
        <div class="features">
            <div class="feature"><span>&#x1F3A8;</span> Color-coded headers</div>
            <div class="feature"><span>&#x1F534;</span> Withdrawals in red</div>
            <div class="feature"><span>&#x1F7E2;</span> Deposits in green</div>
            <div class="feature"><span>&#x1F4CA;</span> Alternating row colors</div>
            <div class="feature"><span>&#x1F4CC;</span> Frozen header row</div>
            <div class="feature"><span>&#x1F53D;</span> Auto-filter enabled</div>
            <div class="feature"><span>&#x2795;</span> Summary totals</div>
            <div class="feature"><span>&#x1F3E6;</span> Multi-bank support</div>
        </div>
    </div>
    <script>
        const fileInput = document.getElementById('fileInput');
        const fileName = document.getElementById('fileName');
        const convertBtn = document.getElementById('convertBtn');
        const uploadArea = document.getElementById('uploadArea');
        const form = document.getElementById('uploadForm');
        const progress = document.getElementById('progress');
        const result = document.getElementById('result');
        const errorDiv = document.getElementById('error');
        fileInput.addEventListener('change', function() {
            if (this.files.length > 0) {
                const f = this.files[0];
                const sizeMB = (f.size / 1024 / 1024).toFixed(1);
                fileName.textContent = f.name + ' (' + sizeMB + ' MB)';
                convertBtn.disabled = false;
            }
        });
        uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.classList.add('dragover'); });
        uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
        uploadArea.addEventListener('drop', (e) => { e.preventDefault(); uploadArea.classList.remove('dragover'); });
        document.querySelectorAll('.mode-option').forEach(opt => {
            opt.addEventListener('click', () => {
                document.querySelectorAll('.mode-option').forEach(o => o.classList.remove('active'));
                opt.classList.add('active');
            });
        });
        function showError(msg) {
            progress.style.display = 'none';
            result.style.display = 'none';
            errorDiv.style.display = 'block';
            errorDiv.textContent = 'Error: ' + msg;
            convertBtn.disabled = false;
        }
        form.addEventListener('submit', function(e) {
            e.preventDefault();
            convertBtn.disabled = true;
            progress.style.display = 'block';
            result.style.display = 'none';
            errorDiv.style.display = 'none';
            const formData = new FormData(form);
            fetch('/convert', { method: 'POST', body: formData })
                .then(async response => {
                    const contentType = response.headers.get('content-type') || '';
                    if (!response.ok) {
                        if (contentType.includes('application/json')) {
                            const data = await response.json();
                            throw new Error(data.error || 'Unknown server error');
                        } else {
                            const text = await response.text();
                            throw new Error('Server error (' + response.status + '). The PDF may be too large or in an unsupported format.');
                        }
                    }
                    if (contentType.includes('application/json')) {
                        const data = await response.json();
                        throw new Error(data.error || 'Unknown error');
                    }
                    const txnCount = response.headers.get('X-Transaction-Count') || '?';
                    const bankName = response.headers.get('X-Bank-Name') || 'Bank';
                    const blob = await response.blob();
                    return { blob, txnCount, bankName };
                })
                .then(({ blob, txnCount, bankName }) => {
                    progress.style.display = 'none';
                    result.style.display = 'block';
                    document.getElementById('resultText').textContent =
                        'Extracted ' + txnCount + ' transactions from ' + bankName + ' statement.';
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = fileInput.files[0].name.replace('.pdf', '.xlsx').replace('.PDF', '.xlsx');
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                    URL.revokeObjectURL(url);
                    convertBtn.disabled = false;
                })
                .catch(err => {
                    showError(err.message || 'Something went wrong. Please try again.');
                });
        });
    </script>
</body>
</html>
"""

BANK_PROFILES = {
    "icici": {
        "name": "ICICI Bank",
        "keywords": ["ICICI", "icicibank"],
        "skip_keywords": [
            'Detailed', 'Statement', 'Name:', 'Address:', 'A/C No',
            'Jt. Holder', 'Transaction Date from', 'Transaction Period',
            'Statement Request', 'Advanced Search', 'Amount from',
            'Cheque number', 'Transaction remarks', 'Transaction type:',
            'Sl Tran Value', 'No Id Date', 'Page No', 'Closing Balance',
            'Branch:', 'Branch Address', 'A/C Type', 'Cust ID',
            'IFSC Code', 'Account Currency', 'Download', 'from:',
            'WARD', 'PRADESH', 'CAA', 'GANDHI', 'DR',
        ],
    },
    "sbi": {
        "name": "SBI",
        "keywords": ["State Bank", "SBI", "sbi.co.in"],
        "skip_keywords": [
            'State Bank', 'Account Statement', 'Account Number',
            'Address', 'Branch', 'IFSC', 'CIF No', 'Nomination',
            'IFS Code', 'MICR Code', 'Page', 'Opening Balance',
        ],
    },
    "hdfc": {
        "name": "HDFC Bank",
        "keywords": ["HDFC", "hdfcbank"],
        "skip_keywords": [
            'HDFC BANK', 'Statement of Account', 'Account No',
            'Branch', 'Address', 'IFSC', 'Nomination', 'Page',
            'Opening Balance', 'Closing Balance',
        ],
    },
    "generic": {
        "name": "Generic Bank",
        "keywords": [],
        "skip_keywords": [
            'Page', 'Statement', 'Account', 'Branch', 'Address',
            'IFSC', 'Opening Balance', 'Closing Balance',
        ],
    },
}


def detect_bank(text):
    upper = text.upper()
    for key, profile in BANK_PROFILES.items():
        if key == "generic":
            continue
        for kw in profile["keywords"]:
            if kw.upper() in upper:
                return key
    return "generic"


def extract_text_mode(pdf_path):
    pdf = pdfplumber.open(pdf_path)
    total_pages = len(pdf.pages)
    sample = ""
    for i in range(min(3, total_pages)):
        t = pdf.pages[i].extract_text()
        if t:
            sample += t + "\n"
    bank_key = detect_bank(sample)
    profile = BANK_PROFILES[bank_key]
    skip_keywords = profile["skip_keywords"]
    date_pat = re.compile(r'\d{2}/\w{3}/\d{2,4}|\d{2}-\w{3}-\d{2,4}|\d{2}/\d{2}/\d{2,4}|\d{2}-\d{2}-\d{2,4}')
    all_rows = []
    for i in range(total_pages):
        text = pdf.pages[i].extract_text()
        if not text:
            continue
        for line in text.split('\n'):
            stripped = line.strip()
            if not stripped:
                continue
            if any(kw.lower() in stripped.lower() for kw in skip_keywords):
                continue
            if re.match(r'^(\d+)\s+', stripped) and date_pat.search(stripped):
                all_rows.append(stripped)
    pdf.close()
    return all_rows, bank_key, profile


def extract_table_mode(pdf_path):
    pdf = pdfplumber.open(pdf_path)
    total_pages = len(pdf.pages)
    sample = ""
    for i in range(min(3, total_pages)):
        t = pdf.pages[i].extract_text()
        if t:
            sample += t + "\n"
    bank_key = detect_bank(sample)
    profile = BANK_PROFILES[bank_key]
    header_keywords = ["date", "narration", "description", "particular", "debit", "credit",
                        "balance", "withdrawal", "deposit", "amount", "txn", "ref", "cheque",
                        "tran", "value"]
    all_data = []
    headers = None
    for i in range(total_pages):
        tables = pdf.pages[i].extract_tables()
        for table in tables:
            for row in table:
                if not row or not any(cell and cell.strip() for cell in row if cell):
                    continue
                cleaned = [cell.strip() if cell else "" for cell in row]
                row_text = " ".join(cleaned).lower()
                if headers is None:
                    if sum(1 for kw in header_keywords if kw in row_text) >= 2:
                        headers = cleaned
                        continue
                if headers:
                    all_data.append(cleaned)
    pdf.close()
    return all_data, headers, bank_key, profile


def parse_text_rows(all_rows, bank_key):
    date_pat = re.compile(r'\d{2}/\w{3}/\d{2,4}|\d{2}-\w{3}-\d{2,4}|\d{2}/\d{2}/\d{2,4}|\d{2}-\d{2}-\d{2,4}')
    money_pat = re.compile(r'[\d,]+\.\d{2}')
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
        withdrawal, deposit, balance = "", "", ""
        if len(amounts) >= 3:
            withdrawal, deposit, balance = amounts[-3], amounts[-2], amounts[-1]
        elif len(amounts) == 2:
            withdrawal, balance = amounts[0], amounts[-1]
        elif len(amounts) == 1:
            balance = amounts[0]
        if bank_key == "icici":
            parsed.append({
                'Sl No': sl_no, 'Tran Id': tran_id, 'Value Date': value_date,
                'Transaction Date': txn_date, 'Posted Date': posted_date,
                'Transaction Remarks': remainder,
                'Withdrawal (Dr)': withdrawal, 'Deposit (Cr)': deposit, 'Balance': balance,
            })
        else:
            parsed.append({
                'Sl No': sl_no, 'Date': value_date, 'Description': remainder,
                'Withdrawal': withdrawal, 'Deposit': deposit, 'Balance': balance,
            })
    return pd.DataFrame(parsed)


def format_excel(excel_buffer):
    wb = load_workbook(excel_buffer)
    ws = wb.active
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11, name="Calibri")
    data_font = Font(size=10, name="Calibri")
    thin_border = Border(
        left=Side(style="thin", color="B0B0B0"), right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"), bottom=Side(style="thin", color="B0B0B0"),
    )
    alt_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    debit_font = Font(size=10, name="Calibri", color="CC0000")
    credit_font = Font(size=10, name="Calibri", color="006600")
    col_map = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=1, column=c).value
        if v:
            col_map[v] = c
    dr_kw = ['withdrawal', 'debit', 'dr']
    cr_kw = ['deposit', 'credit', 'cr']
    bal_kw = ['balance']
    dr_cols = {col_map[k] for k in col_map if any(m in k.lower() for m in dr_kw)}
    cr_cols = {col_map[k] for k in col_map if any(m in k.lower() for m in cr_kw)}
    bal_cols = {col_map[k] for k in col_map if any(m in k.lower() for m in bal_kw)}
    money_cols = dr_cols | cr_cols | bal_cols
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    for r in range(2, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if r % 2 == 0:
                cell.fill = alt_fill
            if c in money_cols:
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '#,##0.00'
                if c in dr_cols and cell.value:
                    cell.font = debit_font
                elif c in cr_cols and cell.value:
                    cell.font = credit_font
                else:
                    cell.font = data_font
            else:
                cell.font = data_font
    for col_cells in ws.columns:
        max_len = max((len(str(cell.value)) for cell in col_cells if cell.value), default=8)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(max_len + 3, 10), 50)
    data_last = ws.max_row
    sr = data_last + 2
    ws.cell(row=sr, column=1, value="SUMMARY").font = Font(bold=True, size=12, name="Calibri", color="1F4E79")
    offset = 1
    for ci in dr_cols:
        cl = ws.cell(row=1, column=ci).column_letter
        ws.cell(row=sr + offset, column=max(1, ci - 1), value="Total Withdrawals:").font = Font(bold=True, size=10)
        c = ws.cell(row=sr + offset, column=ci)
        c.value = f"=SUM({cl}2:{cl}{data_last})"
        c.number_format = '#,##0.00'
        c.font = Font(bold=True, size=11, color="CC0000")
        offset += 1
    for ci in cr_cols:
        cl = ws.cell(row=1, column=ci).column_letter
        ws.cell(row=sr + offset, column=max(1, ci - 1), value="Total Deposits:").font = Font(bold=True, size=10)
        c = ws.cell(row=sr + offset, column=ci)
        c.value = f"=SUM({cl}2:{cl}{data_last})"
        c.number_format = '#,##0.00'
        c.font = Font(bold=True, size=11, color="006600")
        offset += 1
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{ws.cell(row=1, column=ws.max_column).column_letter}{data_last}"
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


@app.route("/convert", methods=["POST"])
def convert():
    try:
        if "pdf_file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
        pdf_file = request.files["pdf_file"]
        if pdf_file.filename == "":
            return jsonify({"error": "No file selected"}), 400
        mode = request.form.get("mode", "text")
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        pdf_file.save(tmp.name)
        tmp.close()
        try:
            if mode == "text":
                all_rows, bank_key, profile = extract_text_mode(tmp.name)
                df = parse_text_rows(all_rows, bank_key)
            else:
                all_data, headers, bank_key, profile = extract_table_mode(tmp.name)
                if headers and all_data:
                    max_cols = len(headers)
                    normalized = []
                    for row in all_data:
                        if len(row) < max_cols:
                            row += [""] * (max_cols - len(row))
                        elif len(row) > max_cols:
                            row = row[:max_cols]
                        normalized.append(row)
                    df = pd.DataFrame(normalized, columns=headers)
                else:
                    df = pd.DataFrame()
            # Clean up temp file
            if os.path.exists(tmp.name):
                os.unlink(tmp.name)
            if df.empty or len(df) == 0:
                return jsonify({"error": "No transactions found. Try switching to Table-based mode."}), 400
            # Convert money columns
            money_kws = ['withdrawal', 'deposit', 'balance', 'debit', 'credit', 'amount', 'dr', 'cr']
            for col in df.columns:
                if any(kw in col.lower() for kw in money_kws):
                    df[col] = df[col].apply(
                        lambda x: float(str(x).replace(',', ''))
                        if x and re.match(r'^[\d,]+\.?\d*$', str(x).replace(',', '').strip() or '0')
                        else None
                    )
            df = df.dropna(how='all').reset_index(drop=True)
            # Generate Excel
            excel_buf = io.BytesIO()
            df.to_excel(excel_buf, index=False, sheet_name="Bank Statement")
            excel_buf.seek(0)
            formatted = format_excel(excel_buf)
            output_name = pdf_file.filename.replace(".pdf", ".xlsx").replace(".PDF", ".xlsx")
            response = send_file(
                formatted,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True,
                download_name=output_name,
            )
            response.headers["X-Transaction-Count"] = str(len(df))
            response.headers["X-Bank-Name"] = profile["name"]
            response.headers["Access-Control-Expose-Headers"] = "X-Transaction-Count, X-Bank-Name"
            return response
        except Exception as e:
            if os.path.exists(tmp.name):
                os.unlink(tmp.name)
            raise e
    except Exception as e:
        app.logger.error(f"Convert error: {traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("\n" + "=" * 50)
    print("  PDF to Excel Converter")
    print("  Open in browser: http://localhost:5000")
    print("=" * 50 + "\n")
    app.run(debug=True, host="0.0.0.0", port=5000)
