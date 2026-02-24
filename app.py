import os, zipfile, shutil, uuid, base64, json, io
import pandas as pd
import fitz
from flask import Flask, request, render_template_string, send_file, jsonify
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = "/tmp/uploads"
REPORT_FOLDER = "/tmp/reports"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url="https://api.deepseek.com")

ALLOWED_BILL_EXT = {".jpg", ".jpeg", ".png", ".pdf"}
IGNORE_FOLDERS = {"payment screenshots", "payments", "misc", "receipts", "__macosx", ".ds_store", "payment"}

HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Purchase Audit System</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Segoe UI', sans-serif; background: #0f172a; color: #e2e8f0; min-height: 100vh; }
        .header { background: linear-gradient(135deg, #1e293b, #0f172a); padding: 20px 40px; border-bottom: 1px solid #334155; display: flex; align-items: center; gap: 15px; }
        .header h1 { font-size: 24px; font-weight: 700; color: #fff; }
        .badge { background: #3b82f6; color: white; padding: 3px 10px; border-radius: 20px; font-size: 12px; }
        .container { max-width: 900px; margin: 40px auto; padding: 0 20px; }
        .card { background: #1e293b; border-radius: 16px; padding: 30px; border: 1px solid #334155; margin-bottom: 20px; }
        .card h2 { font-size: 18px; font-weight: 600; margin-bottom: 20px; color: #f1f5f9; }
        .upload-area { border: 2px dashed #334155; border-radius: 12px; padding: 50px; text-align: center; cursor: pointer; transition: all 0.3s; }
        .upload-area:hover, .upload-area.dragover { border-color: #3b82f6; background: rgba(59,130,246,0.05); }
        .upload-area .icon { font-size: 48px; margin-bottom: 15px; }
        .upload-area p { color: #94a3b8; margin-bottom: 5px; }
        .upload-area strong { color: #3b82f6; }
        .btn { background: #3b82f6; color: white; border: none; padding: 12px 30px; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer; transition: all 0.2s; width: 100%; margin-top: 15px; }
        .btn:hover { background: #2563eb; }
        .btn:disabled { background: #475569; cursor: not-allowed; }
        .progress-area { display: none; }
        .progress-bar-wrap { background: #0f172a; border-radius: 8px; height: 8px; margin: 15px 0; overflow: hidden; }
        .progress-bar { background: linear-gradient(90deg, #3b82f6, #8b5cf6); height: 100%; width: 0%; transition: width 0.5s; border-radius: 8px; }
        .status-text { color: #94a3b8; font-size: 14px; text-align: center; }
        .result-area { display: none; }
        .stats { display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; margin-bottom: 20px; }
        .stat { background: #0f172a; border-radius: 10px; padding: 20px; text-align: center; border: 1px solid #334155; }
        .stat .num { font-size: 32px; font-weight: 700; color: #3b82f6; }
        .stat .label { font-size: 13px; color: #94a3b8; margin-top: 5px; }
        .download-btn { background: linear-gradient(135deg, #10b981, #059669); color: white; border: none; padding: 15px 30px; border-radius: 10px; font-size: 16px; font-weight: 600; cursor: pointer; width: 100%; transition: all 0.2s; }
        .download-btn:hover { transform: translateY(-2px); box-shadow: 0 10px 25px rgba(16,185,129,0.3); }
        .log { background: #0f172a; border-radius: 8px; padding: 15px; max-height: 250px; overflow-y: auto; font-family: monospace; font-size: 13px; margin-top: 15px; }
        .log-item { padding: 3px 0; border-bottom: 1px solid #1e293b; }
        .log-item.ok { color: #10b981; }
        .log-item.err { color: #ef4444; }
        .log-item.info { color: #94a3b8; }
        input[type="file"] { display: none; }
        .powered { text-align: center; color: #475569; font-size: 12px; margin-top: 30px; padding-bottom: 30px; }
        .csv-upload { border: 1px dashed #334155; border-radius: 8px; padding: 15px; text-align: center; cursor: pointer; margin-top: 10px; transition: all 0.3s; }
        .csv-upload:hover { border-color: #3b82f6; }
    </style>
</head>
<body>
    <div class="header">
        <h1>🧾 Purchase Audit System</h1>
        <span class="badge">AI Powered by DeepSeek</span>
    </div>
    <div class="container">
        <div class="card">
            <h2>📁 Upload Bills ZIP</h2>
            <div class="upload-area" id="dropZone" onclick="document.getElementById('fileInput').click()">
                <div class="icon">📦</div>
                <p><strong>Click to upload</strong> or drag & drop</p>
                <p>ZIP file containing bill images (JPG, PNG, PDF)</p>
                <p id="fileName" style="color:#3b82f6; margin-top:10px; font-weight:600;"></p>
            </div>
            <input type="file" id="fileInput" accept=".zip" onchange="handleFile(this)">
            <label style="color:#94a3b8; font-size:14px; display:block; margin-top:15px;">Bill CSV (optional — for audit matching):</label>
            <div class="csv-upload" onclick="document.getElementById('csvInput').click()">
                <p style="color:#94a3b8;"><strong style="color:#3b82f6;">Click</strong> to upload Bill.csv</p>
                <p id="csvName" style="color:#3b82f6; font-size:13px; margin-top:5px;"></p>
            </div>
            <input type="file" id="csvInput" accept=".csv" onchange="document.getElementById('csvName').textContent = '✓ ' + (this.files[0]?.name || '')">
            <button class="btn" id="uploadBtn" onclick="startUpload()" disabled>🚀 Start Audit</button>
        </div>
        <div class="card progress-area" id="progressCard">
            <h2>⚙️ Processing Bills with AI...</h2>
            <div class="progress-bar-wrap"><div class="progress-bar" id="progressBar"></div></div>
            <p class="status-text" id="statusText">Initializing...</p>
            <div class="log" id="logArea"></div>
        </div>
        <div class="card result-area" id="resultCard">
            <h2>✅ Audit Complete!</h2>
            <div class="stats">
                <div class="stat"><div class="num" id="statTotal">0</div><div class="label">Bills Processed</div></div>
                <div class="stat"><div class="num" id="statMatched" style="color:#10b981;">0</div><div class="label">Matched</div></div>
                <div class="stat"><div class="num" id="statMismatch" style="color:#ef4444;">0</div><div class="label">Mismatches / Unmatched</div></div>
            </div>
            <button class="download-btn" onclick="downloadReport()">📥 Download Excel Report</button>
        </div>
        <div class="powered">Powered by DeepSeek AI + Flask on Render</div>
    </div>
    <script>
        let zipFile = null, reportFile = null;
        function handleFile(input) {
            zipFile = input.files[0];
            document.getElementById('fileName').textContent = zipFile ? '✓ ' + zipFile.name : '';
            document.getElementById('uploadBtn').disabled = !zipFile;
        }
        const dropZone = document.getElementById('dropZone');
        dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
        dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
        dropZone.addEventListener('drop', e => {
            e.preventDefault(); dropZone.classList.remove('dragover');
            const f = e.dataTransfer.files[0];
            if (f && f.name.endsWith('.zip')) { zipFile = f; document.getElementById('fileName').textContent = '✓ ' + f.name; document.getElementById('uploadBtn').disabled = false; }
        });
        function addLog(msg, type='info') {
            const log = document.getElementById('logArea');
            const div = document.createElement('div');
            div.className = 'log-item ' + type;
            div.textContent = msg;
            log.appendChild(div);
            log.scrollTop = log.scrollHeight;
        }
        async function startUpload() {
            if (!zipFile) return;
            const btn = document.getElementById('uploadBtn');
            btn.disabled = true; btn.textContent = 'Processing...';
            document.getElementById('progressCard').style.display = 'block';
            document.getElementById('resultCard').style.display = 'none';
            document.getElementById('logArea').innerHTML = '';
            document.getElementById('progressBar').style.width = '10%';
            document.getElementById('statusText').textContent = 'Uploading...';
            const formData = new FormData();
            formData.append('zip_file', zipFile);
            const csvFile = document.getElementById('csvInput').files[0];
            if (csvFile) formData.append('csv_file', csvFile);
            addLog('📤 Uploading and extracting ZIP...', 'info');
            try {
                const response = await fetch('/upload', { method: 'POST', body: formData });
                const data = await response.json();
                document.getElementById('progressBar').style.width = '100%';
                if (data.success) {
                    reportFile = data.report_file;
                    document.getElementById('statusText').textContent = 'Done!';
                    data.logs.forEach(l => addLog(l.msg, l.type));
                    document.getElementById('statTotal').textContent = data.total;
                    document.getElementById('statMatched').textContent = data.matched;
                    document.getElementById('statMismatch').textContent = data.mismatch;
                    document.getElementById('resultCard').style.display = 'block';
                } else { addLog('❌ Error: ' + data.error, 'err'); document.getElementById('statusText').textContent = 'Failed'; }
            } catch(e) { addLog('❌ ' + e.message, 'err'); }
            btn.disabled = false; btn.textContent = '🚀 Start Audit';
        }
        function downloadReport() { if (reportFile) window.location.href = '/download/' + reportFile; }
    </script>
</body>
</html>
"""

def is_ignored(path):
    parts = [p.strip().lower() for p in path.replace("\\", "/").split("/")]
    return any(p in IGNORE_FOLDERS for p in parts)

def pdf_to_images(pdf_path, output_dir):
    images = []
    try:
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_path = os.path.join(output_dir, f"pdf_page_{i}.jpg")
            pix.save(img_path)
            images.append(img_path)
    except:
        pass
    return images

def ocr_image(path):
    text = ""
    try:
        import cv2, pytesseract
        img = cv2.imread(path)
        if img is not None:
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            gray = cv2.fastNlMeansDenoising(gray, h=10)
            processed = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2)
            text = pytesseract.image_to_string(processed, config="--psm 6")
    except:
        pass
    return text.strip()

def parse_with_deepseek(raw_text, filename=""):
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": "You are an expert at parsing Indian bill/invoice OCR text. Extract structured fields accurately. Return ONLY valid JSON."},
                {"role": "user", "content": f"""Parse this OCR text from an Indian bill/invoice:

OCR TEXT:
{raw_text[:2000]}

Return ONLY this JSON (no extra text):
{{
  "bill_number": "invoice/bill number",
  "bill_date": "date in DD/MM/YYYY",
  "vendor_name": "vendor/supplier/agency name",
  "customer_name": "customer/hotel/branch name",
  "item_description": "main item or particulars",
  "quantity": "quantity as number",
  "rate": "unit rate as number only",
  "total_amount": "final total as number only"
}}"""}
            ],
            max_tokens=300,
            temperature=0.1
        )
        result = response.choices[0].message.content.strip()
        if "```json" in result:
            result = result.split("```json").split("```").strip()[1]
        elif "```" in result:
            result = result.split("```")[2].split("```")[0].strip()
        return json.loads(result)
    except Exception as e:
        return {"bill_number": "", "bill_date": "", "vendor_name": "",
                "customer_name": "", "item_description": "", "quantity": "",
                "rate": "", "total_amount": "", "error": str(e)}

def generate_excel(results, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    headers = ["File Name", "Folder", "Bill Number", "Bill Date", "Vendor Name",
               "Customer/Hotel", "Item Description", "Quantity", "Rate (₹)", "Total Amount (₹)", "Match Status"]
    hfill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    hfont = Font(bold=True, color="FFFFFF", size=11)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30
    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    alt = PatternFill(start_color="EEF4FF", end_color="EEF4FF", fill_type="solid")
    for i, r in enumerate(results, 2):
        row_data = [r.get('file_name',''), r.get('folder',''), r.get('bill_number',''),
                    r.get('bill_date',''), r.get('vendor_name',''), r.get('customer_name',''),
                    r.get('item_description',''), r.get('quantity',''), r.get('rate',''),
                    r.get('total_amount',''), r.get('match_status','')]
        status = r.get('match_status','')
        row_fill = green if status == 'Matched' else (red if status == 'Not Found' else (alt if i % 2 == 0 else None))
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.alignment = Alignment(vertical="center")
            if row_fill: cell.fill = row_fill
    for col, width in enumerate([20,15,15,12,30,25,25,10,12,15,12], 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width
    ws.freeze_panes = "A2"
    wb.save(output_path)

@app.route('/')
def index():
    return render_template_string(HTML)

@app.route('/upload', methods=['POST'])
def upload():
    session_id = str(uuid.uuid4())[:8]
    work_dir = os.path.join(UPLOAD_FOLDER, session_id)
    os.makedirs(work_dir, exist_ok=True)
    logs = []
    results = []
    try:
        zip_file = request.files.get('zip_file')
        if not zip_file:
            return jsonify({"success": False, "error": "No ZIP file"})
        zip_path = os.path.join(work_dir, "bills.zip")
        zip_file.save(zip_path)
        csv_df = None
        csv_file = request.files.get('csv_file')
        if csv_file:
            csv_path = os.path.join(work_dir, "bills.csv")
            csv_file.save(csv_path)
            try:
                csv_df = pd.read_csv(csv_path)
                logs.append({"msg": f"📋 Loaded CSV: {len(csv_df)} rows", "type": "ok"})
            except:
                logs.append({"msg": "⚠️ Could not read CSV", "type": "err"})
        extract_dir = os.path.join(work_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        with zipfile.ZipFile(zip_path, 'r') as z:
            z.extractall(extract_dir)
        bill_images = []
        for root, dirs, files in os.walk(extract_dir):
            for f in files:
                rel = os.path.relpath(os.path.join(root, f), extract_dir)
                if is_ignored(rel):
                    continue
                ext = os.path.splitext(f)[1].lower()
                full_path = os.path.join(root, f)
                if ext == ".pdf":
                    imgs = pdf_to_images(full_path, root)
                    bill_images.extend([(img, rel) for img in imgs])
                elif ext in ALLOWED_BILL_EXT:
                    bill_images.append((full_path, rel))
        logs.append({"msg": f"📂 Found {len(bill_images)} bill files", "type": "info"})
        for img_path, rel_path in bill_images:
            fname = os.path.basename(rel_path)
            logs.append({"msg": f"🔍 Reading: {fname}", "type": "info"})
            raw_text = ocr_image(img_path)
            if not raw_text:
                logs.append({"msg": f"  ⚠️ No text extracted from {fname}", "type": "err"})
                raw_text = f"filename: {fname}"
            data = parse_with_deepseek(raw_text, fname)
            data['file_name'] = fname
            data['folder'] = os.path.dirname(rel_path)
            match_status = "No CSV"
            if csv_df is not None:
                match_status = "Not Found"
                bill_no = str(data.get('bill_number', '')).strip()
                if bill_no:
                    for col in csv_df.columns:
                        col_lower = col.lower()
                        if any(k in col_lower for k in ['bill', 'invoice', 'no', 'number']):
                            if csv_df[col].astype(str).str.contains(bill_no, na=False, case=False).any():
                                match_status = "Matched"
                                break
            data['match_status'] = match_status
            results.append(data)
            logs.append({"msg": f"  ✅ #{data.get('bill_number','?')} | {data.get('vendor_name','?')[:25]} | ₹{data.get('total_amount','?')} | {match_status}", "type": "ok"})
        report_name = f"audit_{session_id}.xlsx"
        report_path = os.path.join(REPORT_FOLDER, report_name)
        generate_excel(results, report_path)
        matched = sum(1 for r in results if r.get('match_status') == 'Matched')
        mismatch = sum(1 for r in results if r.get('match_status') == 'Not Found')
        logs.append({"msg": f"📊 Excel report ready!", "type": "ok"})
        return jsonify({"success": True, "total": len(results), "matched": matched,
                        "mismatch": mismatch, "report_file": report_name, "logs": logs})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(REPORT_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name=filename)
    return "File not found", 404

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8000))
    app.run(debug=False, host='0.0.0.0', port=port)
