import os
import zipfile
import shutil
import uuid
import re
import cv2
import pytesseract
import pandas as pd
import fitz  # PyMuPDF — for PDF support
from flask import Flask, request, render_template_string, send_file, jsonify
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

app = Flask(__name__)

# ── Folders ────────────────────────────────────────────────────────────────────
UPLOAD_FOLDER = "/tmp/uploads"
REPORT_FOLDER = "/tmp/reports"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

ALLOWED_BILL_EXT = {".jpg", ".jpeg", ".png", ".pdf"}
IGNORE_FOLDERS   = {
    "payment screenshots", "payments", "misc",
    "receipts", "__macosx", ".ds_store", "payment"
}

# ── HTML UI ────────────────────────────────────────────────────────────────────
HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Purchase Audit System</title>
  <style>
    *{margin:0;padding:0;box-sizing:border-box}
    body{font-family:'Segoe UI',sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh}
    .header{background:linear-gradient(135deg,#1e293b,#0f172a);padding:20px 40px;
            border-bottom:1px solid #334155;display:flex;align-items:center;gap:15px}
    .header h1{font-size:24px;font-weight:700;color:#fff}
    .badge{background:#3b82f6;color:#fff;padding:3px 10px;border-radius:20px;font-size:12px}
    .container{max-width:900px;margin:40px auto;padding:0 20px}
    .card{background:#1e293b;border-radius:16px;padding:30px;
          border:1px solid #334155;margin-bottom:20px}
    .card h2{font-size:18px;font-weight:600;margin-bottom:20px;color:#f1f5f9}
    .upload-area{border:2px dashed #334155;border-radius:12px;padding:50px;
                 text-align:center;cursor:pointer;transition:all .3s}
    .upload-area:hover,.upload-area.dragover{border-color:#3b82f6;background:rgba(59,130,246,.05)}
    .upload-area .icon{font-size:48px;margin-bottom:15px}
    .upload-area p{color:#94a3b8;margin-bottom:5px}
    .upload-area strong{color:#3b82f6}
    .btn{background:#3b82f6;color:#fff;border:none;padding:12px 30px;border-radius:8px;
         font-size:15px;font-weight:600;cursor:pointer;transition:all .2s;width:100%;margin-top:15px}
    .btn:hover{background:#2563eb}
    .btn:disabled{background:#475569;cursor:not-allowed}
    .progress-area{display:none}
    .progress-bar-wrap{background:#0f172a;border-radius:8px;height:8px;margin:15px 0;overflow:hidden}
    .progress-bar{background:linear-gradient(90deg,#3b82f6,#8b5cf6);height:100%;
                  width:0%;transition:width .5s;border-radius:8px}
    .status-text{color:#94a3b8;font-size:14px;text-align:center}
    .result-area{display:none}
    .stats{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin-bottom:20px}
    .stat{background:#0f172a;border-radius:10px;padding:20px;text-align:center;
          border:1px solid #334155}
    .stat .num{font-size:32px;font-weight:700;color:#3b82f6}
    .stat .label{font-size:13px;color:#94a3b8;margin-top:5px}
    .download-btn{background:linear-gradient(135deg,#10b981,#059669);color:#fff;border:none;
                  padding:15px 30px;border-radius:10px;font-size:16px;font-weight:600;
                  cursor:pointer;width:100%;transition:all .2s}
    .download-btn:hover{transform:translateY(-2px);box-shadow:0 10px 25px rgba(16,185,129,.3)}
    .log{background:#0f172a;border-radius:8px;padding:15px;max-height:250px;overflow-y:auto;
         font-family:monospace;font-size:13px;margin-top:15px}
    .log-item{padding:3px 0;border-bottom:1px solid #1e293b}
    .log-item.ok{color:#10b981}
    .log-item.err{color:#ef4444}
    .log-item.info{color:#94a3b8}
    input[type="file"]{display:none}
    .powered{text-align:center;color:#475569;font-size:12px;margin-top:30px;padding-bottom:30px}
    .csv-upload{border:1px dashed #334155;border-radius:8px;padding:15px;text-align:center;
                cursor:pointer;margin-top:10px;transition:all .3s}
    .csv-upload:hover{border-color:#3b82f6}
  </style>
</head>
<body>
  <div class="header">
    <h1>&#x1F9FE; Purchase Audit System</h1>
    <span class="badge">OCR + Smart Matching</span>
  </div>
  <div class="container">
    <div class="card">
      <h2>&#x1F4C1; Upload Files</h2>
      <div class="upload-area" id="dropZone"
           onclick="document.getElementById('fileInput').click()">
        <div class="icon">&#x1F4E6;</div>
        <p><strong>Click to upload</strong> or drag &amp; drop</p>
        <p>ZIP file containing bill images (JPG, PNG, PDF)</p>
        <p id="fileName" style="color:#3b82f6;margin-top:10px;font-weight:600;"></p>
      </div>
      <input type="file" id="fileInput" accept=".zip" onchange="handleFile(this)">
      <label style="color:#94a3b8;font-size:14px;display:block;margin-top:15px;">
        Zoho Excel / CSV (for matching):
      </label>
      <div class="csv-upload" onclick="document.getElementById('csvInput').click()">
        <p style="color:#94a3b8;">
          <strong style="color:#3b82f6;">Click</strong> to upload Excel / CSV
        </p>
        <p id="csvName" style="color:#3b82f6;font-size:13px;margin-top:5px;"></p>
      </div>
      <input type="file" id="csvInput" accept=".csv,.xlsx,.xls"
             onchange="document.getElementById('csvName').textContent =
                       '&#10003; ' + (this.files[0]?.name || '')">
      <button class="btn" id="uploadBtn" onclick="startUpload()" disabled>
        &#x1F680; Start Audit
      </button>
    </div>
    <div class="card progress-area" id="progressCard">
      <h2>&#x2699;&#xFE0F; Processing Bills...</h2>
      <div class="progress-bar-wrap">
        <div class="progress-bar" id="progressBar"></div>
      </div>
      <p class="status-text" id="statusText">Initializing...</p>
      <div class="log" id="logArea"></div>
    </div>
    <div class="card result-area" id="resultCard">
      <h2>&#x2705; Audit Complete!</h2>
      <div class="stats">
        <div class="stat">
          <div class="num" id="statTotal">0</div>
          <div class="label">Bills Processed</div>
        </div>
        <div class="stat">
          <div class="num" id="statMatched" style="color:#10b981;">0</div>
          <div class="label">Matched</div>
        </div>
        <div class="stat">
          <div class="num" id="statMismatch" style="color:#ef4444;">0</div>
          <div class="label">Not Found / Mismatch</div>
        </div>
      </div>
      <button class="download-btn" onclick="downloadReport()">
        &#x1F4E5; Download Excel Report
      </button>
    </div>
    <div class="powered">
      Purchase Audit System &mdash; Built by Shubham Hulsure
    </div>
  </div>
  <script>
    let zipFile = null, reportFile = null;
    function handleFile(input) {
      zipFile = input.files[0];
      document.getElementById('fileName').textContent = zipFile ? '\u2713 ' + zipFile.name : '';
      document.getElementById('uploadBtn').disabled = !zipFile;
    }
    const dropZone = document.getElementById('dropZone');
    dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', e => {
      e.preventDefault(); dropZone.classList.remove('dragover');
      const f = e.dataTransfer.files[0];
      if (f && f.name.endsWith('.zip')) {
        zipFile = f;
        document.getElementById('fileName').textContent = '\u2713 ' + f.name;
        document.getElementById('uploadBtn').disabled = false;
      }
    });
    function addLog(msg, type = 'info') {
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
      document.getElementById('resultCard').style.display  = 'none';
      document.getElementById('logArea').innerHTML = '';
      document.getElementById('progressBar').style.width  = '10%';
      document.getElementById('statusText').textContent   = 'Uploading...';
      const formData = new FormData();
      formData.append('zip_file', zipFile);
      const csvFile = document.getElementById('csvInput').files[0];
      if (csvFile) formData.append('csv_file', csvFile);
      addLog('Uploading and extracting ZIP...', 'info');
      try {
        const response = await fetch('/upload', { method: 'POST', body: formData });
        const data = await response.json();
        document.getElementById('progressBar').style.width = '100%';
        if (data.success) {
          reportFile = data.report_file;
          document.getElementById('statusText').textContent = 'Done!';
          data.logs.forEach(l => addLog(l.msg, l.type));
          document.getElementById('statTotal').textContent    = data.total;
          document.getElementById('statMatched').textContent  = data.matched;
          document.getElementById('statMismatch').textContent = data.mismatch;
          document.getElementById('resultCard').style.display = 'block';
        } else {
          addLog('Error: ' + data.error, 'err');
          document.getElementById('statusText').textContent = 'Failed';
        }
      } catch (e) { addLog('Error: ' + e.message, 'err'); }
      btn.disabled = false; btn.textContent = '\uD83D\uDE80 Start Audit';
    }
    function downloadReport() {
      if (reportFile) window.location.href = '/download/' + reportFile;
    }
  </script>
</body>
</html>
"""

# ── Utilities ──────────────────────────────────────────────────────────────────

def is_ignored(path):
    parts = [p.strip().lower() for p in path.replace("\\", "/").split("/")]
    return any(p in IGNORE_FOLDERS for p in parts)

def normalize(text):
    if not text: return ""
    text = str(text).lower()
    text = re.sub(r"[^a-z0-9\s\.]", " ", text)
    return re.sub(r"\s+", " ", text).strip()

def extract_numbers(text):
    return set(re.findall(r"\b\d+(?:\.\d+)?\b", text))

# ── OCR ────────────────────────────────────────────────────────────────────────

def pdf_to_images(pdf_path, output_dir):
    images = []
    try:
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            pix      = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_path = os.path.join(output_dir, f"pdf_page_{i}.jpg")
            pix.save(img_path)
            images.append(img_path)
    except Exception as e:
        print(f"PDF error {pdf_path}: {e}")
    return images

def ocr_image(path):
    try:
        img = cv2.imread(path)
        if img is None: return ""
        gray    = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray    = cv2.GaussianBlur(gray, (3, 3), 0)
        _, proc = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        raw     = pytesseract.image_to_string(proc, config="--psm 6")
        return normalize(raw)
    except Exception as e:
        print(f"OCR error {path}: {e}")
        return ""

# ── Matching ───────────────────────────────────────────────────────────────────

def score_match(bill_no, vendor, item, amount_str, ocr_text):
    if not ocr_text: return 0
    score = 0
    if bill_no and len(bill_no) >= 3:
        if bill_no in ocr_text:
            score += 60
        else:
            score += fuzz.partial_ratio(bill_no, ocr_text) * 0.30
    if vendor and len(vendor) >= 3:
        score += fuzz.partial_ratio(vendor, ocr_text) * 0.20
    if item and len(item) >= 2:
        score += fuzz.partial_ratio(item, ocr_text) * 0.10
    if amount_str:
        amt_stripped = re.sub(r"\.0+$", "", amount_str.strip())
        ocr_numbers  = extract_numbers(ocr_text)
        if amt_stripped in ocr_numbers or amount_str.strip() in ocr_numbers:
            score += 10
    return min(round(score, 1), 100)

def classify(score):
    if score >= 70:
        return "Matched",              "Strong match — bill number + fields confirmed"
    elif score >= 45:
        return "Mismatch / Duplicate", "Partial match — manual verification required"
    else:
        return "Not Found",            "No matching bill image detected"

# ── Excel report ───────────────────────────────────────────────────────────────

def generate_excel(results, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Audit Report"
    headers = [
        "File Name", "Folder", "Bill Number", "Bill Date",
        "Vendor Name", "Customer / Hotel", "Item Description",
        "Quantity", "Rate (Rs)", "Total Amount (Rs)",
        "AI Confidence", "Match Status"
    ]
    hfill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    hfont = Font(bold=True, color="FFFFFF", size=11)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = hfill; cell.font = hfont
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    green  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    alt    = PatternFill(start_color="EEF4FF", end_color="EEF4FF", fill_type="solid")

    for i, r in enumerate(results, 2):
        row_data = [
            r.get("file_name",""), r.get("folder",""), r.get("bill_number",""),
            r.get("bill_date",""), r.get("vendor_name",""), r.get("customer_name",""),
            r.get("item_description",""), r.get("quantity",""), r.get("rate",""),
            r.get("total_amount",""), r.get("confidence",""), r.get("match_status",""),
        ]
        status = r.get("match_status","")
        if   status == "Matched":             row_fill = green
        elif status == "Not Found":           row_fill = red
        elif "Mismatch" in status:            row_fill = yellow
        else:                                 row_fill = alt if i % 2 == 0 else None

        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.alignment = Alignment(vertical="center")
            if row_fill: cell.fill = row_fill

    col_widths = [22, 18, 15, 12, 30, 25, 25, 10, 12, 15, 12, 14]
    for col, width in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width
    ws.freeze_panes = "A2"
    wb.save(output_path)

# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/upload", methods=["POST"])
def upload():
    session_id = str(uuid.uuid4())[:8]
    work_dir   = os.path.join(UPLOAD_FOLDER, session_id)
    os.makedirs(work_dir, exist_ok=True)
    logs = []; results = []

    try:
        zip_file = request.files.get("zip_file")
        if not zip_file:
            return jsonify({"success": False, "error": "No ZIP file provided"})

        zip_path = os.path.join(work_dir, "bills.zip")
        zip_file.save(zip_path)

        csv_df = None
        csv_file = request.files.get("csv_file")
        if csv_file:
            csv_path = os.path.join(work_dir, "data.xlsx")
            csv_file.save(csv_path)
            try:
                csv_df = pd.read_csv(csv_path) if csv_path.lower().endswith(".csv") \
                         else pd.read_excel(csv_path, engine="openpyxl")
                csv_df.columns = csv_df.columns.str.strip()
                logs.append({"msg": f"Loaded reference file: {len(csv_df)} rows", "type": "ok"})
            except Exception as e:
                logs.append({"msg": f"Could not read reference file: {e}", "type": "err"})

        extract_dir = os.path.join(work_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(extract_dir)

        bill_images = []
        for root, dirs, files in os.walk(extract_dir):
            for f in files:
                rel       = os.path.relpath(os.path.join(root, f), extract_dir)
                full_path = os.path.join(root, f)
                if is_ignored(rel): continue
                ext = os.path.splitext(f)[1].lower()
                if ext == ".pdf":
                    imgs = pdf_to_images(full_path, root)
                    bill_images.extend([(img, rel) for img in imgs])
                elif ext in ALLOWED_BILL_EXT:
                    bill_images.append((full_path, rel))

        logs.append({"msg": f"Found {len(bill_images)} bill files in ZIP", "type": "info"})

        ocr_cache = []
        for img_path, rel_path in bill_images:
            fname   = os.path.basename(rel_path)
            ocr_txt = ocr_image(img_path)
            if ocr_txt:
                ocr_cache.append((img_path, ocr_txt))
                logs.append({"msg": f"OCR OK: {fname}", "type": "ok"})
            else:
                logs.append({"msg": f"OCR empty: {fname}", "type": "err"})

        if csv_df is not None:
            logs.append({"msg": "Matching Excel rows to bill images...", "type": "info"})
            for _, row in csv_df.iterrows():
                bill_no = normalize(str(row.get("Bill Number", "")))
                vendor  = normalize(str(row.get("Vendor Name", "")))
                item    = normalize(str(row.get("Item Name", "")))
                amount  = normalize(str(row.get("Item Total", "")))

                best_score = 0; best_path = ""
                for path, text in ocr_cache:
                    s = score_match(bill_no, vendor, item, amount, text)
                    if s > best_score:
                        best_score = s; best_path = path

                status, remark = classify(best_score)
                results.append({
                    "file_name":        os.path.basename(best_path) if best_path else "",
                    "folder":           os.path.dirname(best_path)  if best_path else "",
                    "bill_number":      str(row.get("Bill Number",  "")),
                    "bill_date":        str(row.get("Bill Date",    "")),
                    "vendor_name":      str(row.get("Vendor Name",  "")),
                    "customer_name":    str(row.get("Branch Name",  "")),
                    "item_description": str(row.get("Item Name",    "")),
                    "quantity":         str(row.get("Quantity",     "")),
                    "rate":             str(row.get("Rate",         "")),
                    "total_amount":     str(row.get("Item Total",   "")),
                    "confidence":       best_score,
                    "match_status":     status,
                })
                logs.append({"msg": f"{status} | {row.get('Bill Number','?')} | "
                                    f"{str(row.get('Vendor Name','?'))[:20]} | "
                                    f"score {best_score}",
                             "type": "ok" if status == "Matched" else
                                     ("err" if status == "Not Found" else "info")})
        else:
            for img_path, _ in ocr_cache:
                results.append({
                    "file_name": os.path.basename(img_path), "folder": "",
                    "bill_number":"","bill_date":"","vendor_name":"","customer_name":"",
                    "item_description":"","quantity":"","rate":"","total_amount":"",
                    "confidence":"","match_status":"No CSV",
                })

        report_name = f"audit_{session_id}.xlsx"
        report_path = os.path.join(REPORT_FOLDER, report_name)
        generate_excel(results, report_path)

        matched  = sum(1 for r in results if r["match_status"] == "Matched")
        mismatch = sum(1 for r in results
                       if r["match_status"] in ("Not Found", "Mismatch / Duplicate"))
        logs.append({"msg": f"Report ready — {matched} matched, {mismatch} issues",
                     "type": "ok"})

        return jsonify({"success": True, "total": len(results),
                        "matched": matched, "mismatch": mismatch,
                        "report_file": report_name, "logs": logs})

    except Exception as e:
        return jsonify({"success": False, "error": str(e)})
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)


@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(REPORT_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True,
                         download_name="Purchase_Audit_Report.xlsx")
    return "File not found", 404


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(debug=False, host="0.0.0.0", port=port)
