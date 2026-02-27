import os, zipfile, shutil, uuid, re, threading
import cv2, pytesseract, fitz
import pandas as pd
from flask import Flask, request, render_template_string, send_file, jsonify
from rapidfuzz import fuzz
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

app = Flask(__name__)

UPLOAD_FOLDER = "/tmp/uploads"
REPORT_FOLDER = "/tmp/reports"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

ALLOWED_BILL_EXT = {".jpg", ".jpeg", ".png", ".pdf"}
IGNORE_FOLDERS = {
    "payment screenshots", "payments", "misc",
    "receipts", "__macosx", ".ds_store", "payment"
}

# In-memory job store
jobs = {}

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
    .card{background:#1e293b;border-radius:16px;padding:30px;border:1px solid #334155;margin-bottom:20px}
    .card h2{font-size:18px;font-weight:600;margin-bottom:20px;color:#f1f5f9}
    .upload-area{border:2px dashed #334155;border-radius:12px;padding:50px;
                 text-align:center;cursor:pointer;transition:all .3s}
    .upload-area:hover,.upload-area.dragover{border-color:#3b82f6;background:rgba(59,130,246,.05)}
    .upload-area .icon{font-size:48px;margin-bottom:15px}
    .upload-area p{color:#94a3b8;margin-bottom:5px}
    .upload-area strong{color:#3b82f6}
    .btn{background:#3b82f6;color:#fff;border:none;padding:12px 30px;border-radius:8px;
         font-size:15px;font-weight:600;cursor:pointer;width:100%;margin-top:15px}
    .btn:hover{background:#2563eb}
    .btn:disabled{background:#475569;cursor:not-allowed}
    .progress-area{display:none}
    .progress-bar-wrap{background:#0f172a;border-radius:8px;height:8px;margin:15px 0;overflow:hidden}
    .progress-bar{background:linear-gradient(90deg,#3b82f6,#8b5cf6);height:100%;
                  width:0%;transition:width .6s;border-radius:8px}
    .status-text{color:#94a3b8;font-size:14px;text-align:center;margin-bottom:10px}
    .result-area{display:none}
    .stats{display:grid;grid-template-columns:repeat(3,1fr);gap:15px;margin-bottom:20px}
    .stat{background:#0f172a;border-radius:10px;padding:20px;text-align:center;border:1px solid #334155}
    .stat .num{font-size:32px;font-weight:700;color:#3b82f6}
    .stat .label{font-size:13px;color:#94a3b8;margin-top:5px}
    .download-btn{background:linear-gradient(135deg,#10b981,#059669);color:#fff;border:none;
                  padding:15px 30px;border-radius:10px;font-size:16px;font-weight:600;
                  cursor:pointer;width:100%}
    .download-btn:hover{transform:translateY(-2px)}
    .log{background:#0f172a;border-radius:8px;padding:15px;max-height:280px;overflow-y:auto;
         font-family:monospace;font-size:12px;margin-top:15px;line-height:1.6}
    .log-item{padding:2px 0;border-bottom:1px solid #1a2035}
    .log-item.ok{color:#10b981}
    .log-item.err{color:#ef4444}
    .log-item.info{color:#94a3b8}
    input[type="file"]{display:none}
    .powered{text-align:center;color:#475569;font-size:12px;margin-top:30px;padding-bottom:30px}
    .csv-upload{border:1px dashed #334155;border-radius:8px;padding:15px;text-align:center;
                cursor:pointer;margin-top:10px;transition:all .3s}
    .csv-upload:hover{border-color:#3b82f6}
    .err-box{background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.3);
             border-radius:8px;padding:14px;color:#ef4444;display:none;margin-top:12px;font-size:14px}
  </style>
</head>
<body>
  <div class="header">
    <h1>Purchase Audit System</h1>
    <span class="badge">OCR + Smart Matching</span>
  </div>
  <div class="container">
    <div class="card">
      <h2>Upload Files</h2>
      <div class="upload-area" id="dropZone"
           onclick="document.getElementById('fileInput').click()">
        <div class="icon">&#x1F4E6;</div>
        <p><strong>Click to upload</strong> or drag and drop</p>
        <p>ZIP file containing bill images (JPG, PNG, PDF)</p>
        <p id="fileName" style="color:#3b82f6;margin-top:10px;font-weight:600;"></p>
      </div>
      <input type="file" id="fileInput" accept=".zip" onchange="handleFile(this)">

      <label style="color:#94a3b8;font-size:14px;display:block;margin-top:15px;">
        Zoho Excel or CSV (for matching):
      </label>
      <div class="csv-upload" onclick="document.getElementById('csvInput').click()">
        <p style="color:#94a3b8;"><strong style="color:#3b82f6;">Click</strong> to upload Excel or CSV</p>
        <p id="csvName" style="color:#3b82f6;font-size:13px;margin-top:5px;"></p>
      </div>
      <input type="file" id="csvInput" accept=".csv,.xlsx,.xls"
             onchange="document.getElementById('csvName').textContent='OK: '+(this.files[0]?.name||'')">

      <button class="btn" id="uploadBtn" onclick="startAudit()" disabled>Start Audit</button>
      <div class="err-box" id="errBox"></div>
    </div>

    <div class="card progress-area" id="progressCard">
      <h2>Processing Bills...</h2>
      <div class="progress-bar-wrap"><div class="progress-bar" id="progressBar"></div></div>
      <p class="status-text" id="statusText">Starting...</p>
      <div class="log" id="logArea"></div>
    </div>

    <div class="card result-area" id="resultCard">
      <h2>Audit Complete!</h2>
      <div class="stats">
        <div class="stat"><div class="num" id="statTotal">0</div><div class="label">Bills Processed</div></div>
        <div class="stat"><div class="num" id="statMatched" style="color:#10b981">0</div><div class="label">Matched</div></div>
        <div class="stat"><div class="num" id="statMismatch" style="color:#ef4444">0</div><div class="label">Not Found / Mismatch</div></div>
      </div>
      <button class="download-btn" id="dlBtn">Download Excel Report</button>
    </div>

    <div class="powered">Purchase Audit System - Built by Shubham Hulsure</div>
  </div>

  <script>
    let zipFile = null, currentJobId = null, pollTimer = null;

    function handleFile(input) {
      zipFile = input.files[0];
      document.getElementById('fileName').textContent = zipFile ? 'Selected: ' + zipFile.name : '';
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
        document.getElementById('fileName').textContent = 'Selected: ' + f.name;
        document.getElementById('uploadBtn').disabled = false;
      }
    });

    function addLog(msg, type) {
      const log = document.getElementById('logArea');
      const div = document.createElement('div');
      div.className = 'log-item ' + (type || 'info');
      div.textContent = msg;
      log.appendChild(div);
      log.scrollTop = log.scrollHeight;
    }

    function showErr(msg) {
      const b = document.getElementById('errBox');
      b.textContent = 'Error: ' + msg;
      b.style.display = 'block';
    }

    async function startAudit() {
      if (!zipFile) return;
      document.getElementById('errBox').style.display = 'none';
      document.getElementById('logArea').innerHTML = '';
      document.getElementById('resultCard').style.display = 'none';

      const btn = document.getElementById('uploadBtn');
      btn.disabled = true; btn.textContent = 'Uploading...';
      document.getElementById('progressCard').style.display = 'block';
      document.getElementById('progressBar').style.width = '5%';
      document.getElementById('statusText').textContent = 'Uploading files...';

      const fd = new FormData();
      fd.append('zip_file', zipFile);
      const csvFile = document.getElementById('csvInput').files[0];
      if (csvFile) fd.append('csv_file', csvFile);

      try {
        const res  = await fetch('/start', { method: 'POST', body: fd });
        const data = await res.json();
        if (!data.job_id) { showErr(data.error || 'Upload failed'); btn.disabled = false; btn.textContent = 'Start Audit'; return; }
        currentJobId = data.job_id;
        btn.textContent = 'Processing...';
        addLog('Files uploaded. OCR started in background...', 'ok');
        pollTimer = setInterval(pollStatus, 2500);
      } catch(e) {
        showErr(e.message); btn.disabled = false; btn.textContent = 'Start Audit';
      }
    }

    async function pollStatus() {
      if (!currentJobId) return;
      try {
        const res = await fetch('/status/' + currentJobId);
        const job = await res.json();

        document.getElementById('progressBar').style.width  = job.progress + '%';
        document.getElementById('statusText').textContent   = job.step || 'Processing...';

        if (job.new_logs && job.new_logs.length) {
          job.new_logs.forEach(l => addLog(l.msg, l.type));
        }

        if (job.status === 'done') {
          clearInterval(pollTimer);
          document.getElementById('statTotal').textContent    = job.summary.total;
          document.getElementById('statMatched').textContent  = job.summary.matched;
          document.getElementById('statMismatch').textContent = job.summary.mismatch;
          document.getElementById('resultCard').style.display = 'block';
          document.getElementById('dlBtn').onclick = () => { window.location.href = '/download/' + currentJobId; };
          document.getElementById('uploadBtn').disabled = false;
          document.getElementById('uploadBtn').textContent = 'Start Audit';
        }

        if (job.status === 'error') {
          clearInterval(pollTimer);
          showErr(job.error || 'Something went wrong.');
          document.getElementById('uploadBtn').disabled = false;
          document.getElementById('uploadBtn').textContent = 'Start Audit';
        }
      } catch(e) { /* network blip — keep polling */ }
    }
  </script>
</body>
</html>
"""

# ── Helpers ────────────────────────────────────────────────────────────────────

def is_ignored(path):
    parts = [p.strip().lower() for p in path.replace("\\", "/").split("/")]
    return any(p in IGNORE_FOLDERS for p in parts)

def normalize(text):
    if not text: return ""
    text = re.sub(r"[^a-z0-9\s\.]", " ", str(text).lower())
    return re.sub(r"\s+", " ", text).strip()

def extract_numbers(text):
    return set(re.findall(r"\b\d+(?:\.\d+)?\b", text))

def pdf_to_images(pdf_path, output_dir):
    images = []
    try:
        doc = fitz.open(pdf_path)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_path = os.path.join(output_dir, f"pdf_page_{i}.jpg")
            pix.save(img_path); images.append(img_path)
    except Exception as e:
        print(f"PDF error: {e}")
    return images

def ocr_image(path):
    try:
        img = cv2.imread(path)
        if img is None: return ""
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        _, proc = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        raw = pytesseract.image_to_string(proc, config="--psm 6")
        return normalize(raw)
    except Exception as e:
        print(f"OCR error {path}: {e}"); return ""

def score_match(bill_no, vendor, item, amount_str, ocr_text):
    if not ocr_text: return 0
    score = 0
    if bill_no and len(bill_no) >= 3:
        score += 60 if bill_no in ocr_text else fuzz.partial_ratio(bill_no, ocr_text) * 0.30
    if vendor and len(vendor) >= 3:
        score += fuzz.partial_ratio(vendor, ocr_text) * 0.20
    if item and len(item) >= 2:
        score += fuzz.partial_ratio(item, ocr_text) * 0.10
    if amount_str:
        amt = re.sub(r"\.0+$", "", amount_str.strip())
        if amt in extract_numbers(ocr_text) or amount_str.strip() in extract_numbers(ocr_text):
            score += 10
    return min(round(score, 1), 100)

def classify(score):
    if score >= 70:   return "Matched",              "Strong match"
    elif score >= 45: return "Mismatch / Duplicate", "Partial match - verify manually"
    else:             return "Not Found",             "No matching bill image detected"

def generate_excel(results, output_path):
    wb = Workbook(); ws = wb.active; ws.title = "Audit Report"
    headers = ["File Name","Folder","Bill Number","Bill Date","Vendor Name",
               "Customer/Hotel","Item Description","Quantity","Rate (Rs)",
               "Total Amount (Rs)","AI Confidence","Match Status"]
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
        row_data = [r.get("file_name",""), r.get("folder",""), r.get("bill_number",""),
                    r.get("bill_date",""), r.get("vendor_name",""), r.get("customer_name",""),
                    r.get("item_description",""), r.get("quantity",""), r.get("rate",""),
                    r.get("total_amount",""), r.get("confidence",""), r.get("match_status","")]
        status = r.get("match_status","")
        if   status == "Matched":         row_fill = green
        elif status == "Not Found":       row_fill = red
        elif "Mismatch" in status:        row_fill = yellow
        else:                             row_fill = alt if i % 2 == 0 else None
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.alignment = Alignment(vertical="center")
            if row_fill: cell.fill = row_fill
    col_widths = [22,18,15,12,30,25,25,10,12,15,12,14]
    for col, width in enumerate(col_widths, 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = width
    ws.freeze_panes = "A2"
    wb.save(output_path)

# ── Background worker ──────────────────────────────────────────────────────────

def run_audit(job_id, work_dir, zip_path, csv_path):
    job = jobs[job_id]

    def log(msg, t="info"):
        job["all_logs"].append({"msg": msg, "type": t})
        job["new_logs"].append({"msg": msg, "type": t})

    def update(progress, step):
        job["progress"] = progress
        job["step"] = step

    try:
        # 1 — Extract ZIP
        update(8, "Extracting ZIP file...")
        extract_dir = os.path.join(work_dir, "extracted")
        os.makedirs(extract_dir, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(extract_dir)

        # 2 — Read reference file
        csv_df = None
        if csv_path and os.path.exists(csv_path):
            update(12, "Reading Excel / CSV data...")
            try:
                csv_df = pd.read_csv(csv_path) if csv_path.endswith(".csv") \
                         else pd.read_excel(csv_path, engine="openpyxl")
                csv_df.columns = csv_df.columns.str.strip()
                log(f"Loaded reference file: {len(csv_df)} rows", "ok")
            except Exception as e:
                log(f"Could not read reference file: {e}", "err")

        # 3 — Collect images
        update(16, "Scanning for bill images...")
        bill_images = []
        for root, dirs, files in os.walk(extract_dir):
            for f in files:
                rel = os.path.relpath(os.path.join(root, f), extract_dir)
                full = os.path.join(root, f)
                if is_ignored(rel): continue
                ext = os.path.splitext(f)[1].lower()
                if ext == ".pdf":
                    bill_images.extend([(img, rel) for img in pdf_to_images(full, root)])
                elif ext in ALLOWED_BILL_EXT:
                    bill_images.append((full, rel))

        total_imgs = len(bill_images)
        log(f"Found {total_imgs} bill files in ZIP", "info")

        # 4 — OCR
        ocr_cache = []
        for i, (img_path, rel_path) in enumerate(bill_images):
            update(16 + int((i / max(total_imgs, 1)) * 50),
                   f"OCR scanning {i+1}/{total_imgs}: {os.path.basename(rel_path)}")
            txt = ocr_image(img_path)
            if txt:
                ocr_cache.append((img_path, txt))
                log(f"OCR OK: {os.path.basename(rel_path)}", "ok")
            else:
                log(f"OCR empty: {os.path.basename(rel_path)}", "err")
            job["new_logs"] = job["new_logs"][-30:]  # keep buffer lean

        # 5 — Match
        results = []
        if csv_df is not None:
            update(68, "Matching Excel rows to bills...")
            log("Matching Excel rows to bill images...", "info")
            total_rows = len(csv_df)
            for idx, (_, row) in enumerate(csv_df.iterrows()):
                update(68 + int((idx / max(total_rows, 1)) * 22),
                       f"Matching row {idx+1}/{total_rows}...")
                bill_no = normalize(str(row.get("Bill Number", "")))
                vendor  = normalize(str(row.get("Vendor Name", "")))
                item    = normalize(str(row.get("Item Name", "")))
                amount  = normalize(str(row.get("Item Total", "")))
                best_score = 0; best_path = ""
                for path, text in ocr_cache:
                    s = score_match(bill_no, vendor, item, amount, text)
                    if s > best_score: best_score = s; best_path = path
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
                t = "ok" if status == "Matched" else ("err" if status == "Not Found" else "info")
                log(f"{status} | {row.get('Bill Number','?')} | {str(row.get('Vendor Name',''))[:20]} | score {best_score}", t)
        else:
            for img_path, _ in ocr_cache:
                results.append({"file_name": os.path.basename(img_path),
                                 "folder":"","bill_number":"","bill_date":"",
                                 "vendor_name":"","customer_name":"","item_description":"",
                                 "quantity":"","rate":"","total_amount":"",
                                 "confidence":"","match_status":"No CSV"})

        # 6 — Write report
        update(92, "Writing Excel report...")
        report_path = os.path.join(REPORT_FOLDER, f"audit_{job_id}.xlsx")
        generate_excel(results, report_path)

        matched  = sum(1 for r in results if r["match_status"] == "Matched")
        mismatch = sum(1 for r in results if r["match_status"] in ("Not Found", "Mismatch / Duplicate"))
        log(f"Done! {matched} matched, {mismatch} issues out of {len(results)} rows", "ok")

        job.update({
            "status": "done", "progress": 100, "step": "Complete!",
            "report_path": report_path,
            "summary": {"total": len(results), "matched": matched, "mismatch": mismatch}
        })

    except Exception as e:
        job.update({"status": "error", "error": str(e)})
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/start", methods=["POST"])
def start():
    job_id   = str(uuid.uuid4())[:8]
    work_dir = os.path.join(UPLOAD_FOLDER, job_id)
    os.makedirs(work_dir, exist_ok=True)

    zip_file = request.files.get("zip_file")
    if not zip_file:
        return jsonify({"error": "No ZIP file provided"}), 400

    zip_path = os.path.join(work_dir, "bills.zip")
    zip_file.save(zip_path)

    csv_path = None
    csv_file = request.files.get("csv_file")
    if csv_file:
        ext      = os.path.splitext(csv_file.filename)[1].lower() or ".xlsx"
        csv_path = os.path.join(work_dir, f"data{ext}")
        csv_file.save(csv_path)

    jobs[job_id] = {
        "status": "processing", "progress": 5, "step": "Files received...",
        "all_logs": [], "new_logs": [], "summary": {}, "report_path": None, "error": None
    }

    t = threading.Thread(target=run_audit, args=(job_id, work_dir, zip_path, csv_path))
    t.daemon = True
    t.start()

    return jsonify({"job_id": job_id})

@app.route("/status/<job_id>")
def status(job_id):
    if job_id not in jobs:
        return jsonify({"status": "not_found"}), 404
    job = jobs[job_id]
    out = {k: job[k] for k in ("status","progress","step","summary","error")}
    out["new_logs"] = job["new_logs"]
    job["new_logs"] = []          # flush — only send each log line once
    return jsonify(out)

@app.route("/download/<job_id>")
def download(job_id):
    if job_id not in jobs or not jobs[job_id].get("report_path"):
        return "Report not ready", 404
    rp = jobs[job_id]["report_path"]
    if os.path.exists(rp):
        return send_file(rp, as_attachment=True, download_name="Purchase_Audit_Report.xlsx")
    return "File not found", 404

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    app.run(debug=False, host="0.0.0.0", port=port)
