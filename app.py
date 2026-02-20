import os, zipfile, shutil, uuid, cv2, fitz
import numpy as np
import pytesseract
import pandas as pd
from flask import Flask, request, render_template, send_file, jsonify
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = "data/uploads"
REPORT_FOLDER = "data/reports"
ALLOWED_BILL_EXT = (".jpg", ".jpeg", ".png", ".pdf")
IGNORE_FOLDERS = {"payment screenshots", "payments", "misc", "receipts", "__macosx", ".ds_store"}
REQUIRED_COLS = ["Bill Date", "Bill Number", "Vendor Name", "Branch Name",
                 "SubTotal", "Item Name", "Quantity", "Rate", "Item Total"]

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

def safe_str(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip().lower()

def is_ignored_folder(rel_path):
    parts = [p.strip().lower() for p in rel_path.replace("\\", "/").split("/")]
    return any(p in IGNORE_FOLDERS for p in parts)

def preprocess(img):
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.fastNlMeansDenoising(gray, h=10)
    return cv2.adaptiveThreshold(gray, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2)

def ocr_image(path):
    try:
        img = cv2.imread(path)
        if img is None:
            return ""
        return pytesseract.image_to_string(preprocess(img), config="--psm 6").lower().strip()
    except Exception as e:
        print(f"OCR image error [{path}]: {e}")
        return ""

def ocr_pdf(path):
    text = ""
    try:
        doc = fitz.open(path)
        for page in doc:
            t = page.get_text().strip()
            if not t:
                pix = page.get_pixmap(dpi=200)
                arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                if pix.n == 4:
                    arr = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR)
                t = pytesseract.image_to_string(preprocess(arr), config="--psm 6")
            text += t + "\n"
        doc.close()
    except Exception as e:
        print(f"OCR pdf error [{path}]: {e}")
    return text.lower().strip()

def read_bill(path):
    return ocr_pdf(path) if path.lower().endswith(".pdf") else ocr_image(path)

def match_score(bill_no, vendor, item, total, text):
    return round(
        fuzz.partial_ratio(bill_no, text) * 0.40 +
        fuzz.partial_ratio(vendor, text) * 0.30 +
        fuzz.partial_ratio(item, text) * 0.20 +
        (15 if total and total in text else 0), 2
    )

GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
YELLOW = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
BLUE = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
STATUS_FILL = {"Matched": GREEN, "Not Found": YELLOW, "Duplicate": RED}

def style_report(path):
    wb = load_workbook(path)
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for cell in ws[1]:
            cell.fill = BLUE
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 22
        if sheet_name == "Audit Report":
            status_col = next(
                (c[0].column for c in ws.iter_cols(1, ws.max_column, 1, 1)
                 if c[0].value == "AI Status"), None)
            if status_col:
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    fill = STATUS_FILL.get(row[status_col - 1].value)
                    if fill:
                        for cell in row:
                            cell.fill = fill
        for col in ws.columns:
            w = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(w + 4, 55)
    wb.save(path)

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")

@app.route("/purchase_process", methods=["POST"])
def purchase_process():
    print("\n PURCHASE AUDIT STARTED\n")
    excel_file = request.files.get("excel")
    zip_file = request.files.get("zipfile")
    if not excel_file or excel_file.filename == "":
        return jsonify({"error": "Excel file is missing."}), 400
    if not zip_file or zip_file.filename == "":
        return jsonify({"error": "ZIP file is missing."}), 400
    excel_name = secure_filename(excel_file.filename)
    zip_name = secure_filename(zip_file.filename)
    if not excel_name.lower().endswith((".xlsx", ".xls", ".csv")):
        return jsonify({"error": "Excel must be .xlsx / .xls / .csv"}), 400
    if not zip_name.lower().endswith(".zip"):
        return jsonify({"error": "Bills file must be a .zip"}), 400
    sid = uuid.uuid4().hex[:10]
    sess_dir = os.path.join(UPLOAD_FOLDER, sid)
    os.makedirs(sess_dir, exist_ok=True)
    excel_path = os.path.join(sess_dir, excel_name)
    zip_path = os.path.join(sess_dir, zip_name)
    excel_file.save(excel_path)
    zip_file.save(zip_path)
    try:
        df = (pd.read_csv(excel_path, dtype=str)
              if excel_path.lower().endswith(".csv")
              else pd.read_excel(excel_path, engine="openpyxl", dtype=str))
    except Exception as e:
        return jsonify({"error": f"Cannot read Excel: {e}"}), 400
    df.columns = df.columns.str.strip()
    df = df.fillna("")
    if df.empty:
        return jsonify({"error": "Excel file has no data rows."}), 400
    for col in REQUIRED_COLS:
        if col not in df.columns:
            df[col] = ""
    df = df[REQUIRED_COLS].copy()
    extract_dir = os.path.join(sess_dir, "bills")
    os.makedirs(extract_dir, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_path, "r") as z:
            z.extractall(extract_dir)
    except zipfile.BadZipFile:
        return jsonify({"error": "ZIP file is corrupted or invalid."}), 400
    bill_files = []
    for root, dirs, files in os.walk(extract_dir):
        rel = os.path.relpath(root, extract_dir)
        if is_ignored_folder(rel):
            dirs.clear()
            continue
        for f in files:
            if f.lower().endswith(ALLOWED_BILL_EXT) and not f.startswith("."):
                bill_files.append(os.path.join(root, f))
    total_in_zip = len(bill_files)
    if total_in_zip == 0:
        return jsonify({"error": "No bill files found in ZIP."}), 400
    ocr_results = []
    for i, path in enumerate(bill_files, 1):
        name = os.path.basename(path)
        print(f"OCR [{i}/{total_in_zip}] {name}")
        text = read_bill(path)
        if text.strip():
            ocr_results.append((path, text))
        print(f"Progress: {int(i / total_in_zip * 100)}%")
    raw = []
    bill_seen = {}
    for idx, row in df.iterrows():
        bn = safe_str(row["Bill Number"])
        ven = safe_str(row["Vendor Name"])
        itm = safe_str(row["Item Name"])
        tot = safe_str(row["Item Total"])
        best, best_path = 0, ""
        for path, txt in ocr_results:
            s = match_score(bn, ven, itm, tot, txt)
            if s > best:
                best, best_path = s, path
        raw.append({"bill_no": bn, "score": best, "path": best_path})
        if bn:
            bill_seen.setdefault(bn, []).append(idx)
    statuses, remarks, scores, sources = [], [], [], []
    matched_paths = set()
    for r in raw:
        bn, score, path = r["bill_no"], r["score"], r["path"]
        is_dup = bn and len(bill_seen.get(bn, [])) > 1
        if is_dup:
            statuses.append("Duplicate")
            remarks.append(f"Bill '{bn}' appears multiple times in Excel")
            sources.append(os.path.basename(path) if path else "")
            scores.append(score)
        elif score >= 82:
            statuses.append("Matched")
            remarks.append("Key fields matched (Bill No, Vendor, Item, Amount)")
            sources.append(os.path.basename(path))
            scores.append(score)
            matched_paths.add(path)
        else:
            statuses.append("Not Found")
            remarks.append("No matching bill image found in ZIP")
            sources.append("")
            scores.append(score)
    df["AI Status"] = statuses
    df["AI Remark"] = remarks
    df["Match Score"] = scores
    df["Source File"] = sources
    t_excel = len(df)
    t_matched = statuses.count("Matched")
    t_notfound = statuses.count("Not Found")
    t_duplicate = statuses.count("Duplicate")
    t_zip = total_in_zip
    t_unaccounted = t_zip - len(matched_paths)
    report_name = f"Purchase_Audit_{sid}.xlsx"
    report_path = os.path.join(REPORT_FOLDER, report_name)
    with pd.ExcelWriter(report_path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Audit Report", index=False)
        pd.DataFrame({
            "Metric": [
                "Total Bills in Excel",
                "Matched",
                "Not Found",
                "Duplicates",
                "Total Bills in ZIP",
                "ZIP Bills not in Excel",
            ],
            "Count": [t_excel, t_matched, t_notfound, t_duplicate, t_zip, t_unaccounted]
        }).to_excel(w, sheet_name="Summary", index=False)
    style_report(report_path)
    try:
        shutil.rmtree(sess_dir)
    except Exception:
        pass
    return send_file(
        report_path, as_attachment=True,
        download_name="Purchase_Audit_Report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True, port=8000)
