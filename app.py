import os, zipfile, shutil, uuid, cv2, fitz
import numpy as np
import pytesseract
import pandas as pd
from flask import Flask, request, render_template, send_file, jsonify
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from werkzeug.utils import secure_filename

# ===================== CONFIG =====================
UPLOAD_FOLDER = "data/uploads"
REPORT_FOLDER = "data/reports"
ALLOWED_BILL_EXT = (".jpg", ".jpeg", ".png", ".pdf")
IGNORE_FOLDERS   = {"payment screenshots", "payments", "misc", "receipts", "__macosx", ".ds_store"}

REQUIRED_COLS = ["Bill Date", "Bill Number", "Vendor Name", "Branch Name",
                 "SubTotal", "Item Name", "Quantity", "Rate", "Item Total"]

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB

# ===================== HELPERS ====================

def safe_str(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip().lower()

def is_ignored_folder(rel_path):
    parts = [p.strip().lower() for p in rel_path.replace("\\", "/").split("/")]
    return any(p in IGNORE_FOLDERS for p in parts)

# ===================== OCR ========================

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
    except Exception 
