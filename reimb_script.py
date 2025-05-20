"""
Receipt → Google Sheets Reimbursement Tool
=========================================

• Drag-and-drop one or many receipt images (JPEG / PNG / HEIC)  
• OCR each image with Tesseract  
• Append to the first empty template rows in the user’s Google Sheet  
  (columns A-F and I-J only — G & H formulas untouched)

2025-05-20
• Auth now reads st.secrets["GCREDS_B64"] (Base-64 JSON key)
"""

from __future__ import annotations

import base64, json, re, traceback
from io import BytesIO
from typing import List

import pillow_heif                       # HEIC/HEIF support
pillow_heif.register_heif_opener()

import gspread, numpy as np, pytesseract, streamlit as st
from PIL import Image, UnidentifiedImageError
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials

# ─────────────────────────── Config ────────────────────────────
SERVICE_EMAIL = (
    "jared-eia-reimbursements@reimbursements-460316.iam.gserviceaccount.com"
)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

FIRST_DATA_ROW = 19   # first template row (under header row)
DATE_COL        = 2   # column B (1-based index)

# ─────────── Google-Sheets auth & helpers ────────────
def get_gsheet_client() -> gspread.Client:
    """
    Authorise with a Google service-account.

    • In Streamlit Cloud the secret is stored as a **single-line Base-64 string**
      under st.secrets["GCREDS_B64"].
    • When developing locally it falls back to credentials.json.
    """
    if "GCREDS_B64" in st.secrets:                 # Cloud runtime
        raw_json = base64.b64decode(
            st.secrets["GCREDS_B64"]
        ).decode("utf-8")
        creds_info = json.loads(raw_json)
    else:                                          # local dev
        with open("credentials.json", "r", encoding="utf-8") as fh:
            creds_info = json.load(fh)

    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    return gspread.authorize(creds)


def extract_sheet_id(url: str) -> str:
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", url)
    if not m:
        raise ValueError("That doesn’t look like a Google Sheets link.")
    return m.group(1)


def open_first_worksheet(url: str) -> gspread.Worksheet:
    return get_gsheet_client().open_by_key(extract_sheet_id(url)).sheet1


# ────────────────────────── OCR helpers ──────────────────────────
from PIL import Image, ImageFilter, ImageOps

def preprocess(img: Image.Image) -> Image.Image:
    """
    • Convert to 300 DPI grayscale
    • Auto-increase contrast
    • Binarise with Otsu threshold
    • Slightly sharpen
    """
    # upscale small phone pics to ~300 DPI
    if max(img.size) < 1500:
        scale = 1500 / max(img.size)
        img = img.resize(
            (int(img.width * scale), int(img.height * scale)),
            resample=Image.LANCZOS,
        )

    gray = ImageOps.grayscale(img)
    # autocontrast flattens low-contrast scans
    gray = ImageOps.autocontrast(gray, cutoff=2)
    # PIL’s built-in point-threshold using Otsu
    thresh = gray.point(lambda x: 255 if x > ImageOps.autocontrast(gray).getextrema()[1] * 0.5 else 0)
    # gentle sharpen
    return thresh.filter(ImageFilter.SHARPEN)


def safe_ocr(img: Image.Image) -> str:
    img = preprocess(img)
    return pytesseract.image_to_string(
        img,
        config="--oem 3 --psm 6",  # LSTM engine, assume a single text block
    )

def load_uploaded_image(uploaded) -> Image.Image:
    try:
        img = Image.open(BytesIO(uploaded.read()))
        img.load()
        return img
    except UnidentifiedImageError as e:
        raise ValueError(
            "Could not open that file as an image (PNG / JPEG / HEIC only). "
            f"(Pillow error: {e})"
        ) from e


def extract_receipt_data(img: Image.Image) -> dict:
    text = safe_ocr(img)
    g = lambda m: m.group(1) if m else ""

    date  = re.search(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", text)
    amt   = re.search(r"\b([\\$€£]?[0-9,]+\.\d{2})\b", text)
    curr  = re.search(r"\b(USD|EUR|GBP|JPY|CAD|AUD|INR|BRL|PEN|CNY)\b", text)

    return {
        "Date":          g(date),
        "Description":   text.splitlines()[0] if text.strip() else "",
        "Expense Type":  "",
        "Local Amount":  g(amt).replace("$", "").replace("€", "").replace("£", ""),
        "Currency":      g(curr),
        "Project/ Grant": "",
        "Receipt (Y/N)":  "Y",
    }


# ───────────────────────── Streamlit UI ─────────────────────────
st.markdown(
    f"""
### 📎 Step 1 — Drag and drop one or more PNG/JPEG/HEIC receipt files
### 🔗 Step 2 — Paste the link to **your own** Google Sheet and share it with `{SERVICE_EMAIL}` as **editor**  
### ✅ Step 3 — Click *Extract & Send*  
""",
)

uploads: List["UploadedFile"] = st.file_uploader(
    "Receipt images", type=["png", "jpg", "jpeg", "heic", "heif"],
    accept_multiple_files=True,
)
sheet_url = st.text_input("Google Sheet URL")

if st.button("Extract & Send"):
    if not uploads:
        st.error("Please upload at least one receipt image first.")
        st.stop()

    try:
        receipts = [
            extract_receipt_data(load_uploaded_image(u)) for u in uploads
        ]
        ws = open_first_worksheet(sheet_url)

        # find first empty template row (Date column blank)
        row = FIRST_DATA_ROW
        while str(ws.cell(row, DATE_COL).value).strip() not in ("", "None", "-"):
            row += 1
        start_row = row

        rows_af, rows_ij = [], []
        for i, r in enumerate(receipts):
            receipt_no = (start_row - FIRST_DATA_ROW) + 1 + i
            rows_af.append([
                receipt_no, r["Date"], r["Description"], r["Expense Type"],
                r["Local Amount"], r["Currency"],
            ])
            rows_ij.append([r["Project/ Grant"], r["Receipt (Y/N)"]])

        end_row = start_row + len(receipts) - 1

        ws.update(f"A{start_row}:F{end_row}", rows_af)
        ws.update(f"I{start_row}:J{end_row}", rows_ij)

        st.success(
            f"Added **{len(receipts)}** receipt(s) to rows {start_row}–{end_row}."
        )

    except ValueError as ve:
        st.error(str(ve))
    except PermissionError:
        st.error(
            "🚫 I don’t have access to that Sheet yet. "
            f"Share it with **{SERVICE_EMAIL}** as *Editor* and try again."
        )
    except APIError as ae:
        st.error("Google Sheets API error:"); st.text(str(ae))
    except Exception:
        st.error("An unexpected error occurred."); st.text(traceback.format_exc())
