"""
Receipt → Google Sheets Reimbursement Tool
=========================================

• Drag-and-drop one or many receipt images (JPEG / PNG / HEIC)  
• OCR each image with Tesseract  
• Append to the first empty template rows in the user’s Google Sheet  
  (columns A-F and I-J only — G & H formulas untouched)

2025-05-22
• Improved store name detection: case-agnostic, structure-based
"""

from __future__ import annotations

import base64, json, re, traceback
from io import BytesIO
from typing import List

import pillow_heif
pillow_heif.register_heif_opener()

import gspread, numpy as np, pytesseract, streamlit as st, cv2
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

FIRST_DATA_ROW = 19
DATE_COL = 2

# ─────────── Google-Sheets auth & helpers ────────────
def get_gsheet_client() -> gspread.Client:
    if "GCREDS_B64" in st.secrets:
        raw_json = base64.b64decode(st.secrets["GCREDS_B64"]).decode("utf-8")
        creds_info = json.loads(raw_json)
    else:
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

# ──────────────────────── OCR & Extraction ────────────────────────
def preprocess_with_cv2(pil_img: Image.Image) -> Image.Image:
    img_np = np.array(pil_img.convert("L"))
    _, binarized = cv2.threshold(img_np, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return Image.fromarray(binarized)

def guess_store_name(merged_lines: list[str]) -> str:
    def score(line, i):
        keywords_to_avoid = [
            'server', 'check', 'guest', 'amount', 'tip', 'total', 'tax',
            'visa', 'auth', 'card', 'transaction', 'subtotal', 'date', 'receipt'
        ]
        if any(kw in line.lower() for kw in keywords_to_avoid):
            return -5  # strongly penalize metadata

        word_count = len(line.split())
        if word_count == 0:
            return -10  # skip blanks

        return word_count - 0.1 * i  # favor more words, earlier in receipt

    return max(merged_lines[:15], key=lambda l: score(l, merged_lines.index(l)), default="")

def extract_receipt_data(img: Image.Image) -> dict:
    pre_img = preprocess_with_cv2(img)

    # Dual OCR passes
    text_4 = pytesseract.image_to_string(pre_img, config="--oem 3 --psm 4")
    text_11 = pytesseract.image_to_string(pre_img, config="--oem 3 --psm 11")

    # Merge and deduplicate
    lines_4 = [line.strip() for line in text_4.splitlines() if line.strip()]
    lines_11 = [line.strip() for line in text_11.splitlines() if line.strip()]
    merged_lines = list(dict.fromkeys(lines_11 + lines_4))

    store = guess_store_name(merged_lines)

    g = lambda m: m.group(1) if m else ""
    date = re.search(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", text_4)
    amt = re.search(r"\b[\\$€£]?([0-9,]+\.\d{2})\b", text_4)
    curr = re.search(r"\b(USD|EUR|GBP|JPY|CAD|AUD|INR|BRL|PEN|CNY)\b", text_4)

    return {
        "Date": g(date),
        "Description": store,
        "Expense Type": "",
        "Local Amount": g(amt).replace(",", ""),
        "Currency": g(curr),
        "Project/ Grant": "",
        "Receipt (Y/N)": "Y",
    }

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
        st.error("Please upload at least one receipt image.")
        st.stop()

    try:
        receipts = [extract_from_easyocr(load_uploaded_image(u)) for u in uploads]
        ws = open_first_worksheet(sheet_url)

        # Find first empty row (column B must be blank)
        row = FIRST_DATA_ROW
        while str(ws.cell(row, DATE_COL).value).strip() not in ("", "None", "-"):
            row += 1
        start_row = row

        # Prepare rows for columns B, C, and E (1-based)
        # Fill in None for column A and D to skip
        rows = []
        for r in receipts:
            date = r["Date"]
            desc = r["Store"]
            amt = r["Total"]
            row_data = ["", date, desc, "", amt]  # A, B, C, D, E
            rows.append(row_data)

        end_row = start_row + len(rows) - 1

        # Write B–E only
        ws.update(f"B{start_row}:E{end_row}", [r[1:5] for r in rows])

        st.success(f"Added **{len(rows)}** receipt(s) to rows {start_row}–{end_row}.")
    except ValueError as ve:
        st.error(str(ve))
    except PermissionError:
        st.error(f"🚫 Please share the Sheet with **{SERVICE_EMAIL}** as *Editor* and try again.")
    except APIError as ae:
        st.error("Google Sheets API error:"); st.text(str(ae))
    except Exception:
        st.error("Unexpected error occurred."); st.text(traceback.format_exc())

