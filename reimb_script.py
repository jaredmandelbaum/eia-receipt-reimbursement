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

FIRST_DATA_ROW = 19   # first template row (under header row)
DATE_COL        = 2   # column B (1-based index)

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
    img_np = np.array(pil_img.convert("L"))  # grayscale
    _, binarized = cv2.threshold(img_np, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return Image.fromarray(binarized)

def guess_store_name(lines: list[str]) -> str:
    def score(line, i):
        penalty = sum(kw in line.lower() for kw in [
            'server', 'check', 'guest', 'amount', 'tip', 'total', 'tax', 'visa', 'auth'
        ])
        bonus = sum(w.istitle() for w in line.split())
        return -penalty + bonus - 0.2 * i  # earlier lines preferred

    candidates = lines[:8]
    return max(candidates, key=lambda l: score(l, candidates.index(l)), default="")

def extract_receipt_data(img: Image.Image) -> dict:
    pre_img = preprocess_with_cv2(img)
    full_text = pytesseract.image_to_string(pre_img, config="--oem 3 --psm 4")
    header_text = pytesseract.image_to_string(pre_img, config="--oem 3 --psm 11")

    lines = [line.strip() for line in header_text.splitlines() if line.strip()]
    store = guess_store_name(lines)

    g = lambda m: m.group(1) if m else ""
    date  = re.search(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", full_text)
    amt   = re.search(r"\b[\\$€£]?([0-9,]+\.\d{2})\b", full_text)
    curr  = re.search(r"\b(USD|EUR|GBP|JPY|CAD|AUD|INR|BRL|PEN|CNY)\b", full_text)

    return {
        "Date":          g(date),
        "Description":   store or (lines[0] if lines else ""),
        "Expense Type":  "",
        "Local Amount":  g(amt).replace(",", ""),
        "Currency":      g(curr),
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
