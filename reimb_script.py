"""
Receipt → Google Sheets Reimbursement Tool
=========================================

• Drag-and-drop one or many receipt images (JPEG, PNG, HEIC).  
• OCR each image with Tesseract.  
• Append every receipt to the first empty template rows of the user’s
  Google Sheet – columns A-F and I-J only (G & H formulas untouched).

2025-05-20  
• HEIC/HEIF support via pillow-heif + libheif1  
• Robust service-account JSON parsing  
• Multi-file upload, JPEG runtime fix, larger instructions
"""

import json, re, traceback
from io import BytesIO
from typing import List

import pillow_heif               # ⬅️  NEW – HEIC opener
pillow_heif.register_heif_opener()

import gspread, numpy as np, pytesseract, streamlit as st
from PIL import Image, UnidentifiedImageError
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials

# ────────────────────────── Config ──────────────────────────
SERVICE_EMAIL = "jared-eia-reimbursements@reimbursements-460316.iam.gserviceaccount.com"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

FIRST_DATA_ROW = 19   # first template row (under header)
DATE_COL        = 2   # column B, 1-based

# ───────────── Google Sheets helpers ─────────────
def get_gsheet_client() -> gspread.Client:
    if "google_creds" in st.secrets:               # ← table variant
        creds_info = dict(st.secrets["google_creds"])
    else:                                          # fallback for local dev
        creds_info = json.load(open("credentials.json", "r"))

    creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    return gspread.authorize(creds)

def extract_sheet_id(url: str) -> str:
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", url)
    if not m:
        raise ValueError("That doesn’t look like a Google Sheets link.")
    return m.group(1)


def open_first_worksheet(url: str) -> gspread.Worksheet:
    return get_gsheet_client().open_by_key(extract_sheet_id(url)).sheet1


# ────────────── OCR helpers ──────────────
def safe_ocr(img: Image.Image) -> str:
    try:
        return pytesseract.image_to_string(img)
    except TypeError:                              # missing PIL metadata
        return pytesseract.image_to_string(np.array(img))


def load_uploaded_image(uploaded) -> Image.Image:
    try:
        data = uploaded.read()
        img = Image.open(BytesIO(data))
        img.load()                                 # force decode now
        return img
    except UnidentifiedImageError as e:
        raise ValueError(
            "Could not open that file as an image. "
            "Please upload a PNG, JPEG, or HEIC. "
            f"(Pillow error: {e})"
        ) from e
    except Exception as e:
        raise ValueError(f"Unexpected error reading image: {e}") from e


def extract_receipt_data(img: Image.Image) -> dict:
    text = safe_ocr(img)
    g = lambda m: m.group(1) if m else ""

    date  = re.search(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", text)
    amt   = re.search(r"\b([\\$€£]?[0-9,]+\.\d{2})\b",         text)
    curr  = re.search(r"\b(USD|EUR|GBP|JPY|CAD|AUD|INR|BRL|PEN|CNY)\b", text)

    return {
        "Date":          g(date),
        "Description":   text.splitlines()[0] if text.strip() else "",
        "Expense Type":  "",
        "Local Amount":  g(amt).replace("$","").replace("€","").replace("£",""),
        "Currency":      g(curr),
        "Project/ Grant": "",
        "Receipt (Y/N)":  "Y",
    }


# ───────────── Streamlit UI ─────────────
st.markdown(
    f"""
### 📎 Step 1 – Drag-and-drop one or more PNG/JPEG/HEIC receipts  
### 🔗 Step 2 – Paste the link to **your own** Google Sheet and share it with `{SERVICE_EMAIL}` as **Editor** (one-time step)  
### ✅ Step 3 – Click *Extract & Send*  
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
        receipts = [extract_receipt_data(load_uploaded_image(u)) for u in uploads]

        ws = open_first_worksheet(sheet_url)

        # find first empty template row (blank Date column)
        row = FIRST_DATA_ROW
        while str(ws.cell(row, DATE_COL).value).strip() not in ("", "None", "-"):
            row += 1
        start_row = row

        # build payloads
        rows_af, rows_ij = [], []
        for i, r in enumerate(receipts):
            receipt_no = (start_row - FIRST_DATA_ROW) + 1 + i
            rows_af.append([
                receipt_no, r["Date"], r["Description"], r["Expense Type"],
                r["Local Amount"], r["Currency"]
            ])
            rows_ij.append([r["Project/ Grant"], r["Receipt (Y/N)"]])

        end_row = start_row + len(receipts) - 1

        ws.update(f"A{start_row}:F{end_row}", rows_af)   # write A-F
        ws.update(f"I{start_row}:J{end_row}", rows_ij)   # write I-J

        st.success(f"Added **{len(receipts)}** receipt(s) to rows {start_row}-{end_row}!")

    except ValueError as ve:
        st.error(str(ve))
    except PermissionError:
        st.error(
            "🚫 I don’t have access to that Sheet yet.  "
            "Open it, click **Share**, and add\n"
            f"**{SERVICE_EMAIL}** as *Editor*, then click again."
        )
    except APIError as ae:
        st.error("Google Sheets API error:"); st.text(str(ae))
    except Exception:
        st.error("An unexpected error occurred."); st.text(traceback.format_exc())
