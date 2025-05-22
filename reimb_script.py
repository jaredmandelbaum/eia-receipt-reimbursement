from __future__ import annotations

import base64, json, re, traceback
from io import BytesIO
from typing import List

import gspread, easyocr, numpy as np, streamlit as st
from PIL import Image, UnidentifiedImageError
from gspread.exceptions import APIError
from google.oauth2.service_account import Credentials

# ─────────────── Config ───────────────
SERVICE_EMAIL = (
    "jared-eia-reimbursements@reimbursements-460316.iam.gserviceaccount.com"
)
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
FIRST_DATA_ROW = 19
DATE_COL = 2  # column B (1-based index)

# ─────────────── Google Sheets ───────────────
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

# ─────────────── EasyOCR Setup ───────────────
reader = easyocr.Reader(['en'], gpu=False)

# ─────────────── OCR Extraction ───────────────
def extract_from_easyocr(img: Image.Image) -> dict:
    result = reader.readtext(np.array(img), detail=0)
    lines = [line.strip() for line in result if line.strip()]

    # Store name
    store = next(
        (line for line in lines[:10] if len(line) > 4 and not any(
            kw in line.lower() for kw in ["guest", "check", "server", "auth", "amount", "total", "tip", "visa", "receipt"]
        )),
        lines[0] if lines else ""
    )

    # Date
    date = ""
    for line in lines:
        match = re.search(r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b", line)
        if match:
            date = match.group(1)
            break

    # Total
    total = ""
    all_amounts = []
    for line in lines:
        if "total" in line.lower():
            match = re.search(r"([0-9]+\.\d{2})", line)
            if match:
                total = match.group(1)
        all_amounts += re.findall(r"([0-9]+\.\d{2})", line)
    if not total and all_amounts:
        total = max(all_amounts, key=lambda x: float(x))

    return {"Date": date, "Store": store, "Total": total}

def load_uploaded_image(uploaded) -> Image.Image:
    try:
        img = Image.open(BytesIO(uploaded.read()))
        img.load()
        return img
    except UnidentifiedImageError as e:
        raise ValueError(
            "Could not open that file as an image (PNG / JPEG only). "
            f"(Pillow error: {e})"
        ) from e

# ─────────────── Streamlit UI ───────────────
st.markdown(f"""
### 📎 Step 1 — Upload one or more receipt images  
### 🔗 Step 2 — Paste the link to **your own** Google Sheet and share it with `{SERVICE_EMAIL}` as **editor**  
### ✅ Step 3 — Click *Extract & Send*
""")

uploads: List["UploadedFile"] = st.file_uploader(
    "Receipt images", type=["png", "jpg", "jpeg"], accept_multiple_files=True
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

        # Build rows for B (Date), C (Store), D (blank), E (Total)
        values = []
        for r in receipts:
            values.append([r["Date"], r["Store"], "", r["Total"]])

        end_row = start_row + len(values) - 1
        ws.update(f"B{start_row}:E{end_row}", values)

        st.success(f"Added **{len(values)}** receipt(s) to rows {start_row}–{end_row}.")
    except ValueError as ve:
        st.error(str(ve))
    except PermissionError:
        st.error(f"🚫 Please share the Sheet with **{SERVICE_EMAIL}** as *Editor* and try again.")
    except APIError as ae:
        st.error("Google Sheets API error:"); st.text(str(ae))
    except Exception:
        st.error("Unexpected error occurred."); st.text(traceback.format_exc())
