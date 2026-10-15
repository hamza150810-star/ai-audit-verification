
import streamlit as st
import pandas as pd
import pytesseract
import fitz
import io
import re

from pathlib import Path
from PIL import Image, ImageOps, ImageEnhance
from datetime import datetime
from docx import Document

st.set_page_config(page_title="AI Audit Verification", layout="wide")
st.title("AI Audit Verification System")
st.write("Upload bills and compare them with your accounting records.")

st.warning(
    "Automated checks are preliminary audit procedures. "
    "Verify exceptions against original documents before concluding."
)

TABLE_TYPES = ["xlsx", "xls", "xlsm", "xlsb", "csv"]
BILL_TYPES = [
    "pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp",
    "docx", "txt", "xlsx", "xls", "xlsm", "xlsb", "csv"
]


def read_table(file):
    ext = Path(file.name).suffix.lower()
    data = io.BytesIO(file.getvalue())

    if ext == ".csv":
        try:
            return pd.read_csv(data, encoding="utf-8-sig")
        except UnicodeDecodeError:
            data.seek(0)
            return pd.read_csv(data, encoding="latin-1")

    if ext in [".xlsx", ".xlsm"]:
        return pd.read_excel(data, engine="openpyxl")
    if ext == ".xls":
        return pd.read_excel(data, engine="xlrd")
    if ext == ".xlsb":
        return pd.read_excel(data, engine="pyxlsb")

    raise ValueError("Unsupported spreadsheet format.")


def preprocess_image(image):
    image = ImageOps.exif_transpose(image).convert("L")
    image = ImageEnhance.Contrast(image).enhance(1.8)

    if image.width < 1500:
        factor = 1500 / image.width
        image = image.resize(
            (1500, int(image.height * factor))
        )
    return image


def get_document_text(file):
    ext = Path(file.name).suffix.lower()
    raw = file.getvalue()

    if ext == ".pdf":
        pages_text = []
        with fitz.open(stream=raw, filetype="pdf") as pdf:
            for page in pdf:
                text = page.get_text("text").strip()
                if len(text) < 20:
                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(2, 2), alpha=False
                    )
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    text = pytesseract.image_to_string(
                        preprocess_image(img)
                    )
                pages_text.append(text)
        return "\n".join(pages_text)

    if ext in [".png", ".jpg", ".jpeg", ".tif", ".tiff",
               ".bmp", ".webp"]:
        img = Image.open(io.BytesIO(raw))
        return pytesseract.image_to_string(preprocess_image(img))

    if ext == ".docx":
        doc = Document(io.BytesIO(raw))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        return "\n".join(parts)

    if ext == ".txt":
        return raw.decode("utf-8", errors="replace")

    if ext in TABLE_TYPES:
        return read_table(file).to_string(index=False)

    raise ValueError("Unsupported document format.")


def extract_fields(text, filename):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    full = "\n".join(lines)

    def first_match(pattern):
        match = re.search(pattern, full, re.IGNORECASE)
        return match.group(1).strip() if match else ""

    provider = first_match(
        r"(?:provider|company|utility|supplier)\s*[:\-]\s*([^\n]+)"
    )
    reference = first_match(
        r"(?:invoice\s*(?:no|number|#)|bill\s*(?:no|number|#)|"
        r"reference\s*(?:no|number|#)|consumer\s*(?:no|number|#))"
        r"\s*[:\-]?\s*([A-Z0-9\-\/]+)"
    )
    date_value = first_match(
        r"(?:bill date|invoice date|issue date|date)\s*[:\-]?\s*"
        r"(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}|"
        r"\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})"
    )
    amount = first_match(
        r"(?:total amount|total due|amount due|net payable|"
        r"grand total|total payable|payable amount)\s*[:\-]?\s*"
        r"(?:PKR|Rs\.?|₨)?\s*([\d,]+(?:\.\d{1,2})?)"
    )

    # If a clearly labelled amount was not found, suggest the largest
    # currency-like number. This is only a fallback and needs checking.
    if not amount:
        numbers = re.findall(
            r"(?:Rs\.?|PKR|₨)\s*([\d,]+(?:\.\d{1,2})?)",
            full,
            re.IGNORECASE
        )
        if numbers:
            amount = max(
                numbers,
                key=lambda x: float(x.replace(",", ""))
            )

    return {
        "File Name": filename,
        "Provider": provider,
        "Invoice / Reference": reference,
        "Bill Date": date_value,
        "Amount": amount
    }


def clean_text(value):
    if pd.isna(value):
        return ""
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def parse_amount(value):
    if pd.isna(value) or str(value).strip() == "":
        return None
    try:
        cleaned = re.sub(r"[^\d.\-]", "", str(value).replace(",", ""))
        return float(cleaned) if cleaned else None
    except (ValueError, TypeError):
        return None


def guess_column(columns, keywords):
    for col in columns:
        if any(word in str(col).lower() for word in keywords):
            return col
    return "-- Not available --"


def column_picker(df, label, key, keywords, required=False):
    options = ["-- Not available --"] + list(df.columns)
    guess = guess_column(df.columns, keywords)
    default_index = options.index(guess) if guess in options else 0
    return st.selectbox(
        label, options, index=default_index, key=key
    )


def match_bill(bill, df, vendor_col, ref_col, amount_col, tolerance):
    """Return a cautious matching status and candidate row."""
    if df is None or df.empty or amount_col == "-- Not available --":
        return "Unable to check", ""

    bill_amount = parse_amount(bill.get("Amount", ""))
    bill_vendor = clean_text(bill.get("Provider", ""))
    bill_ref = clean_text(bill.get("Invoice / Reference", ""))

    if bill_amount is None:
        return "Amount missing - check manually", ""

    candidates = []

    for idx, row in df.iterrows():
        row_amount = parse_amount(row.get(amount_col))
        row_vendor = (
            clean_text(row.get(vendor_col))
            if vendor_col != "-- Not available --" else ""
        )
        row_ref = (
            clean_text(row.get(ref_col))
            if ref_col != "-- Not available --" else ""
        )

        same_vendor = bool(
            bill_vendor and row_vendor and (
                bill_vendor == row_vendor
                or bill_vendor in row_vendor
                or row_vendor in bill_vendor
            )
        )
        same_ref = bool(bill_ref and row_ref and bill_ref == row_ref)
        same_amount = (
            row_amount is not None
            and abs(row_amount - bill_amount) <= tolerance
        )

        # Reference match takes priority, but amount is checked too.
        if same_ref and (same_vendor or not bill_vendor or not row_vendor):
            if same_amount:
                return "Matched", str(idx + 2)
            return "Reference found - amount mismatch", str(idx + 2)

        if same_vendor and same_amount:
            candidates.append(str(idx + 2))

    if len(candidates) == 1:
        return "Possible match - verify reference", candidates[0]
    if len(candidates) > 1:
        return "Multiple possible matches", ", ".join(candidates)

    # If the provider was not extracted, an amount alone cannot confirm
    # a match. Show this explicitly to avoid false conclusions.
    if not bill_vendor and ref_col == "-- Not available --":
        return "Not matched - provider/reference unavailable", ""

    return "Not found - investigate", ""


# ---------------- UPLOAD ACCOUNTING RECORDS ----------------

st.header("1. Upload Accounting Records")

col1, col2 = st.columns(2)

with col1:
    ledger_file = st.file_uploader(
        "General ledger", type=TABLE_TYPES, key="ledger"
    )

with col2:
    working_file = st.file_uploader(
        "Utility bills working", type=TABLE_TYPES, key="working"
    )

ledger_df = None
working_df = None

for label, file in [
    ("General ledger", ledger_file),
    ("Utility bills working", working_file)
]:
    if file:
        try:
            loaded = read_table(file)
            if label == "General ledger":
                ledger_df = loaded
            else:
                working_df = loaded
            st.success(f"{label} uploaded: {len(loaded)} rows.")
            with st.expander(f"Preview {label}"):
                st.dataframe(loaded.head(50), use_container_width=True)
        except Exception as e:
            st.error(f"Could not read {label}: {e}")


# ---------------- UPLOAD BILLS ----------------

st.header("2. Upload Bill Pictures or Documents")

bill_files = st.file_uploader(
    "Select one or more bill files",
    type=BILL_TYPES,
    accept_multiple_files=True,
    key="bill_files"
)

records = []

if bill_files:
    for i, file in enumerate(bill_files):
        with st.expander(file.name, expanded=(i == 0)):
            ext = Path(file.name).suffix.lower()
            if ext in [".png", ".jpg", ".jpeg", ".tif", ".tiff",
                       ".bmp", ".webp"]:
                st.image(file.getvalue(), use_container_width=True)

            try:
                text = get_document_text(file)
                st.text_area(
                    "Text read from document",
                    text,
                    height=150,
                    key=f"ocr_{i}"
                )
                records.append(extract_fields(text, file.name))
            except Exception as e:
                st.error(f"Could not read {file.name}: {e}")

if records:
    st.header("3. Review Extracted Bill Details")
    st.write(
        "Correct the extracted fields before running reconciliation. "
        "OCR can misread dates, invoice numbers and amounts."
    )
    bills_df = pd.DataFrame(records)
    edited_bills = st.data_editor(
        bills_df,
        num_rows="dynamic",
        use_container_width=True,
        key="edited_bills"
    )

    st.header("4. Match Bills Against Your Records")

    tolerance = st.number_input(
        "Amount difference allowed for a possible match",
        min_value=0.0,
        value=1.0,
        step=1.0
    )

    basis = st.radio(
        "Accounting basis used by the client",
        ["Accrual basis", "Cash basis", "Not sure"],
        horizontal=True
    )

    def configure_mapping(df, prefix):
        if df is None:
            return ("-- Not available --",) * 3

        st.subheader(prefix + " column mapping")
        st.caption("Choose the columns that contain these values.")
        a, b, c = st.columns(3)

        with a:
            vendor = column_picker(
                df, "Vendor / utility provider", prefix + "_vendor",
                ["vendor", "supplier", "provider", "utility", "payee",
                 "particular", "description", "name"]
            )
        with b:
            ref = column_picker(
                df, "Invoice / bill reference", prefix + "_ref",
                ["invoice", "bill no", "reference", "ref no", "consumer"]
            )
        with c:
            amount = column_picker(
                df, "Amount", prefix + "_amount",
                ["amount", "debit", "credit", "gross", "total", "value"]
            )

        return vendor, ref, amount

    if ledger_df is not None:
        ledger_vendor, ledger_ref, ledger_amount = configure_mapping(
            ledger_df, "General ledger"
        )
    else:
        ledger_vendor, ledger_ref, ledger_amount = (
            "-- Not available --",
        ) * 3

    if working_df is not None:
        work_vendor, work_ref, work_amount = configure_mapping(
            working_df, "Utility bills working"
        )
    else:
        work_vendor, work_ref, work_amount = (
            "-- Not available --",
        ) * 3

    if st.button("Run Bill Verification", type="primary"):
        results = []
        duplicate_keys = {}

        # Flag duplicate uploaded documents using provider + reference.
        for idx, bill in edited_bills.iterrows():
            provider = clean_text(bill.get("Provider", ""))
            reference = clean_text(bill.get("Invoice / Reference", ""))
            amount = parse_amount(bill.get("Amount", ""))
            date_value = clean_text(bill.get("Bill Date", ""))

            if provider and reference:
                key = ("ref", provider, reference)
            elif provider and amount is not None and date_value:
                key = ("date_amount", provider, amount, date_value)
            else:
                key = None

            duplicate_keys[idx] = key

        counts = {}
        for key in duplicate_keys.values():
            if key:
                counts[key] = counts.get(key, 0) + 1

        for idx, bill in edited_bills.iterrows():
            ledger_status, ledger_rows = match_bill(
                bill, ledger_df, ledger_vendor, ledger_ref,
                ledger_amount, tolerance
            )

            working_status, working_rows = match_bill(
                bill, working_df, work_vendor, work_ref,
                work_amount, tolerance
            )

            key = duplicate_keys[idx]
            duplicate_status = (
                "Possible duplicate upload"
                if key and counts.get(key, 0) > 1
                else "No duplicate detected"
            )

            if ledger_status == "Matched":
                overall = "Matched in ledger"
            elif "mismatch" in ledger_status.lower():
                overall = "Investigate amount mismatch"
            elif duplicate_status == "Possible duplicate upload":
                overall = "Investigate possible duplicate"
            elif ledger_status.startswith("Not found"):
                overall = "Potential unrecorded bill - investigate"
            else:
                overall = "Needs further review"

            results.append({
                **bill.to_dict(),
                "Ledger Status": ledger_status,
                "Ledger Candidate Row(s)": ledger_rows,
                "Working Status": working_status,
                "Working Candidate Row(s)": working_rows,
                "Duplicate Check": duplicate_status,
                "Overall Result": overall
            })

        results_df = pd.DataFrame(results)
        st.session_state["verification_results"] = results_df
        st.session_state["verification_basis"] = basis

    if "verification_results" in st.session_state:
        results_df = st.session_state["verification_results"]

        st.subheader("Verification Results")
        st.dataframe(results_df, use_container_width=True)

        st.download_button(
            "Download Verification Report (CSV)",
            results_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="bill_verification_report.csv",
            mime="text/csv"
        )

        st.header("5. Draft Journal Entries for Review")

        potential = results_df[
            results_df["Ledger Status"].str.startswith(
                "Not found", na=False
            )
            & (results_df["Duplicate Check"] != "Possible duplicate upload")
        ].copy()

        potential["Amount Numeric"] = potential["Amount"].apply(parse_amount)
        potential = potential[potential["Amount Numeric"].notna()]

        if potential.empty:
            st.info(
                "No bills currently meet the draft-entry criteria. "
                "This does not prove that all bills are correctly recorded."
            )
        elif st.session_state.get("verification_basis") == "Accrual basis":
            st.warning(
                "These are draft proposals only. Confirm the service period, "
                "liability, prior postings, materiality and applicable "
                "accounting framework before using any entry."
            )

            entries = []
            for _, bill in potential.iterrows():
                amount = bill["Amount Numeric"]
                entries.append({
                    "Bill": bill["File Name"],
                    "Date": bill["Bill Date"],
                    "Account Dr": "Relevant Utility Expense",
                    "Account Cr": "Accrued Utility Expenses / Payable",
                    "Debit": amount,
                    "Credit": amount,
                    "Status": "Draft - requires auditor approval"
                })

            entries_df = pd.DataFrame(entries)
            st.dataframe(entries_df, use_container_width=True)

            st.download_button(
                "Download Draft Journal Entries",
                entries_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="draft_journal_entries.csv",
                mime="text/csv"
            )
        else:
            st.info(
                "No accrual journal entries have been proposed because "
                "the selected basis is cash basis or not confirmed. "
                "Investigate payment timing and the client's applicable "
                "accounting requirements before deciding on an adjustment."
            )

st.divider()
st.caption(
    "Matching uses selected columns, reference numbers, provider names "
    "and amount tolerance. Candidate row numbers refer to spreadsheet "
    "row positions, not verified invoice IDs. A match is not proof of "
    "validity, and a non-match is not proof of an omission."
)
