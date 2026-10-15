
import streamlit as st
import pandas as pd
import pytesseract
import fitz
import io
import json
import re

from pathlib import Path
from PIL import Image, ImageOps, ImageEnhance
from docx import Document


st.set_page_config(
    page_title="AI Audit Verification",
    page_icon="📊",
    layout="wide"
)

st.title("AI Audit Verification System")
st.write(
    "Upload accounting records and supporting documents "
    "for audit assistance."
)

st.warning(
    "This application assists with audit verification. "
    "Always verify extracted information against original evidence."
)

# Common supported formats
TABLE_FORMATS = ["xlsx", "xls", "xlsm", "xlsb", "csv"]
DOCUMENT_FORMATS = [
    "pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp", "webp",
    "xlsx", "xls", "xlsm", "xlsb", "csv",
    "docx", "txt", "json"
]


def read_table(uploaded_file):
    """Read common spreadsheet and CSV formats."""
    extension = Path(uploaded_file.name).suffix.lower()
    file_bytes = uploaded_file.getvalue()

    if extension == ".csv":
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(
                io.BytesIO(file_bytes), encoding="latin-1"
            )

    if extension in [".xlsx", ".xlsm"]:
        return pd.read_excel(
            io.BytesIO(file_bytes), engine="openpyxl"
        )

    if extension == ".xls":
        return pd.read_excel(
            io.BytesIO(file_bytes), engine="xlrd"
        )

    if extension == ".xlsb":
        return pd.read_excel(
            io.BytesIO(file_bytes), engine="pyxlsb"
        )

    raise ValueError(
        "Unsupported spreadsheet format. Upload Excel or CSV."
    )


def preprocess_image(image):
    """Improve image readability before OCR."""
    image = ImageOps.exif_transpose(image)
    image = image.convert("RGB")
    image = ImageOps.grayscale(image)
    image = ImageEnhance.Contrast(image).enhance(1.8)

    # Enlarge small images to help OCR.
    if image.width < 1500:
        ratio = 1500 / image.width
        image = image.resize(
            (1500, int(image.height * ratio))
        )

    return image


def extract_pdf_text(file_bytes):
    """Read selectable PDF text or use OCR for scanned pages."""
    text_parts = []

    with fitz.open(stream=file_bytes, filetype="pdf") as pdf:
        for page_number, page in enumerate(pdf, start=1):
            page_text = page.get_text("text").strip()

            if len(page_text) < 20:
                pix = page.get_pixmap(
                    matrix=fitz.Matrix(2, 2),
                    alpha=False
                )
                image = Image.open(
                    io.BytesIO(pix.tobytes("png"))
                )
                page_text = pytesseract.image_to_string(
                    preprocess_image(image)
                )

            text_parts.append(
                f"--- Page {page_number} ---\n{page_text}"
            )

    return "\n\n".join(text_parts)


def extract_document_text(uploaded_file):
    """Extract text from common audit document formats."""
    extension = Path(uploaded_file.name).suffix.lower()
    file_bytes = uploaded_file.getvalue()

    if extension == ".pdf":
        return extract_pdf_text(file_bytes)

    if extension in [
        ".png", ".jpg", ".jpeg", ".tif", ".tiff",
        ".bmp", ".webp"
    ]:
        image = Image.open(io.BytesIO(file_bytes))
        return pytesseract.image_to_string(
            preprocess_image(image)
        )

    if extension == ".docx":
        document = Document(io.BytesIO(file_bytes))
        paragraphs = [
            paragraph.text for paragraph in document.paragraphs
            if paragraph.text.strip()
        ]

        for table in document.tables:
            for row in table.rows:
                paragraphs.append(
                    " | ".join(cell.text for cell in row.cells)
                )

        return "\n".join(paragraphs)

    if extension == ".txt":
        try:
            return file_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            return file_bytes.decode("latin-1")

    if extension == ".json":
        return json.dumps(
            json.loads(file_bytes.decode("utf-8-sig")),
            indent=2,
            ensure_ascii=False
        )

    if extension in TABLE_FORMATS:
        dataframe = read_table(uploaded_file)
        return dataframe.to_string(index=False)

    raise ValueError("This file format is not supported.")


def extract_bill_fields(text, filename):
    """Suggest bill details using simple text patterns."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    combined_text = "\n".join(lines)

    provider = ""
    invoice_number = ""
    bill_date = ""
    amount = ""

    provider_pattern = re.compile(
        r"(?:provider|company|utility|supplier)\s*[:\-]\s*(.+)",
        re.IGNORECASE
    )
    invoice_pattern = re.compile(
        r"(?:invoice\s*(?:no|number|#)|bill\s*(?:no|number|#)|"
        r"reference\s*(?:no|number|#))\s*[:\-]?\s*([A-Z0-9\-\/]+)",
        re.IGNORECASE
    )
    date_pattern = re.compile(
        r"\b(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}|"
        r"\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})\b"
    )
    amount_pattern = re.compile(
        r"(?:total\s*(?:amount|due|payable)?|amount\s*due|"
        r"net\s*payable|grand\s*total)\s*[:\-]?\s*"
        r"(?:PKR|Rs\.?|₨)?\s*([\d,]+(?:\.\d{1,2})?)",
        re.IGNORECASE
    )

    for line in lines:
        if not provider:
            match = provider_pattern.search(line)
            if match:
                provider = match.group(1).strip()

        if not invoice_number:
            match = invoice_pattern.search(line)
            if match:
                invoice_number = match.group(1).strip()

        if not bill_date:
            match = date_pattern.search(line)
            if match:
                bill_date = match.group(1)

        if not amount:
            match = amount_pattern.search(line)
            if match:
                amount = match.group(1).replace(",", "")

    return {
        "File Name": filename,
        "Provider": provider,
        "Invoice / Reference": invoice_number,
        "Bill Date": bill_date,
        "Amount": amount,
        "Review Status": "Needs manual verification"
    }


# --------------------------------------------------
# SECTION 1: GENERAL LEDGER
# --------------------------------------------------

st.header("1. Upload General Ledger")

ledger_file = st.file_uploader(
    "Select your general ledger",
    type=TABLE_FORMATS,
    key="ledger"
)

if ledger_file:
    try:
        ledger_df = read_table(ledger_file)
        st.success("General ledger uploaded successfully.")
        st.write("Ledger preview:")
        st.dataframe(ledger_df.head(100), use_container_width=True)
    except Exception as error:
        st.error(f"Could not read the ledger: {error}")


# --------------------------------------------------
# SECTION 2: UTILITY BILL WORKING
# --------------------------------------------------

st.header("2. Upload Utility Bills Working")

working_file = st.file_uploader(
    "Select your utility bills working",
    type=TABLE_FORMATS,
    key="working"
)

if working_file:
    try:
        working_df = read_table(working_file)
        st.success("Utility working uploaded successfully.")
        st.write("Utility working preview:")
        st.dataframe(
            working_df.head(100),
            use_container_width=True
        )
    except Exception as error:
        st.error(f"Could not read the utility working: {error}")


# --------------------------------------------------
# SECTION 3: SUPPORTING DOCUMENTS
# --------------------------------------------------

st.header("3. Upload Supporting Bills and Documents")

st.write(
    "Upload multiple bills together. PDF files, photographs, "
    "scanned bills, Word documents, text files and spreadsheets "
    "are supported."
)

bill_files = st.file_uploader(
    "Choose bills or supporting documents",
    type=DOCUMENT_FORMATS,
    accept_multiple_files=True,
    key="bills"
)

bill_records = []

if bill_files:
    st.info(f"{len(bill_files)} file(s) uploaded.")

    for index, uploaded_file in enumerate(bill_files):
        with st.expander(
            f"Document {index + 1}: {uploaded_file.name}",
            expanded=(index == 0)
        ):
            extension = Path(uploaded_file.name).suffix.lower()

            if extension in [
                ".png", ".jpg", ".jpeg", ".tif", ".tiff",
                ".bmp", ".webp"
            ]:
                st.image(
                    uploaded_file.getvalue(),
                    caption=uploaded_file.name,
                    use_container_width=True
                )

            try:
                extracted_text = extract_document_text(uploaded_file)

                st.text_area(
                    "Extracted text — check it against the original",
                    value=extracted_text,
                    height=220,
                    key=f"text_{index}"
                )

                record = extract_bill_fields(
                    extracted_text, uploaded_file.name
                )
                bill_records.append(record)

                st.write("Suggested bill details")
                st.json(record)

            except Exception as error:
                st.error(
                    f"Could not process this file: {error}"
                )

if bill_records:
    st.header("4. Review Extracted Bill Register")

    register_df = pd.DataFrame(bill_records)

    st.write(
        "Correct any missing or inaccurate values before using "
        "this register for audit work."
    )

    edited_df = st.data_editor(
        register_df,
        num_rows="dynamic",
        use_container_width=True,
        key="bill_register"
    )

    csv_data = edited_df.to_csv(index=False).encode("utf-8-sig")

    st.download_button(
        "Download Bill Register (CSV)",
        data=csv_data,
        file_name="extracted_bill_register.csv",
        mime="text/csv"
    )

st.divider()
st.caption(
    "Note: OCR and automatic field extraction may be inaccurate. "
    "This version uploads and extracts documents; it does not yet "
    "automatically reconcile every extracted bill against the ledger "
    "or generate final journal entries."
)
