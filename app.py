
import streamlit as st
import pandas as pd
import pytesseract
import fitz
from PIL import Image, ImageOps, ImageEnhance
from io import BytesIO
from datetime import date
import re
import os

st.set_page_config(
    page_title="AI Audit Verification",
    page_icon="📊",
    layout="wide"
)

st.title("AI Audit Verification")
st.write(
    "Upload utility bills as images, PDFs or Excel files, "
    "extract their details, and compare them with the client's ledger."
)

st.warning(
    "Extracted text may contain errors. Verify bill dates, "
    "references and amounts against the original documents."
)


def read_excel(file):
    file.seek(0)
    return pd.read_excel(file, engine="openpyxl")


def preprocess_image(image):
    image = ImageOps.exif_transpose(image)
    image = ImageOps.grayscale(image)
    image = ImageEnhance.Contrast(image).enhance(1.6)
    return image


def extract_image_text(image):
    image = preprocess_image(image)
    return pytesseract.image_to_string(
        image,
        config="--psm 6"
    )


def extract_pdf_text(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    results = []

    for page_number, page in enumerate(doc, start=1):
        text = page.get_text("text").strip()

        # Scanned PDF pages may contain no selectable text.
        if not text:
            pix = page.get_pixmap(
                matrix=fitz.Matrix(2, 2),
                alpha=False
            )
            image = Image.frombytes(
                "RGB",
                [pix.width, pix.height],
                pix.samples
            )
            text = extract_image_text(image)

        results.append(
            f"--- Page {page_number} ---\n{text}"
        )

    doc.close()
    return "\n\n".join(results)


def find_first(pattern, text):
    match = re.search(
        pattern,
        text,
        flags=re.IGNORECASE | re.MULTILINE
    )
    return match.group(1).strip() if match else ""


def extract_bill_fields(text):
    # These are suggestions only; layouts differ by provider.
    provider = find_first(
        r"(?:provider|company|utility|supplier)\s*[:\-]\s*(.+)",
        text
    )

    reference = find_first(
        r"(?:bill\s*(?:no|number)|invoice\s*(?:no|number)|"
        r"reference\s*(?:no|number)|consumer\s*(?:no|number))"
        r"\s*[:#\-]?\s*([A-Z0-9\/\-]+)",
        text
    )

    bill_date = find_first(
        r"(?:bill\s*date|issue\s*date|date)\s*[:\-]?\s*"
        r"(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4})",
        text
    )

    amount = find_first(
        r"(?:total\s*(?:amount|due|payable)|amount\s*payable|"
        r"current\s*charges|bill\s*amount)\s*[:\-]?\s*"
        r"(?:Rs\.?|PKR)?\s*([\d,]+(?:\.\d{1,2})?)",
        text
    )

    return {
        "Suggested Provider": provider,
        "Suggested Bill Reference": reference,
        "Suggested Bill Date": bill_date,
        "Suggested Amount (Rs.)": amount,
        "Manual Verification Required": "Yes"
    }


st.header("1. Upload client records")

ledger_file = st.file_uploader(
    "General ledger (.xlsx)",
    type=["xlsx"],
    key="ledger"
)

bills_working_file = st.file_uploader(
    "Utility bills working (.xlsx), if available",
    type=["xlsx"],
    key="bills_working"
)

st.header("2. Upload utility bill documents")

bill_files = st.file_uploader(
    "Select one or more bill images or PDFs",
    type=["png", "jpg", "jpeg", "pdf"],
    accept_multiple_files=True,
    key="bill_documents"
)

if bill_files:
    extracted_results = []

    for uploaded in bill_files:
        st.divider()
        st.subheader(uploaded.name)

        try:
            file_bytes = uploaded.getvalue()

            if uploaded.type == "application/pdf" or (
                uploaded.name.lower().endswith(".pdf")
            ):
                extracted_text = extract_pdf_text(file_bytes)

            else:
                image = Image.open(BytesIO(file_bytes))
                st.image(
                    image,
                    caption="Uploaded original bill",
                    use_container_width=True
                )
                extracted_text = extract_image_text(image)

            st.text_area(
                "Extracted text — check against the original",
                value=extracted_text,
                height=220,
                key="text_" + uploaded.name
            )

            fields = extract_bill_fields(extracted_text)
            fields["Source File"] = uploaded.name
            extracted_results.append(fields)

        except Exception as error:
            st.error(
                f"Could not read {uploaded.name}. "
                "Try a clearer image or another PDF."
            )
            st.exception(error)

    if extracted_results:
        st.header("3. Review extracted bill details")

        extracted_df = pd.DataFrame(extracted_results)

        edited_df = st.data_editor(
            extracted_df,
            use_container_width=True,
            num_rows="dynamic",
            key="edited_bill_details"
        )

        st.download_button(
            "Download Extracted Bill Register",
            data=edited_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="extracted_utility_bill_register.csv",
            mime="text/csv"
        )

        st.info(
            "The extracted bill details are not yet matched to the "
            "ledger in this image-upload module. Review and correct "
            "them first. The next step is to connect this register "
            "to the ledger reconciliation and journal-entry workflow."
        )

if ledger_file is not None:
    try:
        ledger_df = read_excel(ledger_file)
        st.header("4. Ledger preview")
        st.dataframe(
            ledger_df.head(20),
            use_container_width=True
        )
    except Exception as error:
        st.error("Unable to read the ledger Excel file.")
        st.exception(error)

if bills_working_file is not None:
    try:
        working_df = read_excel(bills_working_file)
        st.header("5. Utility working preview")
        st.dataframe(
            working_df.head(20),
            use_container_width=True
        )
    except Exception as error:
        st.error("Unable to read the utility working Excel file.")
        st.exception(error)
