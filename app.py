
import io
import re
import random
from datetime import datetime

import pandas as pd
import streamlit as st
from openai import OpenAI

st.set_page_config(
    page_title="AI Audit Verification",
    page_icon="📊",
    layout="wide"
)

st.title("AI Audit Verification System")
st.write(
    "Upload accounting records, select expense samples, "
    "and prepare AI-assisted audit working papers."
)


# ==========================================================
# 1. FILE READING HELPERS
# ==========================================================

def read_tabular_file(uploaded_file, sheet_name=None):
    """Read an uploaded Excel or CSV file safely."""
    if uploaded_file is None:
        return None

    file_bytes = uploaded_file.getvalue()
    filename = uploaded_file.name.lower()

    if filename.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes))

    if filename.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(
            io.BytesIO(file_bytes),
            sheet_name=sheet_name if sheet_name is not None else 0
        )

    raise ValueError("Please upload an Excel or CSV file.")


def get_excel_sheets(uploaded_file):
    """Return worksheet names for Excel files."""
    if uploaded_file is None:
        return []

    if uploaded_file.name.lower().endswith(".csv"):
        return []

    workbook = pd.ExcelFile(
        io.BytesIO(uploaded_file.getvalue())
    )
    return workbook.sheet_names


def choose_sheet(uploaded_file, label, key):
    """Display a worksheet selector for Excel files."""
    if uploaded_file is None:
        return None

    sheets = get_excel_sheets(uploaded_file)

    if not sheets:
        return None

    return st.selectbox(
        label,
        sheets,
        key=key
    )


def guess_column(columns, keywords):
    """Find a likely column based on its heading."""
    for column in columns:
        normalized = str(column).strip().lower()
        if any(word in normalized for word in keywords):
            return column
    return columns[0] if len(columns) else None


def parse_amount(series):
    """
    Convert accounting amounts to numeric values.
    Supports commas, currency symbols, and brackets for negatives.
    """
    def convert(value):
        if pd.isna(value):
            return None

        if isinstance(value, (int, float)):
            return float(value)

        text = str(value).strip()

        if not text or text.lower() in {"nan", "none", "n/a", "-"}:
            return None

        negative = text.startswith("(") and text.endswith(")")
        text = text.replace(",", "")
        text = text.replace("(", "").replace(")", "")
        text = re.sub(r"[^0-9.\-]", "", text)

        try:
            number = float(text)
            return -abs(number) if negative else number
        except ValueError:
            return None

    return series.map(convert)


def get_document_text(uploaded_file):
    """Extract text from common supporting document formats."""
    filename = uploaded_file.name.lower()
    file_bytes = uploaded_file.getvalue()

    if filename.endswith(".txt"):
        return file_bytes.decode("utf-8", errors="replace")

    if filename.endswith(".csv"):
        return pd.read_csv(
            io.BytesIO(file_bytes)
        ).head(300).to_csv(index=False)

    if filename.endswith((".xlsx", ".xlsm", ".xls")):
        workbook = pd.ExcelFile(io.BytesIO(file_bytes))
        parts = []

        for sheet in workbook.sheet_names[:10]:
            frame = pd.read_excel(
                io.BytesIO(file_bytes),
                sheet_name=sheet
            )
            parts.append(
                f"Sheet: {sheet}\n"
                + frame.head(100).to_csv(index=False)
            )

        return "\n".join(parts)

    if filename.endswith(".pdf"):
        import fitz

        document = fitz.open(stream=file_bytes, filetype="pdf")
        return "\n".join(
            page.get_text() for page in document[:30]
        )

    if filename.endswith(".docx"):
        from docx import Document

        document = Document(io.BytesIO(file_bytes))
        return "\n".join(
            paragraph.text for paragraph in document.paragraphs
        )

    if filename.endswith((".png", ".jpg", ".jpeg", ".tiff", ".bmp")):
        from PIL import Image
        import pytesseract

        image = Image.open(io.BytesIO(file_bytes))
        return pytesseract.image_to_string(image)

    return "Text extraction is not supported for this file type."


# ==========================================================
# 2. UPLOAD GENERAL LEDGER AND REFERENCE SAMPLING SHEET
# ==========================================================

st.header("1. Upload Accounting Records")

ledger_file = st.file_uploader(
    "Upload complete General Ledger",
    type=["xlsx", "xls", "xlsm", "csv"],
    key="expense_ledger_upload"
)

reference_file = st.file_uploader(
    "Upload previous Sampling Sheet / Reference Template",
    type=["xlsx", "xls", "xlsm", "csv"],
    key="sampling_reference_upload"
)

working_file = st.file_uploader(
    "Optional: Upload utility bills working or reconciliation",
    type=["xlsx", "xls", "xlsm", "csv"],
    key="utility_working_upload"
)

bill_files = st.file_uploader(
    "Optional: Upload supporting bills and audit evidence",
    type=["pdf", "docx", "txt", "png", "jpg", "jpeg", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
    key="supporting_bills_upload"
)

ledger_df = None
working_df = None
reference_df = None

ledger_sheet = choose_sheet(
    ledger_file,
    "Select General Ledger worksheet",
    "ledger_sheet_selector"
) if ledger_file else None

reference_sheet = choose_sheet(
    reference_file,
    "Select reference sampling worksheet",
    "reference_sheet_selector"
) if reference_file else None

working_sheet = choose_sheet(
    working_file,
    "Select utility working worksheet",
    "working_sheet_selector"
) if working_file else None

if ledger_file:
    try:
        ledger_df = read_tabular_file(ledger_file, ledger_sheet)

        if ledger_df is not None:
            ledger_df.columns = [
                str(column).strip()
                for column in ledger_df.columns
            ]
            st.success(
                f"General Ledger loaded: {len(ledger_df):,} rows."
            )
            st.dataframe(ledger_df.head(10), use_container_width=True)

    except Exception as exc:
        st.error(f"Could not read the General Ledger: {exc}")

if reference_file:
    try:
        reference_df = read_tabular_file(
            reference_file,
            reference_sheet
        )

        if reference_df is not None:
            reference_df.columns = [
                str(column).strip()
                for column in reference_df.columns
            ]
            st.success("Reference sampling sheet loaded.")
            st.dataframe(
                reference_df.head(10),
                use_container_width=True
            )

    except Exception as exc:
        st.error(f"Could not read the reference sheet: {exc}")

if working_file:
    try:
        working_df = read_tabular_file(
            working_file,
            working_sheet
        )

        if working_df is not None:
            working_df.columns = [
                str(column).strip()
                for column in working_df.columns
            ]
            st.success("Utility working file loaded.")

    except Exception as exc:
        st.error(f"Could not read the utility working file: {exc}")


# ==========================================================
# 3. EXPENSE SAMPLING
# ==========================================================

st.divider()
st.header("2. Expense Sampling and 100% Verification")

st.write(
    "Select the relevant ledger columns and specify which expense "
    "heads require 100% verification. Remaining transactions can "
    "be selected using the sampling percentage and value-based "
    "selection."
)

sampling_result = None

if ledger_df is None or ledger_df.empty:
    st.info(
        "Upload your General Ledger to configure expense sampling."
    )

else:
    columns = list(ledger_df.columns)

    col1, col2 = st.columns(2)

    with col1:
        account_col = st.selectbox(
            "Account / Expense Head column",
            columns,
            index=columns.index(
                guess_column(
                    columns,
                    ["account name", "expense head", "account head",
                     "ledger name", "description", "particular"]
                )
            ),
            key="sampling_account_column"
        )

        amount_col = st.selectbox(
            "Transaction Amount column",
            columns,
            index=columns.index(
                guess_column(
                    columns,
                    ["amount", "debit", "transaction value",
                     "local amount", "net amount"]
                )
            ),
            key="sampling_amount_column"
        )

    with col2:
        voucher_default = guess_column(
            columns,
            ["voucher number", "voucher no", "document number",
             "journal number", "reference"]
        )

        date_default = guess_column(
            columns,
            ["posting date", "voucher date", "transaction date", "date"]
        )

        voucher_col = st.selectbox(
            "Voucher / Reference column (optional)",
            ["Not available"] + columns,
            index=(
                columns.index(voucher_default) + 1
                if voucher_default in columns else 0
            ),
            key="sampling_voucher_column"
        )

        date_col = st.selectbox(
            "Transaction Date column (optional)",
            ["Not available"] + columns,
            index=(
                columns.index(date_default) + 1
                if date_default in columns else 0
            ),
            key="sampling_date_column"
        )

    default_heads = (
        "Utilities\n"
        "Fee and Subscription\n"
        "Legal and Professional Charges\n"
        "Salaries\n"
        "Depreciation"
    )

    heads_text = st.text_area(
        "Expense heads for 100% verification",
        value=default_heads,
        help=(
            "Enter one expense head per line. Add any other heads "
            "shown in your reference sampling sheet."
        ),
        key="full_verification_heads"
    )

    sample_percentage = st.slider(
        "Sampling percentage for the remaining population",
        min_value=1,
        max_value=100,
        value=20,
        step=1,
        help=(
            "Applied to transactions not already selected for "
            "100% verification. This is a practical selection "
            "percentage, not a statistically calculated sample size."
        ),
        key="expense_sampling_percentage"
    )

    include_high_value = st.checkbox(
        "Include the highest-value transactions from the remaining population",
        value=True,
        key="include_high_value_items"
    )

    random_seed = st.number_input(
        "Random selection seed (use the same number to reproduce the sample)",
        min_value=0,
        max_value=999999,
        value=42,
        step=1,
        key="sampling_random_seed"
    )

    if st.button(
        "Generate Expense Sampling Workbook",
        type="primary",
        key="generate_sampling_workbook"
    ):
        try:
            data = ledger_df.copy()
            data["_Original_Row_Number"] = range(2, len(data) + 2)

            data["_Parsed_Amount"] = parse_amount(data[amount_col])

            account_text = (
                data[account_col]
                .fillna("")
                .astype(str)
                .str.strip()
            )

            data["_Account_Head"] = account_text

            requested_heads = [
                head.strip()
                for head in re.split(r"[\n,;]+", heads_text)
                if head.strip()
            ]

            if not requested_heads:
                st.error(
                    "Enter at least one expense head for 100% verification."
                )
                st.stop()

            # Match the requested categories against the account/head column.
            # Matching is case-insensitive and uses partial text matching.
            full_verification_mask = pd.Series(
                False, index=data.index
            )

            for head in requested_heads:
                full_verification_mask |= account_text.str.contains(
                    re.escape(head),
                    case=False,
                    na=False,
                    regex=True
                )

            full_verification = data.loc[
                full_verification_mask
            ].copy()

            remaining_population = data.loc[
                ~full_verification_mask
            ].copy()

            invalid_amount_rows = data.loc[
                data["_Parsed_Amount"].isna()
            ].copy()

            # Select a percentage of the remaining population.
            population_count = len(remaining_population)
            target_count = (
                int(round(population_count * sample_percentage / 100))
                if population_count else 0
            )

            selected_indices = []

            if population_count and target_count:
                if include_high_value:
                    # Select the highest absolute-value items first.
                    valid_amounts = remaining_population[
                        remaining_population["_Parsed_Amount"].notna()
                    ].copy()

                    high_value_count = min(
                        target_count,
                        max(1, int(round(target_count * 0.25)))
                    )

                    if not valid_amounts.empty:
                        high_value_rows = valid_amounts.reindex(
                            valid_amounts["_Parsed_Amount"]
                            .abs()
                            .sort_values(ascending=False)
                            .index
                        ).head(high_value_count)

                        selected_indices.extend(
                            high_value_rows.index.tolist()
                        )

                remaining_candidates = remaining_population.drop(
                    index=selected_indices,
                    errors="ignore"
                )

                needed = max(
                    0,
                    target_count - len(selected_indices)
                )

                if needed and not remaining_candidates.empty:
                    sample_size = min(
                        needed,
                        len(remaining_candidates)
                    )

                    random_sample = remaining_candidates.sample(
                        n=sample_size,
                        random_state=int(random_seed)
                    )

                    selected_indices.extend(
                        random_sample.index.tolist()
                    )

            expense_samples = remaining_population.loc[
                remaining_population.index.isin(selected_indices)
            ].copy()

            # Flag duplicate voucher references where a usable voucher
            # column was selected. These are candidates for review only.
            if voucher_col != "Not available":
                voucher_values = (
                    data[voucher_col].fillna("").astype(str).str.strip()
                )

                nonblank_voucher = voucher_values.ne("")

                duplicate_mask = (
                    nonblank_voucher
                    & data.duplicated(
                        subset=[voucher_col, amount_col],
                        keep=False
                    )
                )

                potential_duplicates = data.loc[
                    duplicate_mask
                ].copy()
            else:
                potential_duplicates = pd.DataFrame(
                    columns=data.columns
                )

            # Add a clear status column to selected rows.
            full_verification.insert(
                0,
                "Verification_Type",
                "100% Verification"
            )

            expense_samples.insert(
                0,
                "Verification_Type",
                "Expense Sample"
            )

            # Summary by expense head.
            summary_rows = []

            for head in requested_heads:
                matching = full_verification[
                    full_verification["_Account_Head"]
                    .str.contains(
                        re.escape(head),
                        case=False,
                        na=False
                    )
                ]

                if not matching.empty:
                    summary_rows.append({
                        "Verification Category": head,
                        "Transactions Identified": len(matching),
                        "Amount Total": matching["_Parsed_Amount"].sum(),
                        "Status": "Selected for 100% verification"
                    })
                else:
                    summary_rows.append({
                        "Verification Category": head,
                        "Transactions Identified": 0,
                        "Amount Total": 0,
                        "Status": "No matching account head found"
                    })

            summary_rows.append({
                "Verification Category": "Remaining Expense Population",
                "Transactions Identified": len(remaining_population),
                "Amount Total": remaining_population[
                    "_Parsed_Amount"
                ].sum(),
                "Status": "Population for percentage-based sampling"
            })

            summary_rows.append({
                "Verification Category": "Selected Expense Sample",
                "Transactions Identified": len(expense_samples),
                "Amount Total": expense_samples[
                    "_Parsed_Amount"
                ].sum(),
                "Status": f"{sample_percentage}% target; includes value-based selection"
            })

            summary_rows.append({
                "Verification Category": "Invalid / Missing Amount",
                "Transactions Identified": len(invalid_amount_rows),
                "Amount Total": 0,
                "Status": "Investigate amount format or missing data"
            })

            summary_df = pd.DataFrame(summary_rows)

            # Remove helper columns from exported data.
            helper_columns = [
                "_Parsed_Amount",
                "_Account_Head"
            ]

            def clean_export(frame):
                return frame.drop(
                    columns=helper_columns,
                    errors="ignore"
                )

            full_verification_export = clean_export(
                full_verification
            )
            samples_export = clean_export(expense_samples)
            population_export = clean_export(remaining_population)
            invalid_export = clean_export(invalid_amount_rows)
            duplicates_export = clean_export(potential_duplicates)

            st.session_state["expense_sampling_workbook_data"] = {
                "summary": summary_df,
                "samples": samples_export,
                "full_verification": full_verification_export,
                "population": population_export,
                "invalid_amounts": invalid_export,
                "potential_duplicates": duplicates_export,
                "reference": reference_df
            }

            st.success("Expense sampling has been generated.")

        except Exception as exc:
            st.error(f"Could not generate the sampling workbook: {exc}")


# ==========================================================
# 4. DISPLAY AND DOWNLOAD SAMPLING WORKBOOK
# ==========================================================

sampling_result = st.session_state.get(
    "expense_sampling_workbook_data"
)

if sampling_result:
    st.subheader("Expense Sampling Results")

    st.dataframe(
        sampling_result["summary"],
        use_container_width=True
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "Expense Samples",
        "100% Verification",
        "Potential Duplicates",
        "Invalid Amounts"
    ])

    with tab1:
        st.dataframe(
            sampling_result["samples"],
            use_container_width=True
        )

    with tab2:
        st.dataframe(
            sampling_result["full_verification"],
            use_container_width=True
        )

    with tab3:
        st.dataframe(
            sampling_result["potential_duplicates"],
            use_container_width=True
        )

    with tab4:
        st.dataframe(
            sampling_result["invalid_amounts"],
            use_container_width=True
        )

    excel_buffer = io.BytesIO()

    with pd.ExcelWriter(
        excel_buffer,
        engine="openpyxl"
    ) as writer:
        sampling_result["summary"].to_excel(
            writer, sheet_name="Sampling Summary", index=False
        )
        sampling_result["samples"].to_excel(
            writer, sheet_name="Expense Samples", index=False
        )
        sampling_result["full_verification"].to_excel(
            writer, sheet_name="100% Verification", index=False
        )
        sampling_result["potential_duplicates"].to_excel(
            writer, sheet_name="Potential Duplicates", index=False
        )
        sampling_result["population"].to_excel(
            writer, sheet_name="Eligible Population", index=False
        )
        sampling_result["invalid_amounts"].to_excel(
            writer, sheet_name="Invalid Amount Rows", index=False
        )

        if sampling_result["reference"] is not None:
            sampling_result["reference"].to_excel(
                writer, sheet_name="Reference Template", index=False
            )

    st.download_button(
        "Download Expense Sampling Workbook",
        data=excel_buffer.getvalue(),
        file_name="expense_sampling_workbook.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary"
    )

    st.caption(
        "The reference worksheet is included for comparison. "
        "The output does not automatically reproduce its exact formatting "
        "or formulas. Review the account-head matches before relying on "
        "the sampling results."
    )


# ==========================================================
# 5. AI AUDIT ASSISTANT
# ==========================================================

st.divider()
st.header("3. AI Audit Assistant")

st.write(
    "Give an instruction in your own words. The AI will examine "
    "the uploaded records and prepare an audit analysis or draft "
    "working paper based on the evidence available."
)

audit_task = st.text_area(
    "What do you want the AI to do?",
    value=(
        "Prepare a utility bills audit working paper. Compare the "
        "general ledger, utility bills working and uploaded bills. "
        "Identify possible unrecorded bills, duplicate bills, "
        "amount differences, period cut-off issues and unusual items. "
        "Show the evidence for each exception, identify missing "
        "information, and propose draft adjusting entries where "
        "appropriate. Do not assume that a missing match proves "
        "an error."
    ),
    height=180,
    key="ai_audit_task"
)

audit_area = st.selectbox(
    "Audit area",
    [
        "General audit analysis",
        "Utility bills",
        "Trial balance",
        "Financial statements",
        "Revenue",
        "Expenses",
        "Receivables and payables",
        "Bank and cash",
        "Fixed assets",
        "Payroll",
        "Tax and statutory compliance",
        "Inventory",
        "Journal entries",
        "Custom procedure"
    ],
    key="ai_audit_area"
)

report_style = st.selectbox(
    "Output format",
    [
        "Detailed audit working paper",
        "Audit exception report",
        "Analytical review",
        "Reconciliation",
        "Draft adjusting journal entries",
        "Audit programme and checklist"
    ],
    key="ai_report_style"
)

period = st.text_input(
    "Financial period",
    placeholder="e.g. Year ended 30 June 2026",
    key="ai_audit_period"
)


# ==========================================================
# 6. BUILD AI CONTEXT
# ==========================================================

def add_table_context(context_list, label, dataframe, max_rows=300):
    if dataframe is not None and not dataframe.empty:
        preview = dataframe.head(max_rows).to_csv(index=False)

        context_list.append(
            f"\n--- {label}: first "
            f"{min(len(dataframe), max_rows)} rows of "
            f"{len(dataframe)} total rows ---\n"
            f"{preview[:25000]}"
        )

        if len(dataframe) > max_rows:
            context_list.append(
                f"\nNOTE: {label} contains {len(dataframe)} rows. "
                f"Only the first {max_rows} rows were included. "
                "Do not draw conclusions about rows not supplied."
            )


def build_audit_context():
    parts = []

    # Use the actual uploaded ledger and working data.
    add_table_context(
        parts, "GENERAL LEDGER", ledger_df
    )
    add_table_context(
        parts, "UTILITY BILLS WORKING", working_df
    )
    add_table_context(
        parts, "REFERENCE SAMPLING SHEET", reference_df
    )

    # Include generated sampling results if available.
    current_sampling = st.session_state.get(
        "expense_sampling_workbook_data"
    )

    if current_sampling:
        add_table_context(
            parts,
            "SELECTED EXPENSE SAMPLES",
            current_sampling["samples"]
        )
        add_table_context(
            parts,
            "100% VERIFICATION POPULATION",
            current_sampling["full_verification"]
        )
        add_table_context(
            parts,
            "POTENTIAL DUPLICATES",
            current_sampling["potential_duplicates"]
        )

    if bill_files:
        for uploaded_file in bill_files:
            try:
                document_text = get_document_text(uploaded_file)

                parts.append(
                    f"\n--- SUPPORTING FILE: "
                    f"{uploaded_file.name} ---\n"
                    f"{document_text[:12000]}"
                )

            except Exception as exc:
                parts.append(
                    f"\n--- SUPPORTING FILE: "
                    f"{uploaded_file.name} ---\n"
                    f"Could not read file: {exc}"
                )

    if "edited_bills" in st.session_state:
        reviewed_bills = st.session_state["edited_bills"]

        if isinstance(reviewed_bills, pd.DataFrame):
            add_table_context(
                parts, "REVIEWED BILL REGISTER", reviewed_bills
            )

    return "\n".join(parts)


# ==========================================================
# 7. GENERATE AUDIT WORKING PAPER
# ==========================================================

if st.button(
    "Generate Audit Working Paper",
    type="primary",
    key="generate_ai_working_paper"
):

    context = build_audit_context()

    if not audit_task.strip():
        st.error("Please enter an instruction.")

    elif not context.strip():
        st.error(
            "Please upload a General Ledger, working paper, "
            "reference sheet or supporting document first."
        )

    else:
        try:
            api_key = st.secrets.get("OPENAI_API_KEY", "")
        except Exception:
            api_key = ""

        if not api_key:
            st.error(
                "OPENAI_API_KEY is missing. Add it in your "
                "Streamlit app's Secrets settings."
            )

        else:
            instructions = """
You are an AI audit assistant supporting a professional auditor.

Analyse only the supplied accounting records and evidence.

MANDATORY RULES:
1. Never invent transactions, balances, references, dates or evidence.
2. Distinguish observed facts, calculations, assumptions and hypotheses.
3. Identify the files, columns, rows and figures supporting each finding.
4. Identify missing information and further evidence required.
5. Do not claim a procedure was completed without sufficient evidence.
6. Quantify reconciliation differences and explain the comparison basis.
7. Label potential duplicates as candidates requiring investigation.
8. Do not conclude a cut-off error from invoice date alone.
9. Present proposed journal entries as drafts only.
10. Ensure draft debits equal credits.
11. Do not automatically post, approve or alter accounting records.
12. Do not declare fraud from anomalies alone.
13. State scope limitations and incomplete records.
14. Use clear, professional audit language.

Suggested structure:
- Working paper title and period
- Objective
- Records received
- Procedures performed
- Results and calculations
- Exceptions with supporting evidence
- Missing evidence and follow-up requests
- Draft adjustments, if justified
- Conclusion limited to the available evidence
- Reviewer points and limitations

If data is incomplete, explain the limitation. Never fabricate
a complete conclusion.
"""

            user_input = f"""
AUDIT AREA: {audit_area}
REQUESTED OUTPUT: {report_style}
FINANCIAL PERIOD: {period or "Not specified"}

USER INSTRUCTION:
{audit_task}

SUPPLIED ACCOUNTING DATA AND DOCUMENT EXTRACTS:
{context}

Analyse the supplied evidence and produce the requested work product.
Show calculations and clearly label any limitations.
"""

            try:
                with st.spinner(
                    "Analysing the uploaded records..."
                ):
                    client = OpenAI(api_key=api_key)

                    response = client.responses.create(
                        model="gpt-5.5",
                        instructions=instructions,
                        input=user_input,
                        max_output_tokens=8000
                    )

                    result = response.output_text

                st.session_state["ai_audit_result"] = result

            except Exception as exc:
                st.error(
                    "The AI request failed. Check your API key, "
                    "model access, usage limits and network connection."
                )
                st.caption(str(exc))


# ==========================================================
# 8. DISPLAY AND DOWNLOAD AI WORKING PAPER
# ==========================================================

if st.session_state.get("ai_audit_result"):
    st.subheader("AI-Generated Draft")

    st.markdown(
        st.session_state["ai_audit_result"]
    )

    st.download_button(
        "Download Working Paper (Markdown)",
        data=st.session_state["ai_audit_result"],
        file_name="ai_audit_working_paper.md",
        mime="text/markdown"
    )

    st.warning(
        "This is an AI-generated draft, not an audit opinion. "
        "Verify the figures against the complete source records "
        "and have the work reviewed before relying on it."
    )
