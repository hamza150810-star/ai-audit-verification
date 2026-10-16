
# --------------------------------------------------
# AI AUDIT ASSISTANT: USER INSTRUCTIONS + WORKPAPERS
# --------------------------------------------------

import os
from openai import OpenAI

import streamlit as st
import pandas as pd
import pytesseract
import fitz
import io
import re

from pathlib import Path
from datetime import datetime
from PIL import Image, ImageOps, ImageEnhance
from docx import Document

st.set_page_config(
    page_title="AI Audit Verification",
    page_icon="📊",
    layout="wide"
)

st.title("AI Audit Verification System")
st.write("Upload accounting records and supporting documents.")

st.divider()

st.divider()
st.header("6. AI Audit Assistant")
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

# Prepare only the data available in this app session.
# Limit row counts and text length to control processing costs.

context_parts = []


def add_table_context(label, dataframe, max_rows=300):
    if dataframe is not None and not dataframe.empty:
        preview = dataframe.head(max_rows).to_csv(index=False)
        context_parts.append(
            f"\n--- {label}: first {min(len(dataframe), max_rows)} "
            f"rows of {len(dataframe)} total rows ---\n"
            f"{preview[:25000]}"
        )
        if len(dataframe) > max_rows:
            context_parts.append(
                f"\nNOTE: {label} has {len(dataframe)} rows. "
                f"Only the first {max_rows} rows are included. "
                "Do not draw conclusions about rows not supplied."
            )


add_table_context("GENERAL LEDGER", ledger_df)
add_table_context("GENERAL LEDGER", ledger_df)
add_table_context("UTILITY BILLS WORKING", working_df)

# Include uploaded bill/document text where readable.
if bill_files:
    for uploaded_file in bill_files:
        try:
            document_text = get_document_text(uploaded_file)
            context_parts.append(
                f"\n--- SUPPORTING FILE: {uploaded_file.name} ---\n"
                f"{document_text[:12000]}"
            )
        except Exception as exc:
            context_parts.append(
                f"\n--- SUPPORTING FILE: {uploaded_file.name} ---\n"
                f"Could not read file: {str(exc)}"
            )

# Include the reviewed bill register, if available.
if "edited_bills" in st.session_state:
    reviewed_bills = st.session_state["edited_bills"]
    context_parts.append(
        "\n--- REVIEWED BILL REGISTER ---\n"
        + reviewed_bills.head(300).to_csv(index=False)[:20000]
    )

context = "\n".join(context_parts)

st.caption(
    "The AI can only analyse the information supplied to it. "
    "Large workbooks may be truncated in this first version."
)

if st.button("Generate Audit Working Paper", type="primary"):

    if not audit_task.strip():
        st.error("Please enter an instruction.")
    elif not context.strip():
        st.error(
            "Please upload at least one ledger, working paper "
            "or supporting document first."
        )
    elif not st.secrets.get("OPENAI_API_KEY", ""):
        st.error(
            "OPENAI_API_KEY is missing. Add it under your "
            "Streamlit app's Secrets settings."
        )
    else:
        instructions = """
You are an AI audit assistant supporting a professional auditor.

Your job is to analyse the supplied accounting records and prepare
useful audit working papers according to the user's instruction.

MANDATORY RULES:
1. Never invent transactions, balances, evidence, references or dates.
2. Distinguish observed facts, calculations, assumptions and hypotheses.
3. State which files, columns, rows and figures support each finding.
4. If required information is missing, identify it and ask for it.
5. Do not claim an audit procedure was completed if the evidence
   required to perform it was not provided.
6. For reconciliations, show the comparison basis and quantify
   differences. Do not claim that a match proves validity.
7. For duplicate checks, label results as potential duplicates until
   supporting details establish whether they are genuine duplicates.
8. For cut-off, consider service period, invoice date, posting date,
   payment date, accounting basis and reporting date. Do not conclude
   a cut-off error from invoice date alone.
9. For journal entries, provide a draft only. Show account, debit,
   credit, amount, rationale and evidence needed. Ensure debits equal
   credits. Explain when the correct account or period is uncertain.
10. Do not automatically post, approve or alter accounting records.
11. Do not declare fraud from anomalies alone.
12. Where applicable, ask which accounting framework and jurisdiction
    apply instead of assuming.
13. Identify scope limitations, incomplete records and exceptions.
14. Use clear professional audit language.

Suggested working paper structure:
- Working paper title and reporting period
- Objective
- Records received
- Procedures performed
- Results and reconciliations
- Exceptions with supporting evidence
- Missing evidence and follow-up requests
- Draft adjustments, if justified
- Conclusion limited to the procedures and evidence available
- Reviewer points and limitations

If data is incomplete, prepare the useful sections possible and clearly
list what is still required. Never fabricate a complete conclusion.
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
Show calculations and clearly label any limitation.
"""

        try:
            with st.spinner(
                "Analysing the uploaded records and preparing the draft..."
            ):
                client = OpenAI(
                    api_key=st.secrets["OPENAI_API_KEY"]
                )

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
                "The AI request failed. Check the API key, API access, "
                "available model, usage limits and network connection."
            )
            st.caption(str(exc))


if st.session_state.get("ai_audit_result"):
    st.subheader("AI-Generated Draft")
    st.markdown(st.session_state["ai_audit_result"])

    st.download_button(
        "Download Working Paper (Markdown)",
        data=st.session_state["ai_audit_result"],
        file_name="ai_audit_working_paper.md",
        mime="text/markdown"
    )

    st.warning(
        "This is an AI-generated draft, not an audit opinion. "
        "Verify figures against the complete source records and "
        "have the work reviewed before relying on it."
    )
    
# --------------------------------------------------
# 7. EXPENSE SAMPLING AND 100% VERIFICATION
# --------------------------------------------------

import random
from io import BytesIO

st.divider()
st.header("7. Expense Sampling and 100% Verification")

st.write(
    "Upload the complete expense ledger and your previous sampling "
    "sheet as a reference. The app will prepare an expense sampling "
    "working paper in Excel."
)

ledger_upload = st.file_uploader(
    "Upload complete expense ledger",
    type=["xlsx", "xls"],
    key="expense_ledger_upload"
)

reference_upload = st.file_uploader(
    "Upload previous sampling sheet (reference)",
    type=["xlsx", "xls"],
    key="sampling_reference_upload"
)

if ledger_upload is not None:

    try:
        excel_file = pd.ExcelFile(ledger_upload)

        ledger_sheet = st.selectbox(
            "Select ledger worksheet",
            excel_file.sheet_names,
            key="expense_ledger_sheet"
        )

        ledger_data = pd.read_excel(
            ledger_upload,
            sheet_name=ledger_sheet
        )

        ledger_data = ledger_data.dropna(how="all").copy()

        st.write("Ledger preview")
        st.dataframe(ledger_data.head(10), use_container_width=True)

        if reference_upload is not None:
            reference_excel = pd.ExcelFile(reference_upload)

            reference_sheet = st.selectbox(
                "Select reference sampling worksheet",
                reference_excel.sheet_names,
                key="expense_reference_sheet"
            )

            reference_data = pd.read_excel(
                reference_upload,
                sheet_name=reference_sheet,
                header=None
            )

            st.write("Reference sampling sheet preview")
            st.dataframe(
                reference_data.head(15),
                use_container_width=True
            )

        st.subheader("Map your ledger columns")

        ledger_columns = list(ledger_data.columns)
        column_options = ["Not available"] + ledger_columns

        account_col = st.selectbox(
            "Expense account / head",
            column_options,
            key="expense_account_col"
        )

        amount_col = st.selectbox(
            "Transaction amount",
            column_options,
            key="expense_amount_col"
        )

        date_col = st.selectbox(
            "Transaction date",
            column_options,
            key="expense_date_col"
        )

        voucher_col = st.selectbox(
            "Voucher / reference number",
            column_options,
            key="expense_voucher_col"
        )

        vendor_col = st.selectbox(
            "Vendor / payee",
            column_options,
            key="expense_vendor_col"
        )

        invoice_col = st.selectbox(
            "Invoice / bill reference",
            column_options,
            key="expense_invoice_col"
        )

        description_col = st.selectbox(
            "Narration / description",
            column_options,
            key="expense_description_col"
        )

        st.subheader("Sampling settings")

        sample_percentage = st.slider(
            "Sampling percentage of remaining expenses",
            min_value=1,
            max_value=100,
            value=10,
            key="expense_sample_percentage"
        )

        random_seed = st.number_input(
            "Random seed for reproducible selection",
            min_value=1,
            max_value=999999,
            value=2026,
            key="expense_random_seed"
        )

        st.subheader("Expense heads requiring 100% verification")

        default_heads = (
            "Utilities\n"
            "Fee and Subscription\n"
            "Legal and Professional Charges\n"
            "Salaries\n"
            "Depreciation"
        )

        heads_input = st.text_area(
            "Enter one expense head per line. "
            "Compare these with your reference sampling sheet.",
            value=default_heads,
            height=150,
            key="expense_full_verification_heads"
        )

        st.caption(
            "Review the categories before generating the paper. "
            "Matching is based on the account name and may require "
            "manual correction for differently named accounts."
        )

        if st.button(
            "Generate Expense Sampling Working Paper",
            type="primary",
            key="generate_expense_sampling"
        ):

            if account_col == "Not available":
                st.error("Select the expense account column.")
                st.stop()

            if amount_col == "Not available":
                st.error("Select the transaction amount column.")
                st.stop()

            work = ledger_data.copy()

            # Preserve the original worksheet row for tracing.
            work.insert(
                0,
                "Source Excel Row",
                range(2, len(work) + 2)
            )

            # Convert amounts into numeric values.
            work["Sampling Amount"] = pd.to_numeric(
                work[amount_col]
                .astype(str)
                .str.replace(",", "", regex=False)
                .str.replace("Rs.", "", regex=False)
                .str.replace("PKR", "", regex=False)
                .str.strip(),
                errors="coerce"
            )

            invalid_amounts = work[
                work["Sampling Amount"].isna()
            ].copy()

            population = work[
                work["Sampling Amount"].notna()
                & (work["Sampling Amount"] != 0)
            ].copy()

            heads = [
                line.strip().lower()
                for line in heads_input.splitlines()
                if line.strip()
            ]

            account_text = (
                population[account_col]
                .fillna("")
                .astype(str)
                .str.lower()
            )

            full_mask = account_text.apply(
                lambda account: any(
                    head in account for head in heads
                )
            )

            full_verification = population[full_mask].copy()
            eligible = population[~full_mask].copy()

            full_verification["Selection Category"] = (
                "100% verification"
            )

            full_verification["Verification Status"] = "Pending"
            full_verification["Auditor Comments"] = ""

            # Identify simple risk indicators.
            if len(eligible):

                abs_amount = eligible["Sampling Amount"].abs()

                high_value_threshold = abs_amount.quantile(0.90)

                risk_reasons = []

                for _, row in eligible.iterrows():

                    reasons = []

                    amount = abs(row["Sampling Amount"])

                    if amount >= high_value_threshold:
                        reasons.append("High-value item")

                    if (
                        voucher_col != "Not available"
                        and not str(row.get(voucher_col, "")).strip()
                        or (
                            voucher_col != "Not available"
                            and pd.isna(row.get(voucher_col))
                        )
                    ):
                        reasons.append("Missing voucher reference")

                    if (
                        invoice_col != "Not available"
                        and (
                            pd.isna(row.get(invoice_col))
                            or not str(row.get(invoice_col, "")).strip()
                        )
                    ):
                        reasons.append("Missing invoice reference")

                    if (
                        description_col != "Not available"
                        and pd.isna(row.get(description_col))
                    ):
                        reasons.append("Missing description")

                    if amount > 0 and amount % 1000 == 0:
                        reasons.append("Round-sum amount")

                    risk_reasons.append(
                        "; ".join(reasons)
                        if reasons
                        else "No automated risk flag"
                    )

                eligible["Risk Indicators"] = risk_reasons

                eligible["Risk Score"] = eligible[
                    "Risk Indicators"
                ].apply(
                    lambda x: 0 if x == "No automated risk flag"
                    else len(x.split(";"))
                )

                sample_count = min(
                    len(eligible),
                    max(
                        1,
                        round(
                            len(eligible)
                            * sample_percentage / 100
                        )
                    )
                )

                # Aim for half risk-based and half random selection.
                risk_count = (sample_count + 1) // 2

                risk_pool = eligible.sort_values(
                    ["Risk Score", "Sampling Amount"],
                    ascending=[False, False],
                    key=lambda col: col.abs()
                    if col.name == "Sampling Amount"
                    else col
                )

                risk_samples = risk_pool.head(risk_count)

                remaining = eligible.drop(
                    index=risk_samples.index
                )

                random_count = sample_count - len(risk_samples)

                random_samples = (
                    remaining.sample(
                        n=min(random_count, len(remaining)),
                        random_state=int(random_seed)
                    )
                    if random_count > 0
                    else remaining.head(0)
                )

                selected = pd.concat(
                    [risk_samples, random_samples]
                ).drop_duplicates()

                selected["Selection Method"] = [
                    "Risk-based"
                    if index in risk_samples.index
                    else "Random"
                    for index in selected.index
                ]

                selected["Verification Status"] = "Pending"
                selected["Auditor Comments"] = ""

                exceptions = eligible[
                    eligible["Risk Indicators"]
                    != "No automated risk flag"
                ].copy()

                exceptions["Review Status"] = (
                    "Requires auditor investigation"
                )

            else:
                selected = eligible.copy()
                exceptions = eligible.copy()

            # Population summary.
            summary = pd.DataFrame({
                "Metric": [
                    "Source worksheet",
                    "Source transaction rows",
                    "Rows with invalid/missing amount",
                    "Expense population included",
                    "100% verification population",
                    "Remaining eligible population",
                    "Selected sample count",
                    "Requested sample percentage",
                    "Actual sample percentage",
                    "Random seed",
                    "Generated on"
                ],
                "Value": [
                    ledger_sheet,
                    len(ledger_data),
                    len(invalid_amounts),
                    len(population),
                    len(full_verification),
                    len(eligible),
                    len(selected),
                    f"{sample_percentage}%",
                    (
                        f"{len(selected) / len(eligible) * 100:.2f}%"
                        if len(eligible)
                        else "0%"
                    ),
                    int(random_seed),
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ]
            })

            # Excel output.
            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                summary.to_excel(
                    writer,
                    sheet_name="Sampling Summary",
                    index=False
                )

                selected.to_excel(
                    writer,
                    sheet_name="Expense Samples",
                    index=False
                )

                full_verification.to_excel(
                    writer,
                    sheet_name="100% Verification",
                    index=False
                )

                exceptions.to_excel(
                    writer,
                    sheet_name="Potential Exceptions",
                    index=False
                )

                eligible.to_excel(
                    writer,
                    sheet_name="Eligible Population",
                    index=False
                )

                invalid_amounts.to_excel(
                    writer,
                    sheet_name="Invalid Amount Rows",
                    index=False
                )

                if reference_upload is not None:
                    reference_data.to_excel(
                        writer,
                        sheet_name="Reference Template",
                        index=False,
                        header=False
                    )

                for worksheet in writer.book.worksheets:
                    worksheet.freeze_panes = "A2"
                    worksheet.auto_filter.ref = worksheet.dimensions

                    for cells in worksheet.columns:
                        max_length = max(
                            len(str(cell.value or ""))
                            for cell in list(cells)[:200]
                        )

                        worksheet.column_dimensions[
                            cells[0].column_letter
                        ].width = min(
                            max(max_length + 2, 12),
                            40
                        )

            st.success(
                "Expense sampling working paper generated."
            )

            st.write("Sampling summary")
            st.dataframe(summary, use_container_width=True)

            st.write("Selected expense samples")
            st.dataframe(selected, use_container_width=True)

            st.write("100% verification population")
            st.dataframe(
                full_verification,
                use_container_width=True
            )

            st.download_button(
                "Download Expense Sampling Working Paper",
                data=output.getvalue(),
                file_name="Expense_Sampling_Working_Paper.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key="download_expense_sampling"
            )

    except Exception as exc:
        st.error(
            "Could not process the workbook. Check that the file is "
            "a valid Excel workbook and the selected columns are correct."
        )
        st.exception(exc)

st.caption(
    "Audit note: risk indicators are not proof of error or fraud. "
    "Confirm population completeness, materiality, audit objectives "
    "and mandatory testing requirements before relying on the sample."
)
