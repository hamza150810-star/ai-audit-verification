
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
