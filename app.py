
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="AI Audit Verification",
    page_icon="📊",
    layout="wide"
)

st.title("AI Audit Verification")
st.write(
    "Reconcile utility bills against the general ledger, "
    "identify potential omissions, and prepare a reviewable "
    "Excel exception report."
)

st.warning(
    "This tool identifies potential audit exceptions. "
    "Review invoices, payment evidence, accounting policy, "
    "and the applicable reporting framework before posting "
    "any correcting journal entry."
)

st.subheader("1. Review settings")

period_end = st.date_input(
    "Reporting period end",
    value=date.today()
)

basis = st.selectbox(
    "Accounting basis for the utility expense review",
    [
        "Accrual basis",
        "Cash basis",
        "Not confirmed"
    ]
)

tolerance = st.number_input(
    "Permitted amount difference (Rs.)",
    min_value=0.0,
    value=1.0,
    step=1.0
)

expense_account = st.text_input(
    "Proposed debit account",
    value="Utility Expense"
)

payable_account = st.text_input(
    "Proposed credit account",
    value="Accrued Utility Expenses / Payable"
)

st.caption(
    "For a cash-basis review, unmatched bills are flagged for "
    "investigation; an accrual journal entry is not generated "
    "automatically."
)

st.divider()
st.subheader("2. Upload your Excel files")

ledger_file = st.file_uploader(
    "Upload the client's general ledger (.xlsx)",
    type=["xlsx"],
    key="ledger"
)

bills_file = st.file_uploader(
    "Upload the utility bills working (.xlsx)",
    type=["xlsx"],
    key="bills"
)


def read_excel_file(file, label):
    """Read the first worksheet from an uploaded workbook."""
    try:
        file.seek(0)
        book = pd.ExcelFile(file, engine="openpyxl")
        df = book.parse(book.sheet_names[0])
        df.columns = df.columns.astype(str).str.strip()

        if df.empty:
            st.error(label + " is empty.")
            return None

        return df

    except Exception as error:
        st.error("Could not read " + label)
        st.exception(error)
        return None


def choose_column(df, label, key, optional=False):
    """Let the user map their own Excel column headings."""
    options = list(df.columns)

    if optional:
        options = ["(Not available)"] + options

    selected = st.selectbox(
        label,
        options,
        key=key
    )

    if optional and selected == "(Not available)":
        return None

    return selected


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip().casefold()


def clean_amount(series):
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("Rs.", "", regex=False)
        .str.replace("PKR", "", regex=False)
        .str.strip(),
        errors="coerce"
    )


if ledger_file is not None and bills_file is not None:

    ledger = read_excel_file(
        ledger_file, "General ledger"
    )

    bills = read_excel_file(
        bills_file, "Utility bills working"
    )

    if ledger is not None and bills is not None:

        st.subheader("3. Map your Excel columns")

        st.info(
            "Choose the column that represents each field. "
            "Your headings do not need to have specific names."
        )

        left, right = st.columns(2)

        with left:
            st.markdown("**Utility bills working**")

            b_vendor = choose_column(
                bills, "Utility provider/vendor",
                "b_vendor"
            )
            b_invoice = choose_column(
                bills, "Bill or invoice number",
                "b_invoice", optional=True
            )
            b_date = choose_column(
                bills, "Bill date or billing-period date",
                "b_date", optional=True
            )
            b_amount = choose_column(
                bills, "Bill amount",
                "b_amount"
            )

        with right:
            st.markdown("**General ledger**")

            l_vendor = choose_column(
                ledger, "Vendor/provider in ledger",
                "l_vendor", optional=True
            )
            l_invoice = choose_column(
                ledger, "Invoice/reference number in ledger",
                "l_invoice", optional=True
            )
            l_date = choose_column(
                ledger, "Posting date",
                "l_date", optional=True
            )
            l_account = choose_column(
                ledger, "Account name",
                "l_account", optional=True
            )
            l_description = choose_column(
                ledger, "Transaction description",
                "l_description", optional=True
            )
            l_amount = choose_column(
                ledger,
                "Ledger amount column (select expense/debit amount)",
                "l_amount"
            )

        st.caption(
            "Upload a ledger amount column that represents the "
            "expense/debit value. If your ledger uses separate "
            "debit and credit columns, select the debit column."
        )

        st.subheader("4. Run reconciliation")

        if st.button(
            "Reconcile Utility Bills",
            type="primary"
        ):

            # Work on copies; the uploaded files are not modified.
            led = ledger.copy().reset_index(drop=True)
            util = bills.copy().reset_index(drop=True)

            led["_amount"] = clean_amount(led[l_amount])
            util["_amount"] = clean_amount(util[b_amount])

            if l_vendor:
                led["_vendor"] = led[l_vendor].map(clean_text)
            else:
                led["_vendor"] = ""

            util["_vendor"] = util[b_vendor].map(clean_text)

            if l_invoice:
                led["_invoice"] = led[l_invoice].map(clean_text)
            else:
                led["_invoice"] = ""

            if b_invoice:
                util["_invoice"] = util[b_invoice].map(clean_text)
            else:
                util["_invoice"] = ""

            if l_date:
                led["_posting_date"] = pd.to_datetime(
                    led[l_date], errors="coerce", dayfirst=True
                )
            else:
                led["_posting_date"] = pd.NaT

            if b_date:
                util["_bill_date"] = pd.to_datetime(
                    util[b_date], errors="coerce", dayfirst=True
                )
            else:
                util["_bill_date"] = pd.NaT

            cutoff = pd.Timestamp(period_end)
            used_ledger_rows = set()
            statuses = []
            matched_rows = []
            matched_amounts = []
            differences = []
            actions = []
            journal_entries = []
            matched_ledger_status = {}

            for i, bill in util.iterrows():

                amount = bill["_amount"]
                vendor = bill["_vendor"]
                invoice = bill["_invoice"]
                bill_date = bill["_bill_date"]

                # Bills after the reporting date are outside this
                # period-end test.
                if pd.notna(bill_date) and bill_date > cutoff:
                    statuses.append("Outside reporting period")
                    matched_rows.append("")
                    matched_amounts.append(None)
                    differences.append(None)
                    actions.append(
                        "Review in the next reporting period"
                    )
                    continue

                if pd.isna(amount):
                    statuses.append("Invalid or missing bill amount")
                    matched_rows.append("")
                    matched_amounts.append(None)
                    differences.append(None)
                    actions.append(
                        "Check the source bill and amount"
                    )
                    continue

                available = [
                    j for j in led.index
                    if j not in used_ledger_rows
                    and pd.notna(led.loc[j, "_amount"])
                ]

                match_index = None
                match_method = ""

                # First preference: provider + invoice reference.
                if invoice and l_invoice and l_vendor:
                    candidates = [
                        j for j in available
                        if led.loc[j, "_invoice"] == invoice
                        and led.loc[j, "_vendor"] == vendor
                    ]

                    if candidates:
                        match_index = candidates[0]
                        match_method = "Invoice and vendor"

                # If the reference is absent or differs, try matching
                # provider and amount. This is a candidate, not proof.
                if match_index is None:
                    candidates = [
                        j for j in available
                        if abs(
                            abs(float(led.loc[j, "_amount"]))
                            - abs(float(amount))
                        ) <= tolerance
                        and (
                            not l_vendor
                            or led.loc[j, "_vendor"] == vendor
                        )
                    ]

                    if candidates:
                        match_index = candidates[0]
                        match_method = (
                            "Vendor and amount"
                            if l_vendor
                            else "Amount only - review"
                        )

                # If a reference matched but amounts differ, report it.
                if match_index is not None:
                    used_ledger_rows.add(match_index)
                    matched_ledger_status[match_index] = (
                        "Matched to utility bill"
                    )

                    ledger_amount = float(
                        led.loc[match_index, "_amount"]
                    )
                    difference = abs(ledger_amount) - abs(float(amount))

                    if abs(difference) > tolerance:
                        status = "Matched invoice - amount differs"
                        action = (
                            "Investigate the difference and verify "
                            "whether the ledger amount includes "
                            "multiple bills, adjustments or taxes"
                        )
                    elif match_method == "Amount only - review":
                        status = "Possible match - amount only"
                        action = (
                            "Verify provider, invoice number and "
                            "supporting evidence"
                        )
                    else:
                        status = "Matched - review evidence"
                        action = (
                            "Verify invoice, posting date, account "
                            "classification and payment evidence"
                        )

                    statuses.append(status)
                    matched_rows.append(match_index + 2)
                    matched_amounts.append(ledger_amount)
                    differences.append(difference)
                    actions.append(action)

                else:
                    status = "Potentially missing from ledger"
                    action = (
                        "Search other accounts and periods, "
                        "check payment evidence and confirm the "
                        "accounting basis before adjustment"
                    )

                    statuses.append(status)
                    matched_rows.append("")
                    matched_amounts.append(None)
                    differences.append(None)
                    actions.append(action)

                    # Only draft an accrual if the user selects
                    # accrual basis and the bill is not beyond cutoff.
                    if basis == "Accrual basis":
                        vendor_name = str(bill[b_vendor])
                        bill_ref = (
                            str(bill[b_invoice])
                            if b_invoice and pd.notna(bill[b_invoice])
                            else "No bill reference supplied"
                        )

                        journal_entries.append({
                            "Journal Date": period_end,
                            "Description": (
                                "Proposed accrual for utility bill - "
                                + vendor_name + " - " + bill_ref
                            ),
                            "Account": expense_account,
                            "Debit (Rs.)": abs(float(amount)),
                            "Credit (Rs.)": 0,
                            "Bill Working Row": i + 2,
                            "Review Status": "Draft - approval required"
                        })

                        journal_entries.append({
                            "Journal Date": period_end,
                            "Description": (
                                "Proposed accrual for utility bill - "
                                + vendor_name + " - " + bill_ref
                            ),
                            "Account": payable_account,
                            "Debit (Rs.)": 0,
                            "Credit (Rs.)": abs(float(amount)),
                            "Bill Working Row": i + 2,
                            "Review Status": "Draft - approval required"
                        })

            # Create the utility reconciliation worksheet.
            reconciliation = bills.copy()
            reconciliation["Audit Status"] = statuses
            reconciliation["Matched Ledger Row"] = matched_rows
            reconciliation["Matched Ledger Amount"] = matched_amounts
            reconciliation["Difference (Rs.)"] = differences
            reconciliation["Suggested Audit Action"] = actions

            # Create a copy of the ledger with review status.
            ledger_review = ledger.copy()
            ledger_review["Utility Review Status"] = [
                matched_ledger_status.get(
                    i, "Not matched to utility working - review"
                )
                for i in range(len(ledger))
            ]

            journal_df = pd.DataFrame(journal_entries)

            st.subheader("5. Reconciliation results")

            c1, c2, c3 = st.columns(3)

            c1.metric("Utility Bills", len(reconciliation))

            missing_count = sum(
                status == "Potentially missing from ledger"
                for status in statuses
            )
            c2.metric("Potentially Missing", missing_count)

            difference_count = sum(
                status == "Matched invoice - amount differs"
                for status in statuses
            )
            c3.metric("Amount Differences", difference_count)

            st.dataframe(
                reconciliation,
                use_container_width=True
            )

            if basis != "Accrual basis":
                st.info(
                    "No proposed accrual entries have been generated. "
                    "Confirm the client's accounting basis and "
                    "reporting requirements before proposing adjustments."
                )

            if not journal_df.empty:
                st.subheader("Proposed correcting journal entries")
                st.dataframe(
                    journal_df,
                    use_container_width=True
                )
                st.caption(
                    "These are draft entries only. They must be "
                    "reviewed and approved before posting."
                )

            # Generate workbook with separate review sheets.
            output = BytesIO()

            with pd.ExcelWriter(
                output, engine="openpyxl"
            ) as writer:

                reconciliation.to_excel(
                    writer,
                    index=False,
                    sheet_name="Utility Reconciliation"
                )

                ledger_review.to_excel(
                    writer,
                    index=False,
                    sheet_name="Ledger Review"
                )

                if journal_df.empty:
                    pd.DataFrame({
                        "Information": [
                            "No proposed journal entries generated.",
                            "Review utility exceptions and confirm "
                            "the accounting treatment."
                        ]
                    }).to_excel(
                        writer,
                        index=False,
                        sheet_name="Proposed Journal Entries"
                    )
                else:
                    journal_df.to_excel(
                        writer,
                        index=False,
                        sheet_name="Proposed Journal Entries"
                    )

            output.seek(0)

            # Highlight exceptions in the Excel workbook.
            workbook = load_workbook(output)

            red_fill = PatternFill(
                "solid", fgColor="FFC7CE"
            )
            yellow_fill = PatternFill(
                "solid", fgColor="FFEB9C"
            )
            green_fill = PatternFill(
                "solid", fgColor="C6EFCE"
            )
            header_fill = PatternFill(
                "solid", fgColor="1F4E78"
            )

            for sheet in workbook.worksheets:
                sheet.freeze_panes = "A2"
                sheet.auto_filter.ref = sheet.dimensions

                for cell in sheet[1]:
                    cell.fill = header_fill
                    cell.font = Font(
                        color="FFFFFF", bold=True
                    )
                    cell.alignment = Alignment(
                        vertical="center",
                        wrap_text=True
                    )

                for column_cells in sheet.columns:
                    column_letter = get_column_letter(
                        column_cells[0].column
                    )
                    max_length = max(
                        len(str(cell.value or ""))
                        for cell in column_cells
                    )
                    sheet.column_dimensions[column_letter].width = min(
                        max(max_length + 2, 12), 45
                    )

            sheet = workbook["Utility Reconciliation"]
            status_column = next(
                (
                    cell.column
                    for cell in sheet[1]
                    if cell.value == "Audit Status"
                ),
                None
            )

            if status_column:
                for row in range(2, sheet.max_row + 1):
                    status = str(
                        sheet.cell(row, status_column).value or ""
                    )

                    if (
                        "missing" in status.lower()
                        or "differs" in status.lower()
                        or "invalid" in status.lower()
                    ):
                        fill = red_fill
                    elif (
                        "outside" in status.lower()
                        or "possible match" in status.lower()
                    ):
                        fill = yellow_fill
                    else:
                        fill = green_fill

                    for cell in sheet[row]:
                        cell.fill = fill

            journal_sheet = workbook["Proposed Journal Entries"]

            for row in range(2, journal_sheet.max_row + 1):
                for cell in journal_sheet[row]:
                    cell.fill = yellow_fill

            final_file = BytesIO()
            workbook.save(final_file)
            final_file.seek(0)

            st.download_button(
                "Download Utility Audit Report",
                data=final_file.getvalue(),
                file_name="utility_audit_report.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )

            st.success(
                "Reconciliation completed. Download the workbook "
                "and review all flagged items before making adjustments."
            )
