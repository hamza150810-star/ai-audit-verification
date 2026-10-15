
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="AI Audit Verification",
    page_icon="📊",
    layout="wide"
)

st.title("AI Audit Verification")
st.write(
    "Upload a general ledger to identify potential "
    "duplicate invoices and high-value transactions."
)

threshold = st.number_input(
    "High-value transaction threshold",
    min_value=0,
    value=50000,
    step=5000
)

uploaded_file = st.file_uploader(
    "Upload your ledger (Excel file)",
    type=["xlsx"]
)

if uploaded_file is not None:
    try:
        # Read the uploaded Excel workbook
        df = pd.read_excel(
            uploaded_file,
            engine="openpyxl"
        )

        st.subheader("Ledger Preview")
        st.dataframe(df.head(10), use_container_width=True)

        if st.button("Analyze Ledger", type="primary"):

            # Clean column names to remove extra spaces
            df.columns = df.columns.astype(str).str.strip()

            # Check required columns
            required = ["Vendor", "Invoice", "Amount"]
            missing = [
                col for col in required
                if col not in df.columns
            ]

            if missing:
                st.error(
                    "Missing required columns: "
                    + ", ".join(missing)
                )
                st.info(
                    "Your Excel file must contain columns "
                    "named Vendor, Invoice, and Amount. "
                    "Column names can have extra spaces, "
                    "but must otherwise match."
                )

            else:
                # Convert amounts into numeric values
                df["Amount"] = pd.to_numeric(
                    df["Amount"]
                    .astype(str)
                    .str.replace(",", "", regex=False)
                    .str.replace(" ", "", regex=False),
                    errors="coerce"
                )

                # Identify repeated vendor and invoice combinations
                vendor = (
                    df["Vendor"].astype("string").str.strip()
                )
                invoice = (
                    df["Invoice"].astype("string").str.strip()
                )

                valid_invoice = (
                    vendor.notna()
                    & invoice.notna()
                    & vendor.ne("")
                    & invoice.ne("")
                )

                duplicate_mask = pd.Series(
                    False, index=df.index
                )

                duplicate_mask.loc[valid_invoice] = (
                    df.loc[valid_invoice]
                    .assign(
                        _Vendor=vendor.loc[valid_invoice],
                        _Invoice=invoice.loc[valid_invoice]
                    )
                    .duplicated(
                        subset=["_Vendor", "_Invoice"],
                        keep=False
                    )
                    .to_numpy()
                )

                df["Duplicate Invoice"] = duplicate_mask

                # Identify high-value transactions
                df["High Value"] = (
                    df["Amount"].notna()
                    & (df["Amount"] >= threshold)
                )

                # Describe audit findings
                df["Audit Finding"] = df.apply(
                    lambda row: "; ".join(
                        finding
                        for finding, condition in [
                            (
                                "Potential duplicate invoice",
                                row["Duplicate Invoice"]
                            ),
                            (
                                "High-value transaction",
                                row["High Value"]
                            )
                        ]
                        if condition
                    ),
                    axis=1
                )

                # Prepare exception report
                exceptions = df[
                    df["Audit Finding"] != ""
                ].copy()

                # Display summary metrics
                col1, col2, col3 = st.columns(3)

                col1.metric(
                    "Total Transactions", len(df)
                )
                col2.metric(
                    "Flagged Transactions", len(exceptions)
                )
                col3.metric(
                    "Potential Duplicate Rows",
                    int(df["Duplicate Invoice"].sum())
                )

                st.subheader("Audit Exceptions")

                if exceptions.empty:
                    st.success(
                        "No exceptions were flagged by these tests."
                    )
                else:
                    st.dataframe(
                        exceptions,
                        use_container_width=True
                    )

                # Generate downloadable Excel report
                output = BytesIO()

                with pd.ExcelWriter(
                    output, engine="openpyxl"
                ) as writer:
                    exceptions.to_excel(
                        writer,
                        index=False,
                        sheet_name="Audit Exceptions"
                    )

                st.download_button(
                    label="Download Audit Report",
                    data=output.getvalue(),
                    file_name="audit_exceptions.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    )
                )

                st.caption(
                    "These flags are audit indicators, not proof "
                    "of fraud or misstatement. Review supporting "
                    "evidence before reaching conclusions."
                )

    except Exception as error:
        st.error(
            "The Excel file could not be processed. "
            "Check that it is a valid, unencrypted Excel workbook."
        )
        st.write("Technical error details:")
        st.exception(error)
