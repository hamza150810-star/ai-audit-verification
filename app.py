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
        df = pd.read_excel(uploaded_file)

        st.subheader("Ledger Preview")
        st.dataframe(df.head(10), use_container_width=True)

        if st.button("Analyze Ledger", type="primary"):

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
                    "Your Excel file needs columns named "
                    "Vendor, Invoice, and Amount."
                )
            else:
                df["Amount"] = pd.to_numeric(
                    df["Amount"], errors="coerce"
                )

                df["Duplicate Invoice"] = (
                    df["Invoice"].notna()
                    & df["Vendor"].notna()
                    & df.duplicated(
                        subset=["Vendor", "Invoice"],
                        keep=False
                    )
                )

                df["High Value"] = (
                    df["Amount"] >= threshold
                )

                df["Audit Finding"] = df.apply(
                    lambda row: "; ".join(
                        finding for finding, condition in [
                            (
                                "Potential duplicate invoice",
                                row["Duplicate Invoice"]
                            ),
                            (
                                "High-value transaction",
                                row["High Value"]
                            )
                        ] if condition
                    ),
                    axis=1
                )

                exceptions = df[
                    df["Audit Finding"] != ""
                ].copy()

                col1, col2, col3 = st.columns(3)
                col1.metric("Total Transactions", len(df))
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
                    "Flags are audit indicators, not proof of "
                    "fraud or misstatement. Review supporting "
                    "evidence before reaching conclusions."
                )

    except Exception as error:
        st.error(
            "The file could not be processed. "
            "Please check that it is a valid Excel workbook."
        )
        st.caption(str(error))
