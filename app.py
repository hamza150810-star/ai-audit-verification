
st.subheader("Automatic Ledger Analysis")

if ledger_df is not None and not ledger_df.empty:
    try:
        auto_ledger, detected_columns = prepare_detected_ledger(
            ledger_df
        )

        st.success("Ledger columns detected automatically.")

        detected_rows = []

        for field, details in detected_columns.items():
            detected_rows.append({
                "Field": field.replace("_", " ").title(),
                "Detected Column": (
                    str(details["column"])
                    if details["column"] is not None
                    else "Not detected"
                ),
                "Confidence": (
                    "Review"
                    if details["ambiguous"]
                    else "Candidate identified"
                    if details["column"] is not None
                    else "Missing"
                )
            })

        st.dataframe(
            pd.DataFrame(detected_rows),
            use_container_width=True,
            hide_index=True
        )

        st.caption(
            "The app selects columns automatically. It flags uncertain "
            "matches so the wrong column is not silently treated as correct."
        )

        if st.button(
            "Generate Sampling Workbook Automatically",
            type="primary"
        ):
            st.info(
                "Pass auto_ledger['_Auto_Account'] as the account-head "
                "field and auto_ledger['_Auto_Amount'] as the amount field "
                "to the sampling generator. Use the detected date, voucher "
                "and description fields for the audit trail."
            )

    except Exception as exc:
        st.error(f"Automatic ledger detection failed: {exc}")

import re
import pandas as pd


def _normalise_header(value):
    return re.sub(
        r"[^a-z0-9]+", " ", str(value).lower()
    ).strip()


def _header_score(header, aliases):
    h = _normalise_header(header)
    score = 0

    for alias in aliases:
        a = _normalise_header(alias)

        if h == a:
            score = max(score, 100)
        elif a and a in h:
            score = max(score, 75)
        elif h and h in a:
            score = max(score, 45)

    return score


def detect_ledger_columns(df):
    """
    Detect ledger columns automatically.
    Returns the best candidate for each field and a confidence score.
    """
    if df is None or df.empty:
        raise ValueError("The uploaded ledger is empty.")

    aliases = {
        "account_name": [
            "account name", "ledger name", "ledger account",
            "account head", "expense head", "particulars",
            "account description", "gl name", "gl description",
            "account title", "name of account"
        ],
        "account_code": [
            "account code", "gl code", "ledger code",
            "account number", "account no", "gl account"
        ],
        "amount": [
            "amount", "transaction amount", "local amount",
            "debit amount", "credit amount", "net amount",
            "value", "transaction value", "debit credit"
        ],
        "debit": [
            "debit", "debit amount", "dr amount",
            "debit value", "debit local"
        ],
        "credit": [
            "credit", "credit amount", "cr amount",
            "credit value", "credit local"
        ],
        "date": [
            "date", "voucher date", "posting date",
            "transaction date", "document date", "entry date"
        ],
        "voucher": [
            "voucher number", "voucher no", "voucher",
            "document number", "document no", "journal number",
            "journal no", "reference number", "reference no"
        ],
        "description": [
            "description", "narration", "details",
            "transaction details", "particulars",
            "remarks", "memo", "line description"
        ],
    }

    def numeric_ratio(series):
        values = series.dropna()
        if values.empty:
            return 0.0

        def is_numeric(value):
            if isinstance(value, (int, float)):
                return True

            text = str(value).strip()
            if not text:
                return False

            text = text.replace(",", "").replace("(", "").replace(")", "")
            text = re.sub(r"[^0-9.\-]", "", text)

            try:
                float(text)
                return True
            except ValueError:
                return False

        return sum(is_numeric(v) for v in values) / len(values)

    def date_ratio(series):
        values = series.dropna()
        if values.empty:
            return 0.0

        parsed = pd.to_datetime(values, errors="coerce")
        return float(parsed.notna().mean())

    candidates = {}

    for field, field_aliases in aliases.items():
        scored = []

        for column in df.columns:
            score = _header_score(column, field_aliases)
            series = df[column]

            if field in {"amount", "debit", "credit"}:
                ratio = numeric_ratio(series)

                if ratio >= 0.70:
                    score += 20
                elif ratio >= 0.40:
                    score += 8

                if series.nunique(dropna=True) <= 1:
                    score -= 10

            elif field == "date":
                ratio = date_ratio(series)
                if ratio >= 0.70:
                    score += 20

            elif field in {"account_name", "description"}:
                nonempty = series.dropna().astype(str).str.strip()
                if not nonempty.empty:
                    text_ratio = (
                        ~nonempty.str.fullmatch(
                            r"[-+]?\d[\d,]*(\.\d+)?"
                        )
                    ).mean()
                    score += int(text_ratio * 10)

            if score > 0:
                scored.append((score, column))

        scored.sort(key=lambda item: item[0], reverse=True)

        if scored:
            best_score, best_column = scored[0]
            second_score = scored[1][0] if len(scored) > 1 else 0

            candidates[field] = {
                "column": best_column,
                "score": best_score,
                "ambiguous": (
                    second_score > 0
                    and best_score - second_score < 15
                )
            }
        else:
            candidates[field] = {
                "column": None,
                "score": 0,
                "ambiguous": False
            }

    # Prefer a separate account-name column over an account code.
    if candidates["account_name"]["column"] is None:
        candidates["account_name"] = candidates["account_code"].copy()

    # A general "amount" match should not override separate debit/credit
    # columns unless those separate columns were not detected.
    return candidates


def prepare_detected_ledger(df):
    """
    Return a standardised ledger and detected-column report.
    Original ledger columns are retained.
    """
    result = df.copy()
    detected = detect_ledger_columns(result)

    account_col = detected["account_name"]["column"]
    amount_col = detected["amount"]["column"]
    debit_col = detected["debit"]["column"]
    credit_col = detected["credit"]["column"]

    if account_col is None:
        raise ValueError(
            "Could not confidently identify the account/ledger name column. "
            "Check that the ledger includes its account name or account code."
        )

    # Prefer a real amount column; otherwise derive a net transaction amount
    # from separate debit and credit columns.
    if amount_col is not None:
        result["_Auto_Amount"] = result[amount_col].map(parse_numeric)

    elif debit_col is not None or credit_col is not None:
        debit = (
            result[debit_col].map(parse_numeric)
            if debit_col is not None
            else pd.Series(0.0, index=result.index)
        )
        credit = (
            result[credit_col].map(parse_numeric)
            if credit_col is not None
            else pd.Series(0.0, index=result.index)
        )

        result["_Auto_Amount"] = (
            debit.fillna(0) - credit.fillna(0)
        )

    else:
        raise ValueError(
            "Could not identify an amount, debit, or credit column."
        )

    result["_Auto_Account"] = (
        result[account_col].fillna("").astype(str).str.strip()
    )

    result["_Auto_Source_Row"] = range(2, len(result) + 2)

    for field in ["date", "voucher", "description"]:
        column = detected[field]["column"]
        result[f"_Auto_{field.title()}"] = (
            result[column] if column is not None
            else ""
        )

    return result, detected

st.subheader("Generate audit sampling workbook")

if ledger_df is not None and reference_file is not None:
    account_col = st.selectbox(
        "Ledger account-head column",
        ledger_df.columns,
        key="template_account_col"
    )

    amount_col = st.selectbox(
        "Ledger amount column",
        ledger_df.columns,
        key="template_amount_col"
    )

    date_col = st.selectbox(
        "Ledger date column (optional)",
        ["None"] + list(ledger_df.columns),
        key="template_date_col"
    )

    voucher_col = st.selectbox(
        "Voucher column (optional)",
        ["None"] + list(ledger_df.columns),
        key="template_voucher_col"
    )

    description_col = st.selectbox(
        "Description column (optional)",
        ["None"] + list(ledger_df.columns),
        key="template_description_col"
    )

    heads_text = st.text_area(
        "100% verification account heads — one per line",
        value=(
            "Utilities\n"
            "Electricity\n"
            "Water\n"
            "Gas\n"
            "Fee and Subscription\n"
            "Legal and Professional Charges\n"
            "Salaries\n"
            "Depreciation"
        ),
        key="template_full_verification_heads"
    )

    st.checkbox(
        "Allow partial account-head matching",
        value=False,
        key="sampling_allow_partial_match"
    )

    percentage = st.number_input(
        "Sampling percentage for remaining population",
        min_value=1,
        max_value=100,
        value=20,
        step=1,
        key="template_sampling_percentage"
    )

    high_value = st.checkbox(
        "Select high-value items first",
        value=True,
        key="template_high_value"
    )

    if st.button("Create Sampling Workbook", type="primary"):
        try:
            excel_bytes = build_audit_sampling_workbook(
                ledger_df=ledger_df,
                reference_bytes=reference_file.getvalue(),
                ledger_account_col=account_col,
                ledger_amount_col=amount_col,
                ledger_date_col=(
                    None if date_col == "None" else date_col
                ),
                ledger_voucher_col=(
                    None if voucher_col == "None" else voucher_col
                ),
                ledger_description_col=(
                    None if description_col == "None"
                    else description_col
                ),
                full_verification_heads=[
                    h.strip()
                    for h in re.split(r"[\n,;]+", heads_text)
                    if h.strip()
                ],
                sample_percentage=int(percentage),
                include_high_value=high_value,
            )

            st.download_button(
                "Download Audit Sampling Workbook",
                data=excel_bytes,
                file_name="PEF_Expense_Sampling_Workbook.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary"
            )

            st.success("Workbook generated. Review it before audit use.")

        except Exception as exc:
            st.error(f"Sampling workbook could not be generated: {exc}")

else:
    st.info(
        "Upload both the General Ledger and the reference sampling "
        "workbook to continue."
    )
