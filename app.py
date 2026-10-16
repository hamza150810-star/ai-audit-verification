
import io
import re
import random
from datetime import date, datetime

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="AI Audit Verification",
    page_icon="📊",
    layout="wide",
)

st.title("AI Audit Verification")
st.caption(
    "Automatic General Ledger analysis and audit sampling workbook generator"
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def normalise_text(value):
    if value is None:
        return ""

    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_numeric(value):
    """Convert common accounting number formats to numeric values."""
    if value is None or pd.isna(value):
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    if not text or text in {"-", "—", "–"}:
        return 0.0

    negative = text.startswith("(") and text.endswith(")")
    text = text.replace(",", "").replace(" ", "")
    text = re.sub(r"[^0-9.\-]", "", text)

    if text in {"", "-", ".", "-."}:
        return 0.0

    try:
        number = float(text)
        return -abs(number) if negative else number
    except (ValueError, TypeError):
        return 0.0


def safe_datetime_series(series):
    """
    Parse mixed date formats and time zones safely.

    utc=True normalises timezone-aware values and allows mixed
    timezone inputs to be processed consistently.
    """
    try:
        return pd.to_datetime(
            series,
            errors="coerce",
            utc=True,
            format="mixed",
        )
    except (TypeError, ValueError):
        # Compatibility fallback for older pandas versions.
        try:
            return pd.to_datetime(
                series,
                errors="coerce",
                utc=True,
            )
        except (TypeError, ValueError):
            return pd.Series(
                pd.NaT,
                index=series.index,
                dtype="datetime64[ns, UTC]",
            )


def read_uploaded_ledger(uploaded_file):
    """Read Excel or CSV General Ledger files."""
    if uploaded_file is None:
        return None

    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if name.endswith(".csv"):
        try:
            return pd.read_csv(io.BytesIO(data))
        except UnicodeDecodeError:
            return pd.read_csv(
                io.BytesIO(data),
                encoding="latin1",
            )

    if name.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(
            io.BytesIO(data),
            sheet_name=0,
        )

    raise ValueError("Please upload an Excel or CSV ledger.")


# ============================================================
# AUTOMATIC COLUMN DETECTION
# ============================================================

HEADER_ALIASES = {
    "account_name": [
        "account name", "ledger name", "ledger account",
        "account head", "expense head", "account title",
        "account description", "gl description", "gl name",
        "name of account", "ledger description", "particulars",
        "account"
    ],
    "account_code": [
        "account code", "gl code", "ledger code",
        "account number", "account no", "gl account code"
    ],
    "amount": [
        "amount", "transaction amount", "local amount",
        "net amount", "transaction value", "value",
        "total amount"
    ],
    "debit": [
        "debit", "debit amount", "dr amount",
        "debit value", "debit local", "debit balance"
    ],
    "credit": [
        "credit", "credit amount", "cr amount",
        "credit value", "credit local", "credit balance"
    ],
    "date": [
        "date", "voucher date", "posting date",
        "transaction date", "document date", "entry date"
    ],
    "voucher": [
        "voucher number", "voucher no", "voucher",
        "document number", "document no", "journal number",
        "journal no", "reference number", "reference no",
        "transaction id", "document reference"
    ],
    "description": [
        "description", "narration", "transaction details",
        "remarks", "memo", "line description",
        "transaction description", "details"
    ],
}


def header_score(column, aliases):
    header = normalise_text(column)
    best = 0

    for alias in aliases:
        candidate = normalise_text(alias)

        if header == candidate:
            best = max(best, 100)
        elif candidate and candidate in header:
            best = max(best, 75)

    return best


def numeric_ratio(series):
    values = series.dropna()

    if values.empty:
        return 0.0

    numeric_count = sum(
        isinstance(value, (int, float))
        or bool(
            re.fullmatch(
                r"[\s$£€\-+()]?\d[\d,]*(\.\d+)?[\s$£€()]?",
                str(value).strip(),
            )
        )
        for value in values
    )

    return numeric_count / len(values)


def detect_ledger_columns(df):
    """Detect likely ledger columns from headers and data."""
    if df is None or df.empty:
        raise ValueError("The uploaded General Ledger is empty.")

    detected = {}

    for field, aliases in HEADER_ALIASES.items():
        candidates = []

        for column in df.columns:
            score = header_score(column, aliases)
            values = df[column].dropna()

            if field in {"amount", "debit", "credit"}:
                ratio = numeric_ratio(values)

                if ratio >= 0.70:
                    score += 20
                elif ratio >= 0.40:
                    score += 8

            elif field == "date" and not values.empty:
                # FIX: utc=True handles mixed time zones.
                parsed_dates = safe_datetime_series(values)
                date_ratio = parsed_dates.notna().mean()

                if date_ratio >= 0.70:
                    score += 20

            elif field in {"account_name", "description"}:
                if not values.empty:
                    text_ratio = (
                        ~values.astype(str).str.fullmatch(
                            r"[-+]?\d[\d,]*(\.\d+)?"
                        )
                    ).mean()

                    score += int(text_ratio * 10)

            if score > 0:
                candidates.append((score, column))

        candidates.sort(
            key=lambda item: item[0],
            reverse=True,
        )

        if candidates:
            best_score, best_column = candidates[0]
            second_score = (
                candidates[1][0]
                if len(candidates) > 1
                else 0
            )

            detected[field] = {
                "column": best_column,
                "score": best_score,
                "ambiguous": (
                    second_score > 0
                    and best_score - second_score < 15
                ),
            }
        else:
            detected[field] = {
                "column": None,
                "score": 0,
                "ambiguous": False,
            }

    if detected["account_name"]["column"] is None:
        detected["account_name"] = detected["account_code"].copy()

    return detected


def prepare_detected_ledger(df):
    """Standardise ledger fields without removing source columns."""
    result = df.copy()
    detected = detect_ledger_columns(result)

    account_col = detected["account_name"]["column"]
    amount_col = detected["amount"]["column"]
    debit_col = detected["debit"]["column"]
    credit_col = detected["credit"]["column"]

    if account_col is None:
        raise ValueError(
            "Account-head column could not be detected. "
            "Check the ledger headers."
        )

    if amount_col is not None:
        result["_Audit_Amount"] = result[amount_col].map(parse_numeric)

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

        result["_Audit_Amount"] = debit - credit

    else:
        raise ValueError(
            "No amount, debit, or credit column was detected."
        )

    result["_Audit_Account"] = (
        result[account_col]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    result["_Audit_Source_Row"] = range(2, len(result) + 2)

    for field in ["date", "voucher", "description"]:
        source_column = detected[field]["column"]

        result[f"_Audit_{field.title()}"] = (
            result[source_column]
            if source_column is not None
            else ""
        )

    result = result[
        result["_Audit_Account"].ne("")
    ].copy()

    return result, detected


# ============================================================
# FULL-VERIFICATION ACCOUNT HEADS
# ============================================================

FULL_VERIFICATION_KEYWORDS = {
    "Utilities": [
        "utility", "utilities", "electricity", "lesco",
        "iesco", "mepco", "electric bill", "water bill",
        "wasa", "sui gas", "gas bill", "telephone", "internet"
    ],
    "Fee and Subscription": [
        "fee and subscription", "fees and subscription",
        "subscription", "membership fee",
        "professional membership", "fee subscription"
    ],
    "Legal and Professional Charges": [
        "legal and professional", "legal charges",
        "professional charges", "legal fee",
        "consultancy fee", "consultancy charges", "audit fee"
    ],
    "Salaries": [
        "salary", "salaries", "payroll",
        "staff remuneration", "employees remuneration", "wages"
    ],
    "Depreciation": [
        "depreciation", "depreciation expense",
        "amortisation", "amortization"
    ],
}


def classify_full_verification(account_name, extra_heads=None):
    account = normalise_text(account_name)

    for category, keywords in FULL_VERIFICATION_KEYWORDS.items():
        for keyword in keywords:
            if normalise_text(keyword) in account:
                return category

    for head in extra_heads or []:
        normalised_head = normalise_text(head)

        if normalised_head and normalised_head in account:
            return head

    return None


# ============================================================
# SAMPLING
# ============================================================

def build_sampling_data(
    ledger,
    sample_percentage=20,
    include_high_value=True,
    extra_heads=None,
    seed=42,
):
    df = ledger.copy()

    df["_Full_Verification_Category"] = df[
        "_Audit_Account"
    ].map(
        lambda account: classify_full_verification(
            account,
            extra_heads=extra_heads,
        )
    )

    df["_Absolute_Amount"] = df["_Audit_Amount"].abs()

    full_verification = df[
        df["_Full_Verification_Category"].notna()
    ].copy()

    remaining = df[
        df["_Full_Verification_Category"].isna()
    ].copy()

    selected_indices = []
    rng = random.Random(seed)

    for _, group in remaining.groupby(
        "_Audit_Account",
        dropna=False,
    ):
        if group.empty:
            continue

        sample_count = max(
            1,
            int(len(group) * sample_percentage / 100),
        )
        sample_count = min(sample_count, len(group))

        if include_high_value:
            group = group.sort_values(
                "_Absolute_Amount",
                ascending=False,
            )

            chosen = list(
                group.head(sample_count).index
            )

            if len(chosen) < sample_count:
                remaining_indices = [
                    idx for idx in group.index
                    if idx not in chosen
                ]

                chosen.extend(
                    rng.sample(
                        remaining_indices,
                        min(
                            sample_count - len(chosen),
                            len(remaining_indices),
                        ),
                    )
                )
        else:
            chosen = rng.sample(
                list(group.index),
                sample_count,
            )

        selected_indices.extend(chosen)

    sample_selected = remaining.loc[
        selected_indices
    ].copy()

    population_not_selected = remaining.drop(
        index=selected_indices,
        errors="ignore",
    ).copy()

    summary = (
        df.groupby("_Audit_Account", dropna=False)
        .agg(
            Population_Items=("_Audit_Account", "size"),
            Population_Net_Amount=("_Audit_Amount", "sum"),
            Population_Gross_Amount=("_Absolute_Amount", "sum"),
        )
        .reset_index()
        .rename(columns={"_Audit_Account": "Account Head"})
    )

    return (
        full_verification,
        sample_selected,
        population_not_selected,
        summary,
    )


# ============================================================
# EXCEL OUTPUT
# ============================================================

def dataframe_for_export(df):
    return df.drop(
        columns=[
            col for col in df.columns
            if col.startswith("_Audit_")
            or col == "_Absolute_Amount"
        ],
        errors="ignore",
    ).copy()


def style_worksheet(ws):
    ws.freeze_panes = "A2"

    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="17365D",
    )
    header_font = Font(
        bold=True,
        color="FFFFFF",
    )
    thin_grey = Side(
        style="thin",
        color="D9E2F3",
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        cell.border = Border(bottom=thin_grey)

    for column_cells in ws.columns:
        letter = get_column_letter(column_cells[0].column)
        max_length = 0

        for cell in column_cells:
            if cell.value is not None:
                max_length = max(
                    max_length,
                    len(str(cell.value)),
                )

        ws.column_dimensions[letter].width = min(
            max(max_length + 2, 12),
            40,
        )

    ws.row_dimensions[1].height = 30


def build_audit_sampling_workbook(
    ledger_df,
    reference_bytes,
    sample_percentage=20,
    include_high_value=True,
    extra_heads=None,
):
    if not reference_bytes:
        raise ValueError(
            "Upload the reference sampling workbook."
        )

    ledger, detected = prepare_detected_ledger(ledger_df)

    full, selected, not_selected, summary = build_sampling_data(
        ledger,
        sample_percentage=sample_percentage,
        include_high_value=include_high_value,
        extra_heads=extra_heads,
    )

    workbook = load_workbook(io.BytesIO(reference_bytes))

    output_names = [
        "Auto Summary",
        "100% Verification",
        "Sample Selected",
        "Population Summary",
        "Unselected Population",
        "Detection Report",
    ]

    for sheet_name in output_names:
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]

    def add_dataframe_sheet(sheet_name, df):
        ws = workbook.create_sheet(sheet_name)
        export_df = dataframe_for_export(df)

        ws.append(list(export_df.columns))

        for values in export_df.itertuples(
            index=False,
            name=None,
        ):
            cleaned_values = []

            for value in values:
                if value is None:
                    cleaned_values.append(None)
                elif isinstance(value, (pd.Timestamp, datetime, date)):
                    cleaned_values.append(
                        value.to_pydatetime()
                        if isinstance(value, pd.Timestamp)
                        else value
                    )
                elif pd.isna(value):
                    cleaned_values.append(None)
                elif hasattr(value, "item"):
                    try:
                        cleaned_values.append(value.item())
                    except Exception:
                        cleaned_values.append(value)
                else:
                    cleaned_values.append(value)

            ws.append(cleaned_values)

        style_worksheet(ws)

    # Summary sheet.
    ws = workbook.create_sheet("Auto Summary")
    ws.append(["AI AUDIT VERIFICATION - SAMPLING SUMMARY"])
    ws.append(["Ledger transactions analysed", len(ledger)])
    ws.append(["Full-verification transactions", len(full)])
    ws.append(["Risk-based sample transactions", len(selected)])
    ws.append(["Unselected population transactions", len(not_selected)])
    ws.append(["Sampling percentage", sample_percentage])
    ws.append([
        "Selection method",
        "Keyword-based full verification and percentage-based sampling",
    ])
    ws.append([
        "Review note",
        "Review classifications and selected items before audit use.",
    ])

    ws["A1"].font = Font(
        bold=True,
        color="FFFFFF",
        size=14,
    )
    ws["A1"].fill = PatternFill(
        fill_type="solid",
        fgColor="17365D",
    )
    ws.merge_cells("A1:B1")
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 90

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True,
            )

    add_dataframe_sheet(
        "100% Verification",
        full.sort_values(
            ["_Full_Verification_Category", "_Audit_Account"]
        ) if not full.empty else full,
    )

    add_dataframe_sheet("Sample Selected", selected)
    add_dataframe_sheet("Population Summary", summary)
    add_dataframe_sheet("Unselected Population", not_selected)

    detection_rows = []

    for field, details in detected.items():
        detection_rows.append({
            "Field": field.replace("_", " ").title(),
            "Detected Column": (
                str(details["column"])
                if details["column"] is not None
                else "Not detected"
            ),
            "Confidence Score": details["score"],
            "Review Required": (
                "Yes" if details["ambiguous"] else "No"
            ),
        })

    add_dataframe_sheet(
        "Detection Report",
        pd.DataFrame(detection_rows),
    )

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return (
        output.getvalue(),
        detected,
        len(full),
        len(selected),
    )


# ============================================================
# STREAMLIT INTERFACE
# ============================================================

st.header("1. Upload Files")

left, right = st.columns(2)

with left:
    ledger_file = st.file_uploader(
        "Upload General Ledger",
        type=["xlsx", "xlsm", "xls", "csv"],
        key="ledger_upload",
    )

with right:
    reference_file = st.file_uploader(
        "Upload Reference Sampling Workbook",
        type=["xlsx", "xlsm"],
        key="reference_upload",
    )

ledger_df = None

if ledger_file is not None:
    try:
        ledger_df = read_uploaded_ledger(ledger_file)

        st.success(
            f"Ledger loaded: {len(ledger_df):,} rows and "
            f"{len(ledger_df.columns):,} columns."
        )

        with st.expander("Preview General Ledger"):
            st.dataframe(
                ledger_df.head(10),
                use_container_width=True,
            )

    except Exception as exc:
        st.error(f"Could not read the ledger: {exc}")

if reference_file is not None:
    st.success(
        f"Reference workbook uploaded: {reference_file.name}"
    )

st.divider()
st.header("2. Automatic Ledger Analysis")

if ledger_df is not None and not ledger_df.empty:
    try:
        prepared_ledger, detected_columns = prepare_detected_ledger(
            ledger_df
        )

        detection_rows = []

        for field, details in detected_columns.items():
            detection_rows.append({
                "Field": field.replace("_", " ").title(),
                "Detected Column": (
                    str(details["column"])
                    if details["column"] is not None
                    else "Not detected"
                ),
                "Confidence Score": details["score"],
                "Status": (
                    "Review recommended"
                    if details["ambiguous"]
                    else "Detected"
                    if details["column"] is not None
                    else "Missing"
                ),
            })

        st.dataframe(
            pd.DataFrame(detection_rows),
            use_container_width=True,
            hide_index=True,
        )

        st.caption(
            "Dates are parsed with timezone normalisation to prevent "
            "mixed-timezone errors."
        )

        with st.expander("Review detected account heads"):
            heads_preview = (
                prepared_ledger["_Audit_Account"]
                .value_counts()
                .rename_axis("Account Head")
                .reset_index(name="Transaction Count")
            )

            st.dataframe(
                heads_preview,
                use_container_width=True,
                hide_index=True,
            )

        st.divider()
        st.header("3. Sampling Settings")

        col1, col2 = st.columns(2)

        with col1:
            sample_percentage = st.number_input(
                "Sampling percentage for other account heads",
                min_value=1,
                max_value=100,
                value=20,
                step=1,
            )

        with col2:
            include_high_value = st.checkbox(
                "Prioritise high-value transactions",
                value=True,
            )

        st.write(
            "**Full-verification categories:** Utilities, Fee and "
            "Subscription, Legal and Professional Charges, Salaries, "
            "and Depreciation."
        )

        st.caption(
            "Categories are matched using account-name keywords. "
            "Review the classification against your audit plan."
        )

        if st.button(
            "Generate Audit Sampling Workbook",
            type="primary",
            use_container_width=True,
        ):
            if reference_file is None:
                st.error(
                    "Please upload the reference sampling workbook."
                )
            else:
                try:
                    with st.spinner(
                        "Generating the audit sampling workbook..."
                    ):
                        excel_bytes, detected, full_count, sample_count = (
                            build_audit_sampling_workbook(
                                ledger_df=ledger_df,
                                reference_bytes=reference_file.getvalue(),
                                sample_percentage=int(sample_percentage),
                                include_high_value=include_high_value,
                            )
                        )

                    st.success(
                        f"Workbook generated. Full-verification items: "
                        f"{full_count:,}; sampled items: {sample_count:,}."
                    )

                    st.download_button(
                        "Download Audit Sampling Workbook",
                        data=excel_bytes,
                        file_name="Audit_Sampling_Workbook.xlsx",
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        type="primary",
                        use_container_width=True,
                    )

                except Exception as exc:
                    st.error(
                        f"Workbook generation failed: {exc}"
                    )

    except Exception as exc:
        st.error(
            f"Automatic ledger analysis failed: {exc}"
        )

else:
    st.info(
        "Upload a General Ledger to start automatic analysis."
    )

st.divider()

st.caption(
    "AI Audit Verification | Always review the selected population "
    "and supporting evidence before relying on audit results."
)
