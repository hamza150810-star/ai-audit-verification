
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
    "Automatically analyse a General Ledger, identify account heads, "
    "select audit samples, and export an Excel workbook."
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def normalise_text(value):
    """Standardise text for matching headers and account names."""
    if value is None:
        return ""

    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_numeric(value):
    """Convert common accounting number formats into numeric values."""
    if pd.isna(value):
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
    except ValueError:
        return 0.0


def read_uploaded_ledger(uploaded_file):
    """Read an Excel or CSV ledger."""
    if uploaded_file is None:
        return None

    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()

    if name.endswith(".csv"):
        try:
            return pd.read_csv(io.BytesIO(data))
        except UnicodeDecodeError:
            return pd.read_csv(io.BytesIO(data), encoding="latin1")

    if name.endswith((".xlsx", ".xlsm")):
        return pd.read_excel(io.BytesIO(data), sheet_name=0)

    if name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(data), sheet_name=0)

    raise ValueError("Upload an Excel or CSV General Ledger file.")


# ============================================================
# AUTOMATIC COLUMN DETECTION
# ============================================================

HEADER_ALIASES = {
    "account_name": [
        "account name", "ledger name", "ledger account",
        "account head", "expense head", "account title",
        "account description", "gl description",
        "gl name", "name of account", "ledger description",
        "particulars", "account"
    ],
    "account_code": [
        "account code", "gl code", "ledger code",
        "account number", "account no", "gl account code"
    ],
    "amount": [
        "amount", "transaction amount", "local amount",
        "net amount", "transaction value", "value",
        "debit credit amount", "total amount"
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

    parsed = values.map(parse_numeric)
    return float(parsed.notna().mean())


def detect_ledger_columns(df):
    """Choose likely columns using header names and column contents."""
    if df is None or df.empty:
        raise ValueError("The uploaded General Ledger is empty.")

    detected = {}

    for field, aliases in HEADER_ALIASES.items():
        candidates = []

        for column in df.columns:
            score = header_score(column, aliases)
            values = df[column].dropna()

            if field in {"amount", "debit", "credit"}:
                if not values.empty:
                    numeric_share = sum(
                        1 for value in values
                        if isinstance(value, (int, float))
                        or bool(
                            re.fullmatch(
                                r"[\s$£€\-+()]?\d[\d,]*(\.\d+)?[\s$£€()]?",
                                str(value).strip()
                            )
                        )
                    ) / len(values)

                    if numeric_share >= 0.70:
                        score += 20

            if field == "date" and not values.empty:
                date_share = pd.to_datetime(
                    values, errors="coerce"
                ).notna().mean()

                if date_share >= 0.70:
                    score += 20

            if field in {"account_name", "description"}:
                if not values.empty:
                    nonnumeric_share = (
                        ~values.astype(str).str.fullmatch(
                            r"[-+]?\d[\d,]*(\.\d+)?"
                        )
                    ).mean()
                    score += int(nonnumeric_share * 10)

            if score > 0:
                candidates.append((score, column))

        candidates.sort(key=lambda item: item[0], reverse=True)

        if candidates:
            best_score, best_column = candidates[0]
            second_score = (
                candidates[1][0] if len(candidates) > 1 else 0
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

    # Account-code fallback if no account-name column was identified.
    if detected["account_name"]["column"] is None:
        detected["account_name"] = detected["account_code"].copy()

    return detected


def prepare_detected_ledger(df):
    """Add standardised audit fields without deleting original columns."""
    result = df.copy()
    detected = detect_ledger_columns(result)

    account_col = detected["account_name"]["column"]
    amount_col = detected["amount"]["column"]
    debit_col = detected["debit"]["column"]
    credit_col = detected["credit"]["column"]

    if account_col is None:
        raise ValueError(
            "Account-head column not detected. Check the ledger headers."
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

        # Debit-positive convention for net transaction value.
        result["_Audit_Amount"] = debit - credit

    else:
        raise ValueError(
            "No amount, debit, or credit column was detected."
        )

    result["_Audit_Account"] = (
        result[account_col].fillna("").astype(str).str.strip()
    )
    result["_Audit_Source_Row"] = range(2, len(result) + 2)

    for field in ["date", "voucher", "description"]:
        source_column = detected[field]["column"]
        result[f"_Audit_{field.title()}"] = (
            result[source_column] if source_column is not None else ""
        )

    # Remove blank account rows and rows without meaningful transactions.
    result = result[
        result["_Audit_Account"].ne("")
    ].copy()

    return result, detected


# ============================================================
# AUDIT HEAD CLASSIFICATION AND SAMPLING
# ============================================================

FULL_VERIFICATION_KEYWORDS = {
    "Utilities": [
        "utility", "utilities", "electricity", "lesco", "iesco",
        "mepco", "electric bill", "water bill", "wasa",
        "sui gas", "gas bill", "telephone", "internet"
    ],
    "Fee and Subscription": [
        "fee and subscription", "fees and subscription",
        "subscription", "membership fee", "professional membership",
        "fee subscription"
    ],
    "Legal and Professional Charges": [
        "legal and professional", "legal charges",
        "professional charges", "legal fee", "consultancy fee",
        "consultancy charges", "audit fee"
    ],
    "Salaries": [
        "salary", "salaries", "payroll", "staff remuneration",
        "employees remuneration", "wages"
    ],
    "Depreciation": [
        "depreciation", "depreciation expense",
        "amortisation", "amortization"
    ],
}


def classify_full_verification(account_name, extra_heads=None):
    """Return the matching full-verification category, if any."""
    account = normalise_text(account_name)

    for category, keywords in FULL_VERIFICATION_KEYWORDS.items():
        for keyword in keywords:
            if normalise_text(keyword) in account:
                return category

    for head in extra_heads or []:
        if normalise_text(head) and normalise_text(head) in account:
            return head

    return None


def build_sampling_data(
    ledger,
    sample_percentage=20,
    include_high_value=True,
    extra_heads=None,
    seed=42,
):
    """Create full-verification and risk-based sample populations."""
    df = ledger.copy()

    df["_Full_Verification_Category"] = df["_Audit_Account"].map(
        lambda account: classify_full_verification(
            account, extra_heads=extra_heads
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

    # Sample within each remaining account head.
    rng = random.Random(seed)

    for _, group in remaining.groupby("_Audit_Account", dropna=False):
        group = group.copy()

        if group.empty:
            continue

        sample_count = max(
            1,
            int(len(group) * sample_percentage / 100),
        )
        sample_count = min(sample_count, len(group))

        if include_high_value:
            group = group.sort_values(
                "_Absolute_Amount", ascending=False
            )
            chosen = list(group.head(min(sample_count, len(group))).index)

            # Randomly fill any remaining sample slots.
            if len(chosen) < sample_count:
                remaining_indices = [
                    idx for idx in group.index if idx not in chosen
                ]
                chosen.extend(
                    rng.sample(
                        remaining_indices,
                        min(sample_count - len(chosen), len(remaining_indices)),
                    )
                )
        else:
            chosen = rng.sample(
                list(group.index), sample_count
            )

        selected_indices.extend(chosen)

    sample_selected = remaining.loc[selected_indices].copy()

    # Keep the remaining, non-selected transactions for population review.
    population_not_selected = remaining.drop(
        index=selected_indices, errors="ignore"
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

    full_summary = (
        full_verification.groupby(
            ["_Audit_Account", "_Full_Verification_Category"],
            dropna=False,
        )
        .agg(
            Items_Selected=("_Audit_Account", "size"),
            Amount_Selected=("_Audit_Amount", "sum"),
        )
        .reset_index()
        .rename(
            columns={
                "_Audit_Account": "Account Head",
                "_Full_Verification_Category": "Verification Category",
            }
        )
    )

    sample_summary = (
        sample_selected.groupby("_Audit_Account", dropna=False)
        .agg(
            Items_Selected=("_Audit_Account", "size"),
            Amount_Selected=("_Audit_Amount", "sum"),
        )
        .reset_index()
        .rename(columns={"_Audit_Account": "Account Head"})
    )

    return (
        full_verification,
        sample_selected,
        population_not_selected,
        summary,
        full_summary,
        sample_summary,
    )


# ============================================================
# EXCEL FORMATTING AND EXPORT
# ============================================================

def dataframe_for_export(df):
    """Remove internal helper columns from user-facing output."""
    result = df.copy()

    helper_columns = [
        col for col in result.columns
        if col.startswith("_Audit_")
    ]
    result = result.drop(columns=helper_columns, errors="ignore")

    return result


def style_worksheet(ws):
    """Apply simple professional formatting."""
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill(
        fill_type="solid", fgColor="17365D"
    )
    header_font = Font(
        bold=True, color="FFFFFF"
    )
    thin_grey = Side(style="thin", color="D9E2F3")

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
        column_letter = get_column_letter(column_cells[0].column)
        max_length = 0

        for cell in column_cells:
            try:
                max_length = max(
                    max_length,
                    len(str(cell.value)) if cell.value is not None else 0,
                )
            except Exception:
                pass

        ws.column_dimensions[column_letter].width = min(
            max(max_length + 2, 12), 38
        )

    ws.row_dimensions[1].height = 30


def build_audit_sampling_workbook(
    ledger_df,
    reference_bytes,
    sample_percentage=20,
    include_high_value=True,
    extra_heads=None,
):
    """Preserve reference workbook sheets and append audit output sheets."""
    if not reference_bytes:
        raise ValueError("Upload the reference sampling workbook.")

    ledger, detected = prepare_detected_ledger(ledger_df)

    full, selected, not_selected, summary, full_summary, sample_summary = (
        build_sampling_data(
            ledger,
            sample_percentage=sample_percentage,
            include_high_value=include_high_value,
            extra_heads=extra_heads,
        )
    )

    workbook = load_workbook(io.BytesIO(reference_bytes))

    # Replace only output sheets generated by this app.
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

        for row in [list(export_df.columns)] + export_df.astype(object).values.tolist():
            cleaned_row = []
            for value in row:
                if pd.isna(value):
                    cleaned_row.append(None)
                elif isinstance(value, (pd.Timestamp, datetime, date)):
                    cleaned_row.append(value.to_pydatetime() if isinstance(value, pd.Timestamp) else value)
                elif hasattr(value, "item"):
                    try:
                        cleaned_row.append(value.item())
                    except Exception:
                        cleaned_row.append(value)
                else:
                    cleaned_row.append(value)
            ws.append(cleaned_row)

        style_worksheet(ws)
        return ws

    # Main summary sheet.
    ws = workbook.create_sheet("Auto Summary")
    ws.append(["AI AUDIT VERIFICATION — SAMPLING SUMMARY"])
    ws.append(["Ledger transactions analysed", len(ledger)])
    ws.append(["Full-verification transactions", len(full)])
    ws.append(["Risk-based sample transactions", len(selected)])
    ws.append(["Remaining unsampled transactions", len(not_selected)])
    ws.append(["Sampling percentage for other heads", sample_percentage])
    ws.append(["Selection method", "Full verification keywords + within-head sample"])
    ws.append([
        "Important",
        "Review the detected fields and full-verification classifications before relying on this workbook."
    ])
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].fill = PatternFill(fill_type="solid", fgColor="17365D")
    ws.merge_cells("A1:B1")
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 90
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

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
            "Review Required": "Yes" if details["ambiguous"] else "No",
        })

    add_dataframe_sheet(
        "Detection Report",
        pd.DataFrame(detection_rows),
    )

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue(), detected, len(full), len(selected)


# ============================================================
# USER INTERFACE
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
        st.error(f"Could not read the General Ledger: {exc}")

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
            "Column detection is automatic. Check ambiguous or missing "
            "fields before using the workbook for audit evidence."
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
            "**Automatic full-verification categories:** Utilities; "
            "Fee and Subscription; Legal and Professional Charges; "
            "Salaries; Depreciation."
        )

        st.caption(
            "These categories are identified using account-name keywords. "
            "Review the classifications because account titles vary between clients."
        )

        if st.button(
            "Generate Audit Sampling Workbook",
            type="primary",
            use_container_width=True,
        ):
            if reference_file is None:
                st.error(
                    "Upload the reference sampling workbook before generating."
                )
            else:
                try:
                    with st.spinner("Preparing the audit sampling workbook..."):
                        excel_bytes, detected, full_count, sample_count = (
                            build_audit_sampling_workbook(
                                ledger_df=ledger_df,
                                reference_bytes=reference_file.getvalue(),
                                sample_percentage=int(sample_percentage),
                                include_high_value=include_high_value,
                            )
                        )

                    st.success(
                        f"Workbook created. Full-verification items: "
                        f"{full_count:,}; risk-based sample items: "
                        f"{sample_count:,}."
                    )

                    st.download_button(
                        label="Download Audit Sampling Workbook",
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
        st.error(f"Automatic ledger analysis failed: {exc}")

else:
    st.info(
        "Upload a General Ledger to start automatic analysis."
    )

st.divider()
st.caption(
    "AI Audit Verification | Generated samples are audit aids, not a "
    "substitute for professional judgement or review of the source records."
)
