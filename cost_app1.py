import io
import os
import re
import zipfile
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st
from dateutil.parser import parse
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from db_utils import (
    fetch_case_data,
    fetch_client_list,
    fetch_flexi_fields,
    get_engine,
    test_connection,
)

# ======================================================
# PAGE CONFIGURATION
# ======================================================

st.set_page_config(
    page_title="Cost Approval Tracker",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header {
        background-color: #2E4057;
        padding: 1rem;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #D4EDDA;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #28A745;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #D1ECF1;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #17A2B8;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #FFF3CD;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #FFC107;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><h1 style="color: white; margin: 0;">📊 Cost Approval Tracker</h1></div>', unsafe_allow_html=True)

DEFAULT_SAVE_FOLDER = r"C:\Cost_Approval"

HOLIDAYS = [
    "2025-01-26", "2025-08-15", "2025-10-02", "2025-12-25"
]

# ======================================================
# EXCEL HELPERS
# ======================================================

def clean_filename(name):
    invalid_chars = r'\/:*?"<>|'
    name = str(name)
    for ch in invalid_chars:
        name = name.replace(ch, "_")
    return name.strip()


BRAND_COLOR = "2E4057"
BRAND_ACCENT = "4472C4"
ZEBRA_FILL_COLOR = "F2F6FC"
TITLE_FONT_COLOR = "2E4057"
SUBTITLE_FONT_COLOR = "808080"

THIN_GRAY = Side(style="thin", color="D9D9D9")
CELL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)

# Ageing-bucket colour scale: cool green (fresh) -> warm red (overdue)
BUCKET_STYLES = {
    "0-5":     ("C6E7B0", "1E5C1E"),
    "06-10":   ("E2F0CB", "3B6E22"),
    "11-15":   ("FFF3B0", "7A5B00"),
    "16-20":   ("FFD98E", "8A5000"),
    "21-25":   ("FFB380", "8A3800"),
    "26-30":   ("FF9B8A", "8A1F00"),
    "30+":     ("FF7A7A", "6E0000"),
    "Unknown": ("E0E0E0", "595959"),
}


def _style_header_row(ws, row_idx, first_col, last_col):
    header_fill = PatternFill(start_color=BRAND_COLOR, end_color=BRAND_COLOR, fill_type="solid")
    header_font = Font(bold=True, size=11, color="FFFFFF")
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = CELL_BORDER


def _autofit_columns(ws, header_row, first_col, last_col):
    for col in range(first_col, last_col + 1):
        column_letter = get_column_letter(col)
        max_length = len(str(ws.cell(row=header_row, column=col).value or ""))
        for row in range(header_row + 1, min(ws.max_row + 1, header_row + 200)):
            cell_value = ws.cell(row=row, column=col).value
            if cell_value:
                max_length = max(max_length, len(str(cell_value)))
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 45)


def beautify_excel(file_path, title=None, subtitle=None):
    """Apply a clean, professional look to every sheet: title band, coloured header,
    zebra striping, borders, autofit columns, ageing-bucket colour scale and bold totals."""
    try:
        wb = load_workbook(file_path)

        for ws in wb.worksheets:
            last_col = ws.max_column
            header_row = 1

            if title:
                ws.insert_rows(1, amount=3)
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
                title_cell = ws.cell(row=1, column=1, value=title)
                title_cell.font = Font(bold=True, size=14, color=TITLE_FONT_COLOR)
                title_cell.alignment = Alignment(horizontal='left', vertical='center')

                if subtitle:
                    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
                    sub_cell = ws.cell(row=2, column=1, value=subtitle)
                    sub_cell.font = Font(italic=True, size=10, color=SUBTITLE_FONT_COLOR)
                    sub_cell.alignment = Alignment(horizontal='left', vertical='center')

                ws.row_dimensions[1].height = 22
                header_row = 4

            _style_header_row(ws, header_row, 1, last_col)
            ws.freeze_panes = f"A{header_row + 1}"
            ws.auto_filter.ref = f"{get_column_letter(1)}{header_row}:{get_column_letter(last_col)}{ws.max_row}"

            # Locate special columns for this sheet
            headers = {ws.cell(row=header_row, column=c).value: c for c in range(1, last_col + 1)}
            bucket_col = headers.get("Bucket")
            cost_col = headers.get("approval_cost") or headers.get("Cost")
            grand_total_col = headers.get("Grand Total")

            for row in range(header_row + 1, ws.max_row + 1):
                is_total_row = str(ws.cell(row=row, column=1).value).strip() == "Grand Total"
                zebra = (row - header_row) % 2 == 0

                for col in range(1, last_col + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = CELL_BORDER
                    if col != 1:
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    if zebra and not is_total_row:
                        cell.fill = PatternFill(start_color=ZEBRA_FILL_COLOR, end_color=ZEBRA_FILL_COLOR, fill_type="solid")
                    if is_total_row:
                        cell.font = Font(bold=True)
                        cell.fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")

                if cost_col:
                    ws.cell(row=row, column=cost_col).number_format = '#,##0.00'

                if bucket_col and not is_total_row:
                    bucket_val = ws.cell(row=row, column=bucket_col).value
                    fill_color, font_color = BUCKET_STYLES.get(bucket_val, (None, None))
                    if fill_color:
                        bcell = ws.cell(row=row, column=bucket_col)
                        bcell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
                        bcell.font = Font(bold=True, color=font_color)

                if grand_total_col:
                    gcell = ws.cell(row=row, column=grand_total_col)
                    gcell.font = Font(bold=True)

            _autofit_columns(ws, header_row, 1, last_col)

        wb.save(file_path)
        return True
    except Exception as e:
        st.warning(f"Could not beautify Excel file: {str(e)}")
        return False


def create_pivot_table(df):
    """Bucket x Check_unique_name pivot for on-screen summary"""
    required_columns = ['Check_unique_name', 'Bucket', 'Case_Check_id']
    missing_cols = [col for col in required_columns if col not in df.columns]
    if missing_cols:
        return None

    pivot = pd.pivot_table(
        df,
        values='Case_Check_id',
        index='Check_unique_name',
        columns='Bucket',
        aggfunc='count',
        fill_value=0,
        margins=True,
        margins_name='Grand Total'
    )

    bucket_order = ['0-5', '06-10', '11-15', '16-20', '21-25', '26-30', '30+', 'Unknown']
    existing_buckets = [b for b in bucket_order if b in pivot.columns]
    other_columns = [c for c in pivot.columns if c not in existing_buckets and c != 'Grand Total']
    pivot = pivot[existing_buckets + other_columns + ['Grand Total']]

    pivot = pivot.reset_index()
    pivot.columns.name = None
    return pivot


# ======================================================
# COST EXTRACTION LOGIC
# ======================================================

PATTERNS = [
    re.compile(r'(?:rs|inr|usd|cad|gbp|eur|aed|₹|\$|£)\s?(\d+(?:\.\d+)?)', re.IGNORECASE),
    re.compile(r'(\d+(?:\.\d+)?)\s?(?:inr|usd|gbp|eur|aed|cad|aud|chf)', re.IGNORECASE),
    re.compile(r'(?:rupees)\s?(\d+(?:\.\d+)?)', re.IGNORECASE),
    re.compile(r'(\d+(?:\.\d+)?)\s?(?:pound|pounds|gbp)', re.IGNORECASE),
    re.compile(r'(\d+(?:\.\d+)?)\s?(?:usd|dollar|dollars)', re.IGNORECASE),
    re.compile(r'(?:amount|cost|price|charges|fees|payment|approval|salary|ctc)\D{0,20}(\d+(?:\.\d+)?)', re.IGNORECASE),
    re.compile(r'(\d+(?:\.\d+)?)\s?/-', re.IGNORECASE),
    re.compile(r'(\d+(?:\.\d+)?)\s?(?:\+gst|including gst|incl gst)', re.IGNORECASE),
]

APPROVAL_PATTERN = re.compile(r'(?:require|required)\s+(?:cost|additional cost)\s+approval\s+of\s+(?:usd|gbp|inr|eur|cad)?\s?(\d+(?:\.\d+)?)', re.IGNORECASE)
TOTAL_PATTERN = re.compile(r'total\s+verification\s+cost\s*:-?\s*(\d+(?:\.\d+)?)', re.IGNORECASE)


def normalize_text_fast(text):
    if pd.isna(text) or text == "" or text == "nan":
        return ""
    text = str(text).lower()
    replacements = [(",", " "), ("\n", " "), ("rs.", "rs "), ("inr.", "inr "),
                    ("$", "usd "), ("£", "gbp "), ("€", "eur "), ("₹", "inr ")]
    for old, new in replacements:
        text = text.replace(old, new)
    return text


def extract_cost_fast(text):
    if not text or text == "":
        return None
    text = normalize_text_fast(text)
    extracted_values = []
    for pattern in PATTERNS:
        for match in pattern.finditer(text):
            try:
                value = float(match.group(1))
                if value <= 0 or value > 999999999:
                    continue
                extracted_values.append(value)
            except Exception:
                continue
    if not extracted_values:
        return None
    extracted_values = list(set(v for v in extracted_values if v > 0))
    if not extracted_values:
        return None
    extracted_values.sort()
    if len(extracted_values) > 1:
        approval_match = APPROVAL_PATTERN.search(text)
        if approval_match:
            try:
                preferred = float(approval_match.group(1))
                if preferred in extracted_values:
                    return preferred
            except Exception:
                pass
        total_match = TOTAL_PATTERN.search(text)
        if total_match:
            try:
                preferred = float(total_match.group(1))
                if preferred in extracted_values:
                    return preferred
            except Exception:
                pass
    return extracted_values[0]


def calculate_networkdays(start_date, end_date, holidays_list):
    if start_date > end_date:
        return 0
    business_days = 0
    current_date = start_date
    holidays = []
    for h in holidays_list:
        try:
            holidays.append(parse(h).date())
        except Exception:
            pass
    while current_date <= end_date:
        if current_date.weekday() < 5 and current_date not in holidays:
            business_days += 1
        current_date += timedelta(days=1)
    return business_days


def get_bucket_from_ageing(ageing_days):
    if pd.isna(ageing_days):
        return "Unknown"
    if ageing_days <= 5:
        return "0-5"
    elif ageing_days <= 10:
        return "06-10"
    elif ageing_days <= 15:
        return "11-15"
    elif ageing_days <= 20:
        return "16-20"
    elif ageing_days <= 25:
        return "21-25"
    elif ageing_days <= 30:
        return "26-30"
    else:
        return "30+"


def process_case_data(df, holidays_list):
    """Extract approval cost from insuff_remarks, compute ageing/bucket, keep only rows with a detected cost."""
    df = df.copy()
    df['approval_cost'] = df['insuff_remarks'].apply(extract_cost_fast)
    df_with_cost = df[df['approval_cost'].notna()].copy()
    if df_with_cost.empty:
        return df_with_cost

    today = date.today()
    ageing_list = []
    for _, row in df_with_cost.iterrows():
        raised = row.get('Insuff Raised Date')
        try:
            if pd.notna(raised):
                if isinstance(raised, str):
                    raised_date = parse(raised).date()
                elif isinstance(raised, (datetime, date)):
                    raised_date = raised if isinstance(raised, date) and not isinstance(raised, datetime) else raised.date()
                else:
                    ageing_list.append(None)
                    continue
                ageing_list.append(calculate_networkdays(raised_date, today, holidays_list))
            else:
                ageing_list.append(None)
        except Exception:
            ageing_list.append(None)

    df_with_cost['Ageing'] = ageing_list
    df_with_cost['Bucket'] = df_with_cost['Ageing'].apply(get_bucket_from_ageing)
    return df_with_cost


# ======================================================
# CLIENT TRACKER SPLITTER
# ======================================================

def build_client_trackers(case_df, flex_df):
    """Build one formatted Excel file per client_id, renaming CASE_FLEX_FIELDx columns
    to that client's actual field names using the flexi-field mapping."""
    files = {}
    for client_id, client_df in case_df.groupby('client_id'):
        client_df = client_df.copy()
        client_flex = flex_df[flex_df['client_id'] == client_id]

        rename_dict = {}
        valid_case_columns = []
        for _, row in client_flex.iterrows():
            try:
                field_id = int(float(row['field_id']))
                field_name = str(row['field_name']).strip()
                if not field_name or field_name.lower() == "nan":
                    continue
                source_column = f"CASE_FLEX_FIELD{field_id}"
                if source_column in client_df.columns:
                    rename_dict[source_column] = field_name
                    valid_case_columns.append(source_column)
            except Exception:
                continue

        client_df.rename(columns=rename_dict, inplace=True)
        cols_to_drop = [c for c in client_df.columns if c.startswith("CASE_FLEX_FIELD") and c not in valid_case_columns]
        client_df.drop(columns=cols_to_drop, inplace=True, errors="ignore")

        company_name = client_df['Company_name'].iloc[0] if not client_df.empty else str(client_id)
        file_name = f"{clean_filename(company_name)}.xlsx"

        bucket_summary = create_pivot_table(client_df)

        title = f"Cost Approval Tracker — {company_name}"
        subtitle = f"Generated on {datetime.now().strftime('%d-%b-%Y %H:%M')}  |  Total Cases: {len(client_df)}"

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            client_df.to_excel(writer, sheet_name="Data", index=False)
            if bucket_summary is not None:
                bucket_summary.to_excel(writer, sheet_name="Bucket Summary", index=False)
        buffer.seek(0)

        # beautify_excel needs a file path, so round-trip through a temp path in memory is not possible with openpyxl directly on BytesIO paths
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(buffer.getvalue())
            tmp_path = tmp.name
        beautify_excel(tmp_path, title=title, subtitle=subtitle)
        with open(tmp_path, "rb") as f:
            data = f.read()
        os.remove(tmp_path)

        files[file_name] = data

    return files


# ======================================================
# STREAMLIT APP
# ======================================================

st.markdown("""
<div class="info-box">
    <p>Pulls the <strong>Cost Approval</strong> report directly from the database (no more manual Query Browser export),
    extracts the approval cost from insufficiency remarks, computes ageing/bucket, and splits the result into one
    formatted Excel tracker per client.</p>
</div>
""", unsafe_allow_html=True)

if "engine" not in st.session_state:
    try:
        st.session_state["engine"] = get_engine()
    except Exception as e:
        st.session_state["engine"] = None
        st.error(f"❌ Could not build DB engine: {e}")

engine = st.session_state.get("engine")

with st.sidebar:
    st.subheader("🔌 Database")
    if engine is not None and st.button("Test Connection"):
        ok, msg = test_connection(engine)
        if ok:
            st.success(msg)
        else:
            st.error(msg)

if engine is None:
    st.stop()

# ------------------------------------------------------
# Step 1: Load client list
# ------------------------------------------------------
st.subheader("1️⃣ Select Client(s)")

if st.button("🔄 Load Client List"):
    with st.spinner("Fetching clients with open cost-approval insufficiencies..."):
        try:
            st.session_state["client_list_df"] = fetch_client_list(engine)
        except Exception as e:
            st.error(f"❌ Failed to fetch client list: {e}")

client_list_df = st.session_state.get("client_list_df")

selected_client_ids = None  # None => all clients

if client_list_df is not None and not client_list_df.empty:
    all_clients = st.checkbox("All Clients", value=True)
    if not all_clients:
        options = {
            f"{row['Company_name']} ({row['client_id']})": int(row['client_id'])
            for _, row in client_list_df.iterrows()
        }
        picked = st.multiselect("Client(s)", options=list(options.keys()))
        selected_client_ids = [options[p] for p in picked] if picked else []
        if not selected_client_ids:
            st.info("No client selected — pick at least one, or check 'All Clients'.")
else:
    st.info("Click 'Load Client List' to populate the client selector (optional — you can also generate for all clients directly).")

# ------------------------------------------------------
# Step 2: Generate tracker
# ------------------------------------------------------
st.subheader("2️⃣ Generate Tracker")

save_locally = st.checkbox("Also save a copy to a local folder", value=False)
save_folder = DEFAULT_SAVE_FOLDER
if save_locally:
    save_folder = st.text_input("Save folder", value=DEFAULT_SAVE_FOLDER)

can_generate = client_list_df is None or selected_client_ids is None or len(selected_client_ids) > 0

if st.button("🚀 Generate Cost Approval Tracker", disabled=not can_generate, use_container_width=True):
    try:
        with st.spinner("Fetching case data from database..."):
            case_df = fetch_case_data(engine, selected_client_ids)
        with st.spinner("Fetching client flexi-field mapping..."):
            flex_df = fetch_flexi_fields(engine, selected_client_ids)

        if case_df.empty:
            st.warning("⚠️ No cases found for the selected client(s).")
        else:
            with st.spinner("Extracting approval cost and computing ageing..."):
                processed_df = process_case_data(case_df, HOLIDAYS)

            if processed_df.empty:
                st.warning("⚠️ No rows with a detectable approval cost were found for the selected client(s).")
            else:
                n_clients = processed_df['client_id'].nunique()
                col1, col2, col3 = st.columns(3)
                col1.metric("📄 Cases Pulled", len(case_df))
                col2.metric("💰 Cases With Cost Detected", len(processed_df))
                col3.metric("🏢 Clients", n_clients)

                pivot = create_pivot_table(processed_df)
                if pivot is not None:
                    st.subheader("📊 Bucket-wise Summary")
                    st.dataframe(pivot, use_container_width=True, hide_index=True)

                with st.spinner("Building per-client Excel trackers..."):
                    files = build_client_trackers(processed_df, flex_df)

                if save_locally:
                    os.makedirs(save_folder, exist_ok=True)
                    for fname, data in files.items():
                        with open(os.path.join(save_folder, fname), "wb") as f:
                            f.write(data)
                    st.success(f"✅ Saved {len(files)} tracker(s) to {save_folder}")

                st.markdown(f"""
                <div class="success-box">
                    <h3>✅ Success!</h3>
                    <p>Generated <strong>{len(files)}</strong> client tracker(s)</p>
                </div>
                """, unsafe_allow_html=True)

                if len(files) == 1:
                    fname, data = next(iter(files.items()))
                    st.download_button(
                        label=f"📥 Download {fname}",
                        data=data,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                        for fname, data in files.items():
                            zf.writestr(fname, data)
                    zip_buffer.seek(0)
                    st.download_button(
                        label="📥 Download All Trackers (ZIP)",
                        data=zip_buffer,
                        file_name=f"Cost_Approval_Output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
        st.exception(e)

st.markdown("---")
st.markdown("<p style='text-align: center; color: #666;'>© 2026 Cost Approval Tracker</p>", unsafe_allow_html=True)
