import os
import io
import uuid
from copy import copy
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

APP_TITLE = "Huliot India | Travel Expense Entry"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, "Travelling Expenses Sheet.xlsx")
LOCAL_FILE = os.path.join(BASE_DIR, "local_entries.csv")
DATA_SHEET_NAME = "Entries"
TEMPLATE_SHEET_NAME = "Travel Reimbur. Form"

COLUMNS = [
    "Timestamp", "EntryID", "Name", "Designation", "Location", "ReportMonth", "HQ",
    "FromDate", "ToDate", "Date", "NightHalt", "DepartureFrom", "DepartureTo", "DepTime", "ArrivalTime",
    "Mode", "Fare", "LocalConveyance", "DA", "TelInternet", "Courier",
    "StationaryXerox", "MiscDetails", "Remarks"
]
NUMERIC_COLUMNS = ["Fare", "LocalConveyance", "DA", "TelInternet", "Courier", "StationaryXerox"]

HTML_STYLE = """
<style>
    .stApp { background: #f4f7fb; }
    .main .block-container { padding-top: 1rem; max-width: 1220px; }
    .brand-header {
        background: linear-gradient(135deg, #0f2f73, #2563eb);
        color: white;
        padding: 20px 24px;
        border-radius: 18px;
        box-shadow: 0 8px 24px rgba(37,99,235,.20);
        margin-bottom: 16px;
    }
    .brand-header h1 { margin: 0; font-size: 30px; font-weight: 800; }
    .brand-header p { margin: 6px 0 0; opacity: .92; font-size: 15px; }
    .html-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 16px 18px 18px;
        box-shadow: 0 5px 18px rgba(15, 23, 42, .07);
        margin-bottom: 16px;
    }
    .section-title {
        color: #111827;
        font-weight: 800;
        font-size: 20px;
        margin: 2px 0 12px;
        padding-bottom: 8px;
        border-bottom: 1px solid #e5e7eb;
    }
    .hint-box {
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        color: #1e3a8a;
        padding: 10px 12px;
        border-radius: 12px;
        font-weight: 700;
        margin-bottom: 12px;
    }
    .status-ok {
        background: #ecfdf5;
        border: 1px solid #a7f3d0;
        color: #065f46;
        padding: 10px 12px;
        border-radius: 12px;
        font-weight: 700;
    }
    .status-warn {
        background: #fff7ed;
        border: 1px solid #fed7aa;
        color: #9a3412;
        padding: 10px 12px;
        border-radius: 12px;
        font-weight: 700;
    }
    div[data-testid="stMetric"] {
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        border-radius: 14px;
        padding: 12px;
    }
    .stButton>button, .stDownloadButton>button {
        border-radius: 10px;
        font-weight: 800;
        padding: .62rem 1rem;
    }
    .stDownloadButton>button { background: #16a34a; color: white; border: 0; }
    div[data-testid="stForm"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 5px 18px rgba(15, 23, 42, .07);
    }
    label { font-weight: 700 !important; color: #374151 !important; }
    [data-testid="stDataFrame"] { border-radius: 14px; overflow: hidden; }
    .small-text { color:#64748b; font-size:13px; }
</style>
"""


def money(value):
    try:
        return float(value or 0)
    except Exception:
        return 0.0


@st.cache_resource(show_spinner=False)
def get_google_worksheet():
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        service_account_info = dict(st.secrets["gcp_service_account"])
        credentials = Credentials.from_service_account_info(service_account_info, scopes=scopes)
        client = gspread.authorize(credentials)
        spreadsheet = client.open_by_key(st.secrets["gsheets"]["spreadsheet_id"])

        try:
            worksheet = spreadsheet.worksheet(DATA_SHEET_NAME)
        except Exception:
            worksheet = spreadsheet.add_worksheet(title=DATA_SHEET_NAME, rows=1000, cols=len(COLUMNS) + 5)
            worksheet.append_row(COLUMNS)
            return worksheet

        existing = worksheet.get_all_values()
        if not existing:
            worksheet.append_row(COLUMNS)
        elif existing[0] != COLUMNS:
            worksheet.update("A1", [COLUMNS])
        return worksheet
    except Exception:
        return None


def read_entries() -> pd.DataFrame:
    worksheet = get_google_worksheet()
    if worksheet is not None:
        records = worksheet.get_all_records()
        df = pd.DataFrame(records)
    elif os.path.exists(LOCAL_FILE):
        df = pd.read_csv(LOCAL_FILE)
    else:
        df = pd.DataFrame(columns=COLUMNS)

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[COLUMNS]


def save_entry(row: dict):
    row = {col: row.get(col, "") for col in COLUMNS}
    worksheet = get_google_worksheet()
    if worksheet is not None:
        worksheet.append_row([row[col] for col in COLUMNS], value_input_option="USER_ENTERED")
    else:
        df = read_entries()
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        df.to_csv(LOCAL_FILE, index=False)


def delete_entry(entry_id: str):
    df = read_entries()
    df = df[df["EntryID"].astype(str) != str(entry_id)].copy()
    worksheet = get_google_worksheet()
    if worksheet is not None:
        worksheet.clear()
        worksheet.update("A1", [COLUMNS] + df.fillna("").astype(str).values.tolist())
    else:
        df.to_csv(LOCAL_FILE, index=False)


def copy_row_style(ws, source_row: int, target_row: int):
    for col in range(1, ws.max_column + 1):
        source_cell = ws.cell(source_row, col)
        target_cell = ws.cell(target_row, col)
        if source_cell.has_style:
            target_cell.font = copy(source_cell.font)
            target_cell.border = copy(source_cell.border)
            target_cell.fill = copy(source_cell.fill)
            target_cell.number_format = source_cell.number_format
            target_cell.protection = copy(source_cell.protection)
            target_cell.alignment = copy(source_cell.alignment)
        target_cell._style = copy(source_cell._style)
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def safe_set_cell_value(ws, cell_ref: str, value):
    cell = ws[cell_ref]
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            ws.cell(merged_range.min_row, merged_range.min_col).value = value
            return
    cell.value = value


def safe_clear_cell(ws, row: int, col: int):
    cell = ws.cell(row, col)
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            if row == merged_range.min_row and col == merged_range.min_col:
                cell.value = None
            return
    cell.value = None


def generate_excel(df: pd.DataFrame, employee: dict) -> bytes:
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb[TEMPLATE_SHEET_NAME] if TEMPLATE_SHEET_NAME in wb.sheetnames else wb.active

    safe_set_cell_value(ws, "A5", f"NAME OF THE MKTG  PERSON : {employee.get('name', '')}")
    safe_set_cell_value(ws, "E4", f"STATE: : {employee.get('state', '')}")
    safe_set_cell_value(ws, "N5", datetime.now().strftime("%d-%m-%Y"))
    safe_set_cell_value(ws, "N6", employee.get("hq", ""))
    safe_set_cell_value(ws, "A7", f"FROM: {employee.get('from_date', '')}")
    safe_set_cell_value(ws, "C7", f"TO: {employee.get('to_date', '')}")

    data_start_row = 11
    template_last_data_row = 29
    available_rows = template_last_data_row - data_start_row + 1

    df = df.copy()
    if employee.get("name"):
        df = df[df["Name"].astype(str) == employee.get("name")]
    if employee.get("report_month"):
        df = df[df["ReportMonth"].astype(str) == employee.get("report_month")]

    extra_rows = max(0, len(df) - available_rows)
    if extra_rows:
        ws.insert_rows(template_last_data_row + 1, amount=extra_rows)
        for r in range(template_last_data_row + 1, template_last_data_row + 1 + extra_rows):
            copy_row_style(ws, template_last_data_row, r)

    max_entry_row = data_start_row + max(len(df), available_rows) - 1
    for r in range(data_start_row, max_entry_row + 1):
        for c in range(1, 15):
            safe_clear_cell(ws, r, c)

    total_row = data_start_row + max(len(df), available_rows)

    for excel_row, (_, row) in enumerate(df.iterrows(), start=data_start_row):
        ws.cell(excel_row, 1).value = row.get("Date", "")
        ws.cell(excel_row, 2).value = row.get("NightHalt", "")
        ws.cell(excel_row, 3).value = row.get("DepartureFrom", "")
        ws.cell(excel_row, 4).value = row.get("DepartureTo", "")
        ws.cell(excel_row, 5).value = row.get("DepTime", "")
        ws.cell(excel_row, 6).value = row.get("ArrivalTime", "")
        ws.cell(excel_row, 7).value = row.get("Mode", "")
        ws.cell(excel_row, 8).value = money(row.get("Fare"))
        ws.cell(excel_row, 9).value = money(row.get("LocalConveyance"))
        ws.cell(excel_row, 10).value = money(row.get("DA"))
        ws.cell(excel_row, 11).value = money(row.get("TelInternet"))
        ws.cell(excel_row, 12).value = money(row.get("Courier"))
        ws.cell(excel_row, 13).value = money(row.get("StationaryXerox"))
        ws.cell(excel_row, 14).value = row.get("MiscDetails", "") or row.get("Remarks", "")

    ws.cell(total_row, 1).value = "TOTAL"
    for col in range(8, 14):
        letter = get_column_letter(col)
        ws.cell(total_row, col).value = f"=SUM({letter}{data_start_row}:{letter}{total_row-1})"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(HTML_STYLE, unsafe_allow_html=True)
    st.markdown(
        '<div class="brand-header"><h1>Huliot India</h1><p>Travel Expense Entry App, easy HTML-style form with live Google Sheet saving.</p></div>',
        unsafe_allow_html=True,
    )

    if "employee" not in st.session_state:
        st.session_state.employee = {
            "name": "Mr. Umesh Nikam",
            "designation": "Executive - Technical Support",
            "state": "",
            "hq": "",
            "report_month": datetime.now().strftime("%Y-%m"),
            "from_date": date.today(),
            "to_date": date.today(),
        }

    st.markdown('<div class="html-card"><div class="section-title">Employee / Report Details</div><div class="hint-box">Fill this once, then add daily travel entries below.</div></div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    name = c1.text_input("Name", value=st.session_state.employee["name"])
    designation = c2.text_input("Designation", value=st.session_state.employee["designation"])
    location = c3.text_input("State / Location", value=st.session_state.employee["state"])
    hq = c4.text_input("H. QTRS", value=st.session_state.employee["hq"])

    c5, c6, c7 = st.columns(3)
    report_month = c5.text_input("Report Month", value=st.session_state.employee["report_month"], help="Example: 2026-05")
    from_date = c6.date_input("From Date", value=st.session_state.employee["from_date"])
    to_date = c7.date_input("To Date", value=st.session_state.employee["to_date"])

    employee = {
        "name": name,
        "designation": designation,
        "state": location,
        "hq": hq,
        "report_month": report_month,
        "from_date": from_date.strftime("%d-%m-%Y"),
        "to_date": to_date.strftime("%d-%m-%Y"),
    }
    st.session_state.employee.update({**employee, "from_date": from_date, "to_date": to_date})

    with st.form("entry_form", clear_on_submit=True):
        st.markdown('<div class="section-title">Add Travel Expense</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        entry_date = c1.date_input("Date", value=date.today())
        night_halt = c2.text_input("Night Halt", placeholder="Yes / No / City")
        departure_from = c3.text_input("Departure From", placeholder="Start location")
        departure_to = c4.text_input("Departure To", placeholder="End location")

        c5, c6, c7, c8 = st.columns(4)
        dep_time = c5.text_input("Dep Time", placeholder="09:30 AM")
        arrival_time = c6.text_input("Arrival Time", placeholder="06:30 PM")
        mode = c7.selectbox("Mode", ["2 Wheeler", "4 Wheeler", "Rikshaw", "Bus", "Ola/Uber", "Train", "Flight", "Other"])
        fare = c8.number_input("Fare", min_value=0.0, step=1.0)

        c9, c10, c11, c12 = st.columns(4)
        local_conveyance = c9.number_input("Local Conve.", min_value=0.0, step=1.0)
        da = c10.number_input("DA", min_value=0.0, step=1.0)
        tel_internet = c11.number_input("Tel/Fax/Internet", min_value=0.0, step=1.0)
        courier = c12.number_input("Courier", min_value=0.0, step=1.0)

        c13, c14 = st.columns([1, 3])
        stationary = c13.number_input("Stationary/Xerox", min_value=0.0, step=1.0)
        misc = c14.text_input("Misc / Give Details", placeholder="Purpose / company / site / extra details")
        remarks = st.text_area("Remarks", placeholder="Meeting details or important note")

        submitted = st.form_submit_button("Save Entry", use_container_width=True)
        if submitted:
            if not departure_from or not departure_to:
                st.error("Please fill Departure From and Departure To.")
            else:
                row = {
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "EntryID": str(uuid.uuid4()),
                    "Name": name,
                    "Designation": designation,
                    "Location": location,
                    "ReportMonth": report_month,
                    "HQ": hq,
                    "FromDate": from_date.strftime("%d-%m-%Y"),
                    "ToDate": to_date.strftime("%d-%m-%Y"),
                    "Date": entry_date.strftime("%d-%m-%Y"),
                    "NightHalt": night_halt,
                    "DepartureFrom": departure_from,
                    "DepartureTo": departure_to,
                    "DepTime": dep_time,
                    "ArrivalTime": arrival_time,
                    "Mode": mode,
                    "Fare": fare,
                    "LocalConveyance": local_conveyance,
                    "DA": da,
                    "TelInternet": tel_internet,
                    "Courier": courier,
                    "StationaryXerox": stationary,
                    "MiscDetails": misc,
                    "Remarks": remarks,
                }
                save_entry(row)
                st.success("Entry saved successfully in Google Sheet.")

    df = read_entries()
    filtered_df = df.copy()
    if name:
        filtered_df = filtered_df[filtered_df["Name"].astype(str) == name]
    if report_month:
        filtered_df = filtered_df[filtered_df["ReportMonth"].astype(str) == report_month]

    total_expense = sum(filtered_df[col].apply(money).sum() for col in NUMERIC_COLUMNS if col in filtered_df.columns) if not filtered_df.empty else 0

    st.markdown('<div class="html-card"><div class="section-title">Expense Summary</div></div>', unsafe_allow_html=True)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Entries", len(filtered_df))
    m2.metric("Total Expense", f"₹{total_expense:,.2f}")
    m3.metric("Report Month", report_month)
    m4.metric("Storage", "Google Sheet" if get_google_worksheet() is not None else "Local CSV")

    excel_bytes = generate_excel(df, employee)
    st.download_button(
        "Download Same Excel Format",
        data=excel_bytes,
        file_name=f"Travelling_Expenses_{name.replace(' ', '_')}_{report_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown('<div class="html-card"><div class="section-title">Saved Entries</div></div>', unsafe_allow_html=True)
    show_cols = ["Date", "NightHalt", "DepartureFrom", "DepartureTo", "DepTime", "ArrivalTime", "Mode", "Fare", "LocalConveyance", "DA", "TelInternet", "Courier", "StationaryXerox", "MiscDetails", "Remarks", "EntryID"]
    show_cols = [c for c in show_cols if c in filtered_df.columns]
    st.dataframe(filtered_df[show_cols], use_container_width=True, hide_index=True)

    with st.expander("Delete wrong entry"):
        st.caption("Copy EntryID from the Saved Entries table and paste here.")
        delete_id = st.text_input("EntryID")
        if st.button("Delete Entry") and delete_id:
            delete_entry(delete_id)
            st.warning("Entry deleted. Refresh the app once if table does not update immediately.")

    with st.sidebar:
        st.markdown("### App Status")
        if get_google_worksheet() is not None:
            st.markdown('<div class="status-ok">Google Sheets connected. Entries are saved live.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-warn">Google Sheets not connected. Local CSV fallback is active.</div>', unsafe_allow_html=True)
        st.markdown("### Daily Use")
        st.write("1. Fill Employee details")
        st.write("2. Add travel entry")
        st.write("3. Save Entry")
        st.write("4. Download Excel when required")


if __name__ == "__main__":
    main()
