import os
import io
import uuid
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from copy import copy

APP_TITLE = "Huliot India | Travel Expense Entry"
TEMPLATE_FILE = "Travelling Expenses Sheet.xlsx"
LOCAL_FILE = "local_entries.csv"
DATA_SHEET_NAME = "Entries"
TEMPLATE_SHEET_NAME = "Travel Reimbur. Form"  # Change if you want to use another sheet from template

COLUMNS = [
    "Timestamp", "EntryID", "Name", "Designation", "Location", "ReportMonth",
    "Date", "NightHalt", "DepartureFrom", "DepartureTo", "DepTime", "ArrivalTime",
    "Mode", "Fare", "LocalConveyance", "DA", "TelInternet", "Courier",
    "StationaryXerox", "MiscDetails", "Remarks"
]

NUMERIC_COLUMNS = ["Fare", "LocalConveyance", "DA", "TelInternet", "Courier", "StationaryXerox"]


def money(value):
    try:
        return float(value or 0)
    except Exception:
        return 0.0


@st.cache_resource
def get_google_worksheet():
    """Return Google Sheet worksheet if Streamlit secrets are configured, otherwise None."""
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
    else:
        if os.path.exists(LOCAL_FILE):
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
    worksheet = get_google_worksheet()
    df = read_entries()
    df = df[df["EntryID"].astype(str) != str(entry_id)].copy()

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


def generate_excel(df: pd.DataFrame, employee: dict) -> bytes:
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb[TEMPLATE_SHEET_NAME] if TEMPLATE_SHEET_NAME in wb.sheetnames else wb.active

    # Header details, kept in the same existing cells/merged areas as far as possible.
    ws["A5"] = f"NAME OF THE MKTG  PERSON : {employee.get('name', '')}"
    ws["E4"] = f"STATE: : {employee.get('state', '')}"
    ws["N5"] = datetime.now().strftime("%d-%m-%Y")
    ws["N6"] = employee.get("hq", "")
    ws["A7"] = f"FROM: {employee.get('from_date', '')}"
    ws["C7"] = f"TO: {employee.get('to_date', '')}"

    data_start_row = 11
    template_last_data_row = 29  # Row above TOTAL/signature area in attached format
    available_rows = template_last_data_row - data_start_row + 1

    df = df.copy()
    if employee.get("name"):
        df = df[df["Name"].astype(str) == employee.get("name")]
    if employee.get("report_month"):
        df = df[df["ReportMonth"].astype(str) == employee.get("report_month")]

    # If entries are more than template rows, insert rows and copy style.
    extra_rows = max(0, len(df) - available_rows)
    if extra_rows:
        ws.insert_rows(template_last_data_row + 1, amount=extra_rows)
        for r in range(template_last_data_row + 1, template_last_data_row + 1 + extra_rows):
            copy_row_style(ws, template_last_data_row, r)

    # Clear old entry area only, not full template.
    max_entry_row = data_start_row + max(len(df), available_rows) + extra_rows - 1
    for r in range(data_start_row, max_entry_row + 1):
        for c in range(1, 15):
            ws.cell(r, c).value = None

    total_row = data_start_row + max(len(df), available_rows) + extra_rows

    for index, (_, row) in enumerate(df.iterrows(), start=data_start_row):
        ws.cell(index, 1).value = row.get("Date", "")
        ws.cell(index, 2).value = row.get("NightHalt", "")
        ws.cell(index, 3).value = row.get("DepartureFrom", "")
        ws.cell(index, 4).value = row.get("DepartureTo", "")
        ws.cell(index, 5).value = row.get("DepTime", "")
        ws.cell(index, 6).value = row.get("ArrivalTime", "")
        ws.cell(index, 7).value = row.get("Mode", "")
        ws.cell(index, 8).value = money(row.get("Fare"))
        ws.cell(index, 9).value = money(row.get("LocalConveyance"))
        ws.cell(index, 10).value = money(row.get("DA"))
        ws.cell(index, 11).value = money(row.get("TelInternet"))
        ws.cell(index, 12).value = money(row.get("Courier"))
        ws.cell(index, 13).value = money(row.get("StationaryXerox"))
        ws.cell(index, 14).value = row.get("MiscDetails", "") or row.get("Remarks", "")

    # Totals row. This keeps the template layout and writes only formulas/values in total area.
    ws.cell(total_row, 1).value = "TOTAL"
    for col in range(8, 14):
        letter = ws.cell(1, col).column_letter
        ws.cell(total_row, col).value = f"=SUM({letter}{data_start_row}:{letter}{total_row-1})"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown("### Huliot India")
    st.title("Travel Expense Entry App")
    st.caption("Entry saves directly to Google Sheets when deployed. No repeated JSON file download required.")

    with st.sidebar:
        st.header("Employee / Report Details")
        name = st.text_input("Name", "Mr. Umesh Nikam")
        designation = st.text_input("Designation", "Executive - Technical Support")
        location = st.text_input("Location / State", "")
        hq = st.text_input("H. QTRS", "")
        report_month = st.text_input("Report Month", datetime.now().strftime("%Y-%m"))
        from_date = st.date_input("From Date", value=date.today())
        to_date = st.date_input("To Date", value=date.today())

        employee = {
            "name": name,
            "designation": designation,
            "state": location,
            "hq": hq,
            "report_month": report_month,
            "from_date": from_date.strftime("%d-%m-%Y"),
            "to_date": to_date.strftime("%d-%m-%Y"),
        }

    st.subheader("Add Entry")
    with st.form("entry_form", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        entry_date = c1.date_input("Date", value=date.today())
        night_halt = c2.text_input("Night Halt")
        departure_from = c3.text_input("Departure From")
        departure_to = c4.text_input("Departure To")

        c5, c6, c7, c8 = st.columns(4)
        dep_time = c5.text_input("Dep Time")
        arrival_time = c6.text_input("Arrival Time")
        mode = c7.selectbox("Mode", ["2 Wheeler", "4 Wheeler", "Rikshaw", "Bus", "Ola/Uber", "Train", "Flight", "Other"])
        fare = c8.number_input("Fare", min_value=0.0, step=1.0)

        c9, c10, c11, c12 = st.columns(4)
        local_conveyance = c9.number_input("Local Conve.", min_value=0.0, step=1.0)
        da = c10.number_input("DA", min_value=0.0, step=1.0)
        tel_internet = c11.number_input("Tel/Fax/Internet", min_value=0.0, step=1.0)
        courier = c12.number_input("Courier", min_value=0.0, step=1.0)

        c13, c14 = st.columns([1, 3])
        stationary = c13.number_input("Stationary/Xerox", min_value=0.0, step=1.0)
        misc = c14.text_input("Misc / Give Details")
        remarks = st.text_area("Remarks")

        submitted = st.form_submit_button("Save Entry")
        if submitted:
            row = {
                "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "EntryID": str(uuid.uuid4()),
                "Name": name,
                "Designation": designation,
                "Location": location,
                "ReportMonth": report_month,
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
            st.success("Entry saved successfully.")

    st.subheader("Saved Entries")
    df = read_entries()
    filtered_df = df.copy()
    if name:
        filtered_df = filtered_df[filtered_df["Name"].astype(str) == name]
    if report_month:
        filtered_df = filtered_df[filtered_df["ReportMonth"].astype(str) == report_month]

    st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    c1, c2, c3 = st.columns(3)
    total_expense = sum(filtered_df[col].apply(money).sum() for col in NUMERIC_COLUMNS if col in filtered_df.columns) if not filtered_df.empty else 0
    c1.metric("Total Entries", len(filtered_df))
    c2.metric("Total Expense", f"₹{total_expense:,.2f}")

    excel_bytes = generate_excel(df, employee)
    c3.download_button(
        "Download Same Excel Format",
        data=excel_bytes,
        file_name=f"Travelling_Expenses_{name.replace(' ', '_')}_{report_month}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Delete one wrong entry"):
        delete_id = st.text_input("Paste EntryID to delete")
        if st.button("Delete Entry") and delete_id:
            delete_entry(delete_id)
            st.warning("Entry deleted. Refresh page if table is not updated.")

    with st.expander("Storage Status"):
        if get_google_worksheet() is not None:
            st.success("Google Sheets connected. Entries are saved live.")
        else:
            st.info("Google Sheets secrets not found. App is using local CSV fallback. For Streamlit Cloud, configure secrets for permanent live saving.")


if __name__ == "__main__":
    main()
