import os
import io
import uuid
import traceback
from copy import copy
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

APP_TITLE = "Huliot Travel Expense Entry"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, "Travelling Expenses Sheet.xlsx")
TEMPLATE_SHEET_NAME = "Travel Reimbur. Form"
WORKSHEET_NAME_DEFAULT = "Entries"

VEHICLE_OPTIONS = ["", "2 Wheeler", "4 Wheeler", "Public Transport", "Train", "Bus", "Flight", "Other"]
HEADERS = ["id", "empName", "designation", "location", "reportMonthText", "date", "from", "to", "vehicle", "company", "contact", "invoice", "toll", "fuel", "lodging", "food", "tel", "courier", "rikshaw", "remarks", "total", "created_at"]
AMOUNT_COLUMNS = ["toll", "fuel", "lodging", "food", "tel", "courier", "rikshaw"]

def make_str(value):
    if value is None:
        return ""
    return str(value).strip()

def money(value):
    try:
        if value in (None, ""):
            return 0.0
        return float(value)
    except Exception:
        return 0.0

def current_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def calc_total(row):
    return sum(money(row.get(col, 0)) for col in AMOUNT_COLUMNS)

def default_employee():
    return {
        "empName": "Mr.Umesh Nikam",
        "designation": "Executive-Technical Support",
        "location": "Pune",
        "reportMonthText": "01st Jun 2026 To 30th Jun 2026",
    }

def get_secret_value(name, default_value=""):
    try:
        return st.secrets.get(name, default_value)
    except Exception:
        return default_value

@st.cache_resource(show_spinner=False)
def get_google_worksheet():
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        if "gcp_service_account" not in st.secrets:
            raise RuntimeError("Missing [gcp_service_account] in Streamlit Secrets.")

        spreadsheet_id = ""
        if "gsheets" in st.secrets:
            spreadsheet_id = st.secrets["gsheets"].get("spreadsheet_id", "")

        if not spreadsheet_id:
            raise RuntimeError("Missing [gsheets] spreadsheet_id in Streamlit Secrets.")

        worksheet_name = get_secret_value("WORKSHEET_NAME", WORKSHEET_NAME_DEFAULT)
        worksheet_name = worksheet_name or WORKSHEET_NAME_DEFAULT

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]

        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=scopes,
        )

        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(spreadsheet_id)

        try:
            ws = spreadsheet.worksheet(worksheet_name)
        except Exception:
            ws = spreadsheet.add_worksheet(
                title=worksheet_name,
                rows=1000,
                cols=len(HEADERS) + 5,
            )
            ws.append_row(HEADERS, value_input_option="USER_ENTERED")
            return ws

        values = ws.get_all_values()

        if not values:
            ws.append_row(HEADERS, value_input_option="USER_ENTERED")
        elif values[0] != HEADERS:
            ws.clear()
            ws.update("A1", [HEADERS])

        return ws

    except Exception as e:
        st.session_state["gsheet_error"] = traceback.format_exc()
        return None

def read_entries():
    ws = get_google_worksheet()
    if ws is None:
        return pd.DataFrame(columns=HEADERS)
    try:
        records = ws.get_all_records()
        df = pd.DataFrame(records)
    except Exception:
        df = pd.DataFrame(columns=HEADERS)
    for col in HEADERS:
        if col not in df.columns:
            df[col] = ""
    df = df[HEADERS].fillna("")
    return df

def append_entry(row):
    ws = get_google_worksheet()
    if ws is None:
        raise RuntimeError("Google Sheets not connected.")
    row = {col: row.get(col, "") for col in HEADERS}
    ws.append_row([row[col] for col in HEADERS], value_input_option="USER_ENTERED")

def rewrite_entries(df):
    ws = get_google_worksheet()
    if ws is None:
        raise RuntimeError("Google Sheets not connected.")
    for col in HEADERS:
        if col not in df.columns:
            df[col] = ""
    df = df[HEADERS].fillna("").astype(str)
    ws.clear()
    ws.update("A1", [HEADERS] + df.values.tolist())

def copy_row_style(ws, source_row, target_row):
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        tgt = ws.cell(target_row, col)
        if src.has_style:
            tgt.font = copy(src.font)
            tgt.border = copy(src.border)
            tgt.fill = copy(src.fill)
            tgt.number_format = src.number_format
            tgt.protection = copy(src.protection)
            tgt.alignment = copy(src.alignment)
            tgt._style = copy(src._style)
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height

def safe_set(ws, cell_ref, value):
    cell = ws[cell_ref]
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            ws.cell(merged_range.min_row, merged_range.min_col).value = value
            return
    cell.value = value

def clear_cell(ws, row, col):
    coord = ws.cell(row, col).coordinate
    for merged_range in ws.merged_cells.ranges:
        if coord in merged_range:
            if row == merged_range.min_row and col == merged_range.min_col:
                ws.cell(row, col).value = None
            return
    ws.cell(row, col).value = None

def generate_excel(df, employee):
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb[TEMPLATE_SHEET_NAME] if TEMPLATE_SHEET_NAME in wb.sheetnames else wb.active

    safe_set(ws, "A1", f"Travel Reimbursement Form : {employee.get('reportMonthText', '')}")
    safe_set(ws, "A2", f"Name : {employee.get('empName', '')}")
    safe_set(ws, "B2", f"Designation : {employee.get('designation', '')}")
    safe_set(ws, "C2", f"Location : {employee.get('location', '')}")

    filtered = df.copy()

    if employee.get("empName"):
        filtered = filtered[filtered["empName"].astype(str) == employee.get("empName")]
    if employee.get("reportMonthText"):
        filtered = filtered[filtered["reportMonthText"].astype(str) == employee.get("reportMonthText")]

    data_start = 5
    template_end = 36
    fixed_rows = template_end - data_start + 1
    entry_count = len(filtered)

    if entry_count > fixed_rows:
        extra_rows = entry_count - fixed_rows
        ws.insert_rows(template_end + 1, extra_rows)
        for r in range(template_end + 1, template_end + 1 + extra_rows):
            copy_row_style(ws, template_end, r)

    last_data_row = data_start + max(entry_count, fixed_rows) - 1

    for r in range(data_start, last_data_row + 1):
        for c in range(1, 16):
            clear_cell(ws, r, c)

    for r, (_, row) in enumerate(filtered.iterrows(), start=data_start):
        ws.cell(r, 1).value = row.get("date", "")
        ws.cell(r, 2).value = row.get("from", "")
        ws.cell(r, 3).value = row.get("to", "")
        ws.cell(r, 4).value = row.get("company", "")
        ws.cell(r, 5).value = row.get("contact", "")
        ws.cell(r, 6).value = row.get("invoice", "")
        ws.cell(r, 7).value = money(row.get("toll"))
        ws.cell(r, 8).value = money(row.get("fuel"))
        ws.cell(r, 9).value = 0
        ws.cell(r, 10).value = money(row.get("lodging"))
        ws.cell(r, 11).value = money(row.get("food"))
        ws.cell(r, 12).value = money(row.get("tel"))
        ws.cell(r, 13).value = money(row.get("courier"))
        ws.cell(r, 14).value = money(row.get("rikshaw"))
        ws.cell(r, 15).value = f"=SUM(G{r}:N{r})"

    total_row_1 = last_data_row + 1
    total_row_2 = last_data_row + 2

    safe_set(ws, f"N{total_row_1}", "Total ")
    safe_set(ws, f"O{total_row_1}", f"=SUM(O{data_start}:O{last_data_row})")
    safe_set(ws, f"N{total_row_2}", "Total ")
    safe_set(ws, f"O{total_row_2}", f"=O{total_row_1}")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def show_connection_status():
    st.subheader("Google Sheet Connection")
    ws = get_google_worksheet()

    if ws is not None:
        st.success("Storage: Google Sheets connected. Entries are saved live.")
    else:
        st.error("Storage: Google Sheets not connected.")
        with st.expander("Show Google Sheet connection error"):
            st.code(st.session_state.get("gsheet_error", "No error captured."))

    sheet_id = get_secret_value("gsheets", {}).get("spreadsheet_id", "")
    worksheet_name = get_secret_value("WORKSHEET_NAME", WORKSHEET_NAME_DEFAULT) or WORKSHEET_NAME_DEFAULT
    st.caption(f"Sheet ID: {sheet_id} | Worksheet: {worksheet_name}")

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")

    st.title("HULIOT INDIA")
    st.header("Travel Expense Entry App")
    st.caption("Save entries live in Google Sheets and download in your original Excel format.")

    if "employee" not in st.session_state:
        st.session_state.employee = default_employee()

    show_connection_status()
    st.divider()

    st.subheader("Employee Details")
    ec1, ec2, ec3, ec4 = st.columns(4)

    emp_name = ec1.text_input("Name", value=st.session_state.employee.get("empName", ""), key="emp_name")
    designation = ec2.text_input("Designation", value=st.session_state.employee.get("designation", ""), key="emp_designation")
    location = ec3.text_input("Location", value=st.session_state.employee.get("location", ""), key="emp_location")
    report_text = ec4.text_input("Month / Period Text", value=st.session_state.employee.get("reportMonthText", ""), key="emp_report_text")

    st.caption("These details are used in the Excel header and for filtering your saved entries.")

    employee = {
        "empName": make_str(emp_name),
        "designation": make_str(designation),
        "location": make_str(location),
        "reportMonthText": make_str(report_text),
    }
    st.session_state.employee = employee
    st.divider()

    with st.form("entry_form", clear_on_submit=True):
        st.subheader("Add Travel Expense")

        c1, c2, c3, c4 = st.columns(4)
        entry_date = c1.date_input("Date", value=date.today())
        from_loc = c2.text_input("From", placeholder="Start location")
        to_loc = c3.text_input("To", placeholder="End location")
        vehicle = c4.selectbox("2 Wheeler / 4 Wheeler", VEHICLE_OPTIONS, index=1)

        c5, c6 = st.columns(2)
        company = c5.text_input("Company Name / Contact / MEP Consultant", placeholder="Site / company / consultant name")
        contact = c6.text_input("Contact Person / Meeting With", placeholder="Person name")

        c7, c8, c9, c10 = st.columns(4)
        invoice = c7.text_input("Invoice No.", placeholder="Optional")
        toll = c8.number_input("Toll / Parking", min_value=0.0, value=0.0, step=1.0)
        fuel = c9.number_input("Petrol / Diesel", min_value=0.0, value=0.0, step=1.0)
        lodging = c10.number_input("Lodging / Boarding", min_value=0.0, value=0.0, step=1.0)

        c11, c12, c13, c14 = st.columns(4)
        food = c11.number_input("Food / Beverages", min_value=0.0, value=0.0, step=1.0)
        tel = c12.number_input("Tel / Internet", min_value=0.0, value=0.0, step=1.0)
        courier = c13.number_input("Courier / Stationary", min_value=0.0, value=0.0, step=1.0)
        rikshaw = c14.number_input("Rikshaw / Bus / Ola", min_value=0.0, value=0.0, step=1.0)

        remarks = st.text_area("Remarks / Purpose", placeholder="Meeting details, site visit purpose, or notes")

        save_clicked = st.form_submit_button("Save Entry")

        if save_clicked:
            try:
                if not make_str(from_loc) or not make_str(to_loc):
                    st.error("Please enter From and To.")
                else:
                    row = {
                        "id": str(uuid.uuid4())[:8],
                        "empName": employee["empName"],
                        "designation": employee["designation"],
                        "location": employee["location"],
                        "reportMonthText": employee["reportMonthText"],
                        "date": entry_date.strftime("%Y-%m-%d"),
                        "from": make_str(from_loc),
                        "to": make_str(to_loc),
                        "vehicle": make_str(vehicle),
                        "company": make_str(company),
                        "contact": make_str(contact),
                        "invoice": make_str(invoice),
                        "toll": toll,
                        "fuel": fuel,
                        "lodging": lodging,
                        "food": food,
                        "tel": tel,
                        "courier": courier,
                        "rikshaw": rikshaw,
                        "remarks": make_str(remarks),
                        "created_at": current_stamp(),
                    }
                    row["total"] = calc_total(row)
                    append_entry(row)
                    st.success("Entry saved successfully in Google Sheet.")
                    st.rerun()
            except Exception as e:
                st.error(f"Entry not saved. Google Sheet error: {e}")
                with st.expander("Show full error"):
                    st.code(traceback.format_exc())

    df = read_entries()
    filtered = df.copy()

    if employee["empName"]:
        filtered = filtered[filtered["empName"].astype(str) == employee["empName"]]
    if employee["reportMonthText"]:
        filtered = filtered[filtered["reportMonthText"].astype(str) == employee["reportMonthText"]]

    for col in AMOUNT_COLUMNS + ["total"]:
        if col not in filtered.columns:
            filtered[col] = 0

    total_entries = len(filtered)
    sum_toll = filtered["toll"].apply(money).sum() if total_entries else 0
    sum_fuel = filtered["fuel"].apply(money).sum() if total_entries else 0
    sum_stay_food = (filtered["lodging"].apply(money).sum() + filtered["food"].apply(money).sum()) if total_entries else 0
    grand_total = filtered["total"].apply(money).sum() if total_entries else 0

    st.divider()
    st.subheader("Expense Summary")
    s1, s2, s3, s4, s5 = st.columns(5)
    s1.metric("Total Entries", str(total_entries))
    s2.metric("Toll / Parking", f"₹{sum_toll:.2f}")
    s3.metric("Fuel", f"₹{sum_fuel:.2f}")
    s4.metric("Lodging + Food", f"₹{sum_stay_food:.2f}")
    s5.metric("Grand Total", f"₹{grand_total:.2f}")

    b1, b2, b3 = st.columns([1.4, 1.1, 1.1])
    with b1:
        try:
            excel_bytes = generate_excel(df, employee)
            st.download_button(
                "Download Excel Format",
                data=excel_bytes,
                file_name=f"Travelling_Expenses_{employee['empName'].replace(' ', '_').replace('.', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Excel download error: {e}")

    with b2:
        st.download_button(
            "Backup CSV",
            data=filtered.to_csv(index=False).encode("utf-8"),
            file_name="travel_expense_backup.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with b3:
        if st.button("Refresh", use_container_width=True):
            st.rerun()

    st.caption("Excel download uses your original workbook. Only employee details and entry cells are filled.")
    st.divider()

    st.subheader("Saved Entries (Edit & Delete Directly)")
    st.caption("📝 **To Edit:** Double-click any cell to change its value. <br> 🗑️ **To Delete:** Click the grey box on the far left of a row to select it, then press the **Delete / Trash** icon in the top right of the table.", unsafe_allow_html=True)

    if filtered.empty:
        st.info("No entries saved for this employee and period.")
    else:
        # Display the interactive data editor
        edited_df = st.data_editor(
            filtered,
            use_container_width=True,
            hide_index=False,
            num_rows="dynamic", # This allows row deletion
            disabled=["id", "total", "created_at", "empName", "designation", "location", "reportMonthText"], # Lock background columns
            column_order=["date", "from", "to", "company", "contact", "invoice", "vehicle", "toll", "fuel", "lodging", "food", "tel", "courier", "rikshaw", "total", "remarks"],
            key="expense_editor"
        )

        # Button to save the direct table edits to Google Sheets
        if st.button("💾 Save Table Changes to Google Sheets", type="primary"):
            try:
                # Recalculate totals just in case the user edited the money columns
                for idx, row in edited_df.iterrows():
                    edited_df.at[idx, "total"] = calc_total(row)

                # Fetch the main database
                full_df = read_entries()

                # Remove the old versions of this employee's current month records
                mask = (full_df["empName"] == employee["empName"]) & (full_df["reportMonthText"] == employee["reportMonthText"])
                full_df = full_df[~mask]

                # Add the newly edited/deleted records back in
                final_df = pd.concat([full_df, edited_df], ignore_index=True)

                # Overwrite Google Sheets with the updated data
                rewrite_entries(final_df)

                st.success("All edits and deletions have been successfully synced to Google Sheets!")
                st.rerun()
            except Exception as e:
                st.error(f"Failed to save changes: {e}")
                with st.expander("Show full error"):
                    st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
