import os
import io
import uuid
from copy import copy
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

APP_TITLE = "Huliot Travel Expense Entry"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, "Travelling Expenses Sheet.xlsx")
LOCAL_FILE = os.path.join(BASE_DIR, "local_travel_entries.csv")
DATA_SHEET_NAME = "Travel_Expense_Entries"
TEMPLATE_SHEET_NAME = "Travel Reimbur. Form"

COLUMNS = [
    "Timestamp", "EntryID",
    "Name", "Designation", "Location", "ReportText",
    "Date", "From", "To", "CompanyConsultant", "ContactPerson", "InvoiceNo",
    "TollParking", "PetrolDiesel", "TwoWheelerFourWheeler", "LodgingBoarding",
    "FoodBeverages", "TelInternet", "CourierStationary", "RikshawBusOla", "Remarks"
]

AMOUNT_COLUMNS = [
    "TollParking", "PetrolDiesel", "TwoWheelerFourWheeler", "LodgingBoarding",
    "FoodBeverages", "TelInternet", "CourierStationary", "RikshawBusOla"
]

DISPLAY_RENAME = {
    "Date": "Date", "From": "From", "To": "To",
    "CompanyConsultant": "Company / Consultant", "ContactPerson": "Contact",
    "InvoiceNo": "Invoice", "TollParking": "Toll", "PetrolDiesel": "Fuel",
    "TwoWheelerFourWheeler": "2W / 4W", "LodgingBoarding": "Lodging",
    "FoodBeverages": "Food", "TelInternet": "Tel", "CourierStationary": "Courier",
    "RikshawBusOla": "Rikshaw", "Remarks": "Remarks", "EntryID": "EntryID"
}

CSS = """
<style>
    /* Same visual direction as uploaded HTML file */
    .stApp { background: #f4f7fb; color: #1f2937; }
    .main .block-container { max-width: 1180px; padding-top: 0; padding-bottom: 34px; }
    header, .html-header {
        background: #1e40af;
        color: white;
        padding: 18px 16px;
        text-align: center;
        margin: 0 -2rem 18px -2rem;
    }
    .html-header .brand {
        display: inline-block;
        margin-bottom: 8px;
        padding: 5px 10px;
        border: 1px solid rgba(255,255,255,.55);
        border-radius: 999px;
        font-size: 12px;
        letter-spacing: .3px;
        background: rgba(255,255,255,.12);
    }
    .html-header h1 { margin: 0; font-size: 22px; line-height: 1.1; }
    .html-header p { margin: 6px 0 0; font-size: 14px; opacity: .9; }

    div[data-testid="stForm"], .card-box {
        background: white;
        border-radius: 14px;
        padding: 16px;
        margin-bottom: 16px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, .08);
        border: 0;
    }
    .card-title { margin: 0 0 14px 0; font-size: 22px; font-weight: 800; color: #111827; }
    .muted { color: #64748b; font-size: 13px; margin-top: 8px; }
    .ok-text { color: #166534; font-weight: 700; font-size: 13px; }
    .warn-text { color: #92400e; font-weight: 700; font-size: 13px; }

    label, [data-testid="stWidgetLabel"] p {
        font-size: 12px !important;
        font-weight: 700 !important;
        color: #374151 !important;
        margin-bottom: 5px !important;
    }
    input, textarea, select {
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        font-size: 14px !important;
        background: white !important;
    }
    textarea { min-height: 70px !important; }
    div[data-baseweb="input"] > div, div[data-baseweb="select"] > div, div[data-baseweb="textarea"] > div {
        border-radius: 8px !important;
        border-color: #cbd5e1 !important;
    }
    .stButton>button, .stDownloadButton>button {
        border: 0;
        border-radius: 8px;
        padding: 10px 14px;
        font-size: 14px;
        font-weight: 800;
        min-height: 42px;
    }
    div[data-testid="stFormSubmitButton"] button {
        background: #2563eb;
        color: white;
    }
    .stDownloadButton>button {
        background: #16a34a;
        color: white;
    }
    button[kind="secondary"] { background: #e5e7eb; color: #111827; }
    .summary-box {
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        border-radius: 12px;
        padding: 12px;
        text-align: center;
        min-height: 74px;
    }
    .summary-box span { display: block; font-size: 12px; color: #475569; }
    .summary-box strong { display: block; margin-top: 5px; font-size: 18px; color: #1e3a8a; }
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
    .block-gap { height: 8px; }

    @media (max-width: 800px) {
      .html-header h1 { font-size: 19px; }
    }
</style>
"""


def money(value):
    try:
        if value in (None, ""):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def make_str(value):
    if value is None:
        return ""
    return str(value).strip()


def current_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_entry_date(value):
    text = make_str(value)
    for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            pass
    return date.today()


@st.cache_resource(show_spinner=False)
def get_google_worksheet():
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        if "gcp_service_account" not in st.secrets or "gsheets" not in st.secrets:
            return None
        spreadsheet_id = st.secrets["gsheets"].get("spreadsheet_id", "")
        if not spreadsheet_id:
            return None

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes
        )
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(spreadsheet_id)
        try:
            ws = spreadsheet.worksheet(DATA_SHEET_NAME)
        except Exception:
            ws = spreadsheet.add_worksheet(title=DATA_SHEET_NAME, rows=1000, cols=len(COLUMNS) + 2)
            ws.append_row(COLUMNS)
            return ws

        values = ws.get_all_values()
        if not values:
            ws.append_row(COLUMNS)
        elif values[0] != COLUMNS:
            ws.clear()
            ws.update("A1", [COLUMNS])
        return ws
    except Exception:
        return None


def read_entries() -> pd.DataFrame:
    ws = get_google_worksheet()
    if ws is not None:
        try:
            df = pd.DataFrame(ws.get_all_records())
        except Exception:
            df = pd.DataFrame(columns=COLUMNS)
    elif os.path.exists(LOCAL_FILE):
        df = pd.read_csv(LOCAL_FILE)
    else:
        df = pd.DataFrame(columns=COLUMNS)

    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[COLUMNS].fillna("")


def save_entry(row: dict):
    row = {col: row.get(col, "") for col in COLUMNS}
    ws = get_google_worksheet()
    if ws is not None:
        ws.append_row([row[col] for col in COLUMNS], value_input_option="USER_ENTERED")
    else:
        df = read_entries()
        df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
        df.to_csv(LOCAL_FILE, index=False)


def rewrite_entries(df: pd.DataFrame):
    df = df.copy()
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS].fillna("")
    ws = get_google_worksheet()
    if ws is not None:
        ws.clear()
        ws.update("A1", [COLUMNS] + df.astype(str).values.tolist())
    else:
        df.to_csv(LOCAL_FILE, index=False)


def update_entry_by_id(entry_id: str, updated_row: dict):
    df = read_entries()
    mask = df["EntryID"].astype(str) == str(entry_id)
    if not mask.any():
        return False
    updated_row = {col: updated_row.get(col, "") for col in COLUMNS}
    for col in COLUMNS:
        df.loc[mask, col] = updated_row[col]
    rewrite_entries(df)
    return True


def copy_row_style(ws, source_row: int, target_row: int):
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


def safe_set(ws, cell_ref: str, value):
    cell = ws[cell_ref]
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            ws.cell(mr.min_row, mr.min_col).value = value
            return
    cell.value = value


def clear_cell(ws, row: int, col: int):
    coord = ws.cell(row, col).coordinate
    for mr in ws.merged_cells.ranges:
        if coord in mr:
            if row == mr.min_row and col == mr.min_col:
                ws.cell(row, col).value = None
            return
    ws.cell(row, col).value = None


def generate_excel(df: pd.DataFrame, employee: dict) -> bytes:
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb[TEMPLATE_SHEET_NAME] if TEMPLATE_SHEET_NAME in wb.sheetnames else wb.active

    report_text = employee.get("report_text", "")
    safe_set(ws, "A1", f"Travel Reimbursement Form : {report_text}")
    safe_set(ws, "A2", f"Name : {employee.get('name', '')}")
    safe_set(ws, "B2", f"Designation : {employee.get('designation', '')}")
    safe_set(ws, "C2", f"Location : {employee.get('location', '')}")

    filtered = df.copy()
    if employee.get("name"):
        filtered = filtered[filtered["Name"].astype(str) == employee.get("name")]
    if employee.get("report_text"):
        filtered = filtered[filtered["ReportText"].astype(str) == employee.get("report_text")]

    data_start = 5
    template_end = 36
    fixed_rows = template_end - data_start + 1
    entry_count = len(filtered)

    if entry_count > fixed_rows:
        extra = entry_count - fixed_rows
        ws.insert_rows(template_end + 1, extra)
        for r in range(template_end + 1, template_end + 1 + extra):
            copy_row_style(ws, template_end, r)

    last_data_row = data_start + max(entry_count, fixed_rows) - 1
    for r in range(data_start, last_data_row + 1):
        for c in range(1, 16):
            clear_cell(ws, r, c)

    for r, (_, row) in enumerate(filtered.iterrows(), start=data_start):
        ws.cell(r, 1).value = row.get("Date", "")
        ws.cell(r, 2).value = row.get("From", "")
        ws.cell(r, 3).value = row.get("To", "")
        ws.cell(r, 4).value = row.get("CompanyConsultant", "")
        ws.cell(r, 5).value = row.get("ContactPerson", "")
        ws.cell(r, 6).value = row.get("InvoiceNo", "")
        ws.cell(r, 7).value = money(row.get("TollParking"))
        ws.cell(r, 8).value = money(row.get("PetrolDiesel"))
        ws.cell(r, 9).value = money(row.get("TwoWheelerFourWheeler"))
        ws.cell(r, 10).value = money(row.get("LodgingBoarding"))
        ws.cell(r, 11).value = money(row.get("FoodBeverages"))
        ws.cell(r, 12).value = money(row.get("TelInternet"))
        ws.cell(r, 13).value = money(row.get("CourierStationary"))
        ws.cell(r, 14).value = money(row.get("RikshawBusOla"))
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


def default_employee():
    return {
        "name": "Mr.Umesh Nikam",
        "designation": "Executive-Technical Support",
        "location": "Pune",
        "report_text": "01st April 2026 To 30th April 2026",
    }


def summary_box(label, value):
    st.markdown(f'<div class="summary-box"><span>{label}</span><strong>{value}</strong></div>', unsafe_allow_html=True)


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(
        '<div class="html-header"><div class="brand">HULIOT INDIA</div>'
        '<h1>Travel Expense Entry App</h1>'
        '<p>Enter daily travel expenses, save records live, and download in your original Excel format.</p></div>',
        unsafe_allow_html=True,
    )

    if "employee" not in st.session_state:
        st.session_state.employee = default_employee()

    # Employee details card
    st.markdown('<div class="card-box"><h2 class="card-title">Employee Details</h2>', unsafe_allow_html=True)
    ec1, ec2, ec3, ec4 = st.columns(4)
    name = ec1.text_input("Name", value=st.session_state.employee.get("name", ""), key="emp_name")
    designation = ec2.text_input("Designation", value=st.session_state.employee.get("designation", ""), key="emp_designation")
    location = ec3.text_input("Location", value=st.session_state.employee.get("location", ""), key="emp_location")
    report_text = ec4.text_input("Month / Period Text", value=st.session_state.employee.get("report_text", ""), key="emp_report_text")
    st.markdown('<p class="muted">These details are used in the Excel header and for filtering your saved entries.</p></div>', unsafe_allow_html=True)

    employee = {"name": name, "designation": designation, "location": location, "report_text": report_text}
    st.session_state.employee = employee

    # Entry form card. This follows the uploaded HTML form layout.
    with st.form("entry_form", clear_on_submit=True):
        st.markdown('<h2 class="card-title">Add Travel Expense</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        entry_date = c1.date_input("Date", value=date.today())
        from_loc = c2.text_input("From", placeholder="Start location")
        to_loc = c3.text_input("To", placeholder="End location")
        two_four = c4.number_input("2 Wheeler / 4 Wheeler", min_value=0.0, value=0.0, step=1.0)

        c5, c6 = st.columns(2)
        company = c5.text_input("Company Name / Contact / MEP Consultant", placeholder="Site / company / consultant name")
        contact = c6.text_input("Contact Person / Meeting With", placeholder="Person name")

        c7, c8, c9, c10 = st.columns(4)
        invoice = c7.text_input("Invoice No.", placeholder="Optional")
        toll = c8.number_input("Toll / Parking", min_value=0.0, value=0.0, step=1.0)
        petrol = c9.number_input("Petrol / Diesel", min_value=0.0, value=0.0, step=1.0)
        lodging = c10.number_input("Lodging / Boarding", min_value=0.0, value=0.0, step=1.0)

        c11, c12, c13, c14 = st.columns(4)
        food = c11.number_input("Food / Beverages", min_value=0.0, value=0.0, step=1.0)
        tel = c12.number_input("Tel / Internet", min_value=0.0, value=0.0, step=1.0)
        courier = c13.number_input("Courier / Stationary", min_value=0.0, value=0.0, step=1.0)
        rikshaw = c14.number_input("Rikshaw / Bus / Ola", min_value=0.0, value=0.0, step=1.0)

        remarks = st.text_area("Remarks / Purpose", placeholder="Meeting details, site visit purpose, or notes")
        save = st.form_submit_button("Save Entry")

        if save:
            if not make_str(from_loc) or not make_str(to_loc):
                st.error("Please enter From and To.")
            else:
                save_entry({
                    "Timestamp": current_stamp(),
                    "EntryID": str(uuid.uuid4())[:8],
                    "Name": make_str(name),
                    "Designation": make_str(designation),
                    "Location": make_str(location),
                    "ReportText": make_str(report_text),
                    "Date": entry_date.strftime("%d-%m-%Y"),
                    "From": make_str(from_loc),
                    "To": make_str(to_loc),
                    "CompanyConsultant": make_str(company),
                    "ContactPerson": make_str(contact),
                    "InvoiceNo": make_str(invoice),
                    "TollParking": toll,
                    "PetrolDiesel": petrol,
                    "TwoWheelerFourWheeler": two_four,
                    "LodgingBoarding": lodging,
                    "FoodBeverages": food,
                    "TelInternet": tel,
                    "CourierStationary": courier,
                    "RikshawBusOla": rikshaw,
                    "Remarks": make_str(remarks),
                })
                st.success("Entry saved successfully.")

    df = read_entries()
    filtered = df.copy()
    if name:
        filtered = filtered[filtered["Name"].astype(str) == name]
    if report_text:
        filtered = filtered[filtered["ReportText"].astype(str) == report_text]

    for col in AMOUNT_COLUMNS:
        if col not in filtered.columns:
            filtered[col] = 0

    total_entries = len(filtered)
    sum_toll = filtered["TollParking"].apply(money).sum() if total_entries else 0
    sum_fuel = filtered["PetrolDiesel"].apply(money).sum() if total_entries else 0
    sum_stay_food = (filtered["LodgingBoarding"].apply(money).sum() + filtered["FoodBeverages"].apply(money).sum()) if total_entries else 0
    grand_total = sum(filtered[col].apply(money).sum() for col in AMOUNT_COLUMNS) if total_entries else 0

    st.markdown('<div class="card-box"><h2 class="card-title">Expense Summary</h2>', unsafe_allow_html=True)
    s1, s2, s3, s4, s5 = st.columns(5)
    with s1: summary_box("Total Entries", str(total_entries))
    with s2: summary_box("Toll / Parking", f"₹{sum_toll:.2f}")
    with s3: summary_box("Fuel", f"₹{sum_fuel:.2f}")
    with s4: summary_box("Lodging + Food", f"₹{sum_stay_food:.2f}")
    with s5: summary_box("Grand Total", f"₹{grand_total:.2f}")

    st.markdown('<div class="block-gap"></div>', unsafe_allow_html=True)
    excel_bytes = generate_excel(df, employee)
    b1, b2, b3, b4 = st.columns([1.4, 1.1, 1.1, 2.2])
    with b1:
        st.download_button(
            "Download Same Excel Format",
            data=excel_bytes,
            file_name=f"Travelling_Expenses_{name.replace(' ', '_').replace('.', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
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

    if get_google_worksheet() is not None:
        st.markdown('<p class="ok-text">Storage: Google Sheets connected. Entries are saved live.</p>', unsafe_allow_html=True)
    else:
        st.markdown('<p class="warn-text">Storage: Google Sheets not connected. Local CSV fallback is active for testing.</p>', unsafe_allow_html=True)
    st.markdown('<p class="muted">Excel download uses your original workbook. Only employee details and entry cells are filled.</p></div>', unsafe_allow_html=True)

    st.markdown('<div class="card-box"><h2 class="card-title">Saved Entries</h2>', unsafe_allow_html=True)
    if filtered.empty:
        st.info("No entries saved for this employee and period.")
    else:
        show_cols = [
            "Date", "From", "To", "CompanyConsultant", "ContactPerson", "InvoiceNo",
            "TollParking", "PetrolDiesel", "TwoWheelerFourWheeler", "LodgingBoarding", "FoodBeverages",
            "TelInternet", "CourierStationary", "RikshawBusOla", "Remarks", "EntryID"
        ]
        show_cols = [c for c in show_cols if c in filtered.columns]
        display = filtered[show_cols].rename(columns=DISPLAY_RENAME)
        st.dataframe(display, use_container_width=True, hide_index=True, height=330)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-box"><h2 class="card-title">Edit / Delete Entry</h2>', unsafe_allow_html=True)
    if filtered.empty:
        st.info("No entry available to edit.")
    else:
        edit_options = []
        for _, r in filtered.iterrows():
            label = f"{r.get('Date', '')} | {r.get('From', '')} to {r.get('To', '')} | {r.get('CompanyConsultant', '')} | ID: {r.get('EntryID', '')}"
            edit_options.append((label, str(r.get('EntryID', ''))))

        selected_label = st.selectbox("Select entry to edit", [x[0] for x in edit_options], key="edit_select")
        selected_entry_id = dict(edit_options).get(selected_label, "")
        selected_rows = filtered[filtered["EntryID"].astype(str) == selected_entry_id]

        if not selected_rows.empty:
            selected = selected_rows.iloc[0].to_dict()
            with st.form("edit_entry_form"):
                e1, e2, e3, e4 = st.columns(4)
                edit_date = e1.date_input("Date", value=parse_entry_date(selected.get("Date")), key="edit_date")
                edit_from = e2.text_input("From", value=make_str(selected.get("From")), key="edit_from")
                edit_to = e3.text_input("To", value=make_str(selected.get("To")), key="edit_to")
                edit_two_four = e4.number_input("2 Wheeler / 4 Wheeler", min_value=0.0, value=money(selected.get("TwoWheelerFourWheeler")), step=1.0, key="edit_two_four")

                e5, e6 = st.columns(2)
                edit_company = e5.text_input("Company Name / Contact / MEP Consultant", value=make_str(selected.get("CompanyConsultant")), key="edit_company")
                edit_contact = e6.text_input("Contact Person / Meeting With", value=make_str(selected.get("ContactPerson")), key="edit_contact")

                e7, e8, e9, e10 = st.columns(4)
                edit_invoice = e7.text_input("Invoice No.", value=make_str(selected.get("InvoiceNo")), key="edit_invoice")
                edit_toll = e8.number_input("Toll / Parking", min_value=0.0, value=money(selected.get("TollParking")), step=1.0, key="edit_toll")
                edit_petrol = e9.number_input("Petrol / Diesel", min_value=0.0, value=money(selected.get("PetrolDiesel")), step=1.0, key="edit_petrol")
                edit_lodging = e10.number_input("Lodging / Boarding", min_value=0.0, value=money(selected.get("LodgingBoarding")), step=1.0, key="edit_lodging")

                e11, e12, e13, e14 = st.columns(4)
                edit_food = e11.number_input("Food / Beverages", min_value=0.0, value=money(selected.get("FoodBeverages")), step=1.0, key="edit_food")
                edit_tel = e12.number_input("Tel / Internet", min_value=0.0, value=money(selected.get("TelInternet")), step=1.0, key="edit_tel")
                edit_courier = e13.number_input("Courier / Stationary", min_value=0.0, value=money(selected.get("CourierStationary")), step=1.0, key="edit_courier")
                edit_rikshaw = e14.number_input("Rikshaw / Bus / Ola", min_value=0.0, value=money(selected.get("RikshawBusOla")), step=1.0, key="edit_rikshaw")

                edit_remarks = st.text_area("Remarks / Purpose", value=make_str(selected.get("Remarks")), key="edit_remarks")

                u1, u2 = st.columns(2)
                update_clicked = u1.form_submit_button("Update Selected Entry")
                delete_clicked = u2.form_submit_button("Delete Selected Entry")

                if update_clicked:
                    if not make_str(edit_from) or not make_str(edit_to):
                        st.error("Please enter From and To.")
                    else:
                        updated = {
                            "Timestamp": current_stamp(),
                            "EntryID": selected_entry_id,
                            "Name": make_str(name),
                            "Designation": make_str(designation),
                            "Location": make_str(location),
                            "ReportText": make_str(report_text),
                            "Date": edit_date.strftime("%d-%m-%Y"),
                            "From": make_str(edit_from),
                            "To": make_str(edit_to),
                            "CompanyConsultant": make_str(edit_company),
                            "ContactPerson": make_str(edit_contact),
                            "InvoiceNo": make_str(edit_invoice),
                            "TollParking": edit_toll,
                            "PetrolDiesel": edit_petrol,
                            "TwoWheelerFourWheeler": edit_two_four,
                            "LodgingBoarding": edit_lodging,
                            "FoodBeverages": edit_food,
                            "TelInternet": edit_tel,
                            "CourierStationary": edit_courier,
                            "RikshawBusOla": edit_rikshaw,
                            "Remarks": make_str(edit_remarks),
                        }
                        if update_entry_by_id(selected_entry_id, updated):
                            st.success("Entry updated successfully.")
                            st.rerun()
                        else:
                            st.error("Entry not found. Please refresh and try again.")

                if delete_clicked:
                    new_df = df[df["EntryID"].astype(str) != str(selected_entry_id)].copy()
                    rewrite_entries(new_df)
                    st.success("Entry deleted successfully.")
                    st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
