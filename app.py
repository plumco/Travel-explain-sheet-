import os
import io
import uuid
from copy import copy
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

APP_TITLE = "Travel Expense Entry App"
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

CSS = """
<style>
    .stApp { background: #f4f7fb; }
    .main .block-container { max-width: 1160px; padding-top: 0.8rem; padding-bottom: 2rem; }
    .topbar {
        background: #233fb4;
        color: white;
        margin: -0.8rem -1rem 16px -1rem;
        padding: 16px 20px 18px;
        text-align: center;
        border-bottom-left-radius: 0;
        border-bottom-right-radius: 0;
    }
    .brand-pill {
        display: inline-block;
        font-size: 10px;
        border: 1px solid rgba(255,255,255,.55);
        border-radius: 999px;
        padding: 2px 10px;
        margin-bottom: 3px;
        letter-spacing: .3px;
    }
    .topbar h1 { margin: 0; font-size: 21px; line-height: 1.15; font-weight: 800; }
    .topbar p { margin: 4px 0 0; font-size: 12px; opacity: .95; }
    div[data-testid="stForm"], .section-card {
        background: white;
        border-radius: 10px;
        border: 1px solid #e5e7eb;
        box-shadow: 0 4px 14px rgba(15,23,42,.08);
        padding: 16px 16px 14px;
        margin-bottom: 14px;
    }
    .section-title { font-size: 22px; font-weight: 800; color: #111827; margin: 0 0 12px; }
    .note { font-size: 12px; color: #64748b; margin-top: 5px; }
    label, [data-testid="stWidgetLabel"] p { font-size: 11px !important; font-weight: 700 !important; color: #1f2937 !important; }
    input, textarea, select { border-radius: 6px !important; }
    .stButton>button, .stDownloadButton>button {
        border-radius: 6px; font-size: 13px; font-weight: 800; min-height: 36px;
    }
    .stDownloadButton>button { background: #16a34a; color: white; border: 0; }
    .small-summary {
        border: 1px solid #bfdbfe;
        background: #eff6ff;
        border-radius: 8px;
        padding: 11px;
        text-align: center;
    }
    .small-summary div:first-child { font-size: 11px; color: #475569; }
    .small-summary div:last-child { font-size: 16px; color: #1e3a8a; font-weight: 800; margin-top: 4px; }
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
    .excel-preview table { width:100%; border-collapse:collapse; background:white; font-size:11px; }
    .excel-preview th, .excel-preview td { border:1px solid #222; padding:5px; text-align:center; }
    .excel-preview .title { font-size:16px; font-weight:800; }
    .excel-preview .head { font-weight:800; }
</style>
"""


def money(value):
    try:
        if value in (None, ""):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def clean_text(value):
    if value is None:
        return ""
    return str(value).strip()


@st.cache_resource(show_spinner=False)
def get_google_worksheet():
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        if "gcp_service_account" not in st.secrets or "gsheets" not in st.secrets:
            return None
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scopes)
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(st.secrets["gsheets"]["spreadsheet_id"])
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
            # Do not overwrite old sheets. Use this app's exact-format worksheet only.
            ws.clear()
            ws.update("A1", [COLUMNS])
        return ws
    except Exception:
        return None


def read_entries() -> pd.DataFrame:
    ws = get_google_worksheet()
    if ws is not None:
        try:
            records = ws.get_all_records()
            df = pd.DataFrame(records)
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

    report_text = employee.get("report_text", "") or "01st month 2026 To 28th month 2026"
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


def excel_like_preview(employee):
    return f"""
    <div class="excel-preview section-card">
      <table>
        <tr><td colspan="15" class="title">Travel Reimbursement Form : {employee.get('report_text','')}</td></tr>
        <tr><td colspan="3">Name : {employee.get('name','')}</td><td colspan="4">Designation : {employee.get('designation','')}</td><td colspan="8">Location : {employee.get('location','')}</td></tr>
        <tr class="head"><td rowspan="2">DATE</td><td colspan="2">Travel Location</td><td rowspan="2">Company Name / Contact<br>/ MEP Consultant</td><td rowspan="2">Contact Person / With whom<br>you attended meeting</td><td rowspan="2">Invoice No.<br>(if any)</td><td rowspan="2">Toll / Parking<br>Charges</td><td rowspan="2">Petrol / Diesel<br>(Own vehicle)</td><td rowspan="2">2 Wheeler /<br>4 Wheeler</td><td rowspan="2">Lodging /<br>Boarding</td><td rowspan="2">Food /<br>Beverages</td><td rowspan="2">Tel / Internet</td><td rowspan="2">Courier /<br>Stationary</td><td rowspan="2">Rikshaw / Bus / Ola</td><td rowspan="2">TOTAL</td></tr>
        <tr class="head"><td>FROM</td><td>TO</td></tr>
      </table>
    </div>
    """


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown('<div class="topbar"><span class="brand-pill">HULIOT INDIA</span><h1>Travel Expense Entry App</h1><p>Same as attached HTML form and original Excel format.</p></div>', unsafe_allow_html=True)

    if "employee" not in st.session_state:
        st.session_state.employee = {
            "name": "Mr.Umesh Nikam",
            "designation": "Executive-Technical Support",
            "location": "Pune",
            "report_text": "01st month 2026 To 28th month 2026",
        }

    st.markdown('<div class="section-card"><div class="section-title">Employee Details</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns([1, 1.3, 1, 1.4])
    name = c1.text_input("Name", value=st.session_state.employee.get("name", ""))
    designation = c2.text_input("Designation", value=st.session_state.employee.get("designation", ""))
    location = c3.text_input("Location", value=st.session_state.employee.get("location", ""))
    report_text = c4.text_input("Month / Text", value=st.session_state.employee.get("report_text", ""))
    st.markdown('<div class="note">These details are used in the downloaded Excel header.</div></div>', unsafe_allow_html=True)
    employee = {"name": name, "designation": designation, "location": location, "report_text": report_text}
    st.session_state.employee = employee

    st.markdown(excel_like_preview(employee), unsafe_allow_html=True)

    with st.form("entry_form", clear_on_submit=True):
        st.markdown('<div class="section-title">Add Travel Expense</div>', unsafe_allow_html=True)
        r1c1, r1c2, r1c3, r1c4 = st.columns([1, 1.3, 1.3, 1])
        entry_date = r1c1.date_input("Date", value=date.today())
        from_loc = r1c2.text_input("From", placeholder="Start location")
        to_loc = r1c3.text_input("To", placeholder="End location")
        two_four = r1c4.number_input("2 Wheeler / 4 Wheeler", min_value=0.0, step=1.0, value=0.0)

        r2c1, r2c2 = st.columns(2)
        company = r2c1.text_input("Company Name / Contact / MEP Consultant", placeholder="Site / consultant name")
        contact = r2c2.text_input("Contact Person / Meeting With", placeholder="Person name")

        r3c1, r3c2, r3c3, r3c4 = st.columns(4)
        invoice = r3c1.text_input("Invoice No.", placeholder="Optional")
        toll = r3c2.number_input("Toll / Parking", min_value=0.0, step=1.0, value=0.0)
        petrol = r3c3.number_input("Petrol / Diesel", min_value=0.0, step=1.0, value=0.0)
        lodging = r3c4.number_input("Lodging / Boarding", min_value=0.0, step=1.0, value=0.0)

        r4c1, r4c2, r4c3, r4c4 = st.columns(4)
        food = r4c1.number_input("Food / Beverages", min_value=0.0, step=1.0, value=0.0)
        tel = r4c2.number_input("Tel / Internet", min_value=0.0, step=1.0, value=0.0)
        courier = r4c3.number_input("Courier / Stationary", min_value=0.0, step=1.0, value=0.0)
        rikshaw = r4c4.number_input("Rikshaw / Bus / Ola", min_value=0.0, step=1.0, value=0.0)

        remarks = st.text_area("Remarks / Purpose", placeholder="Meeting details, site visit purpose, or notes")
        save = st.form_submit_button("Save Entry", use_container_width=False)
        if save:
            if not from_loc or not to_loc:
                st.error("Please fill From and To.")
            else:
                row = {
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "EntryID": str(uuid.uuid4()),
                    "Name": name,
                    "Designation": designation,
                    "Location": location,
                    "ReportText": report_text,
                    "Date": entry_date.strftime("%d-%m-%Y"),
                    "From": from_loc,
                    "To": to_loc,
                    "CompanyConsultant": company,
                    "ContactPerson": contact,
                    "InvoiceNo": invoice,
                    "TollParking": toll,
                    "PetrolDiesel": petrol,
                    "TwoWheelerFourWheeler": two_four,
                    "LodgingBoarding": lodging,
                    "FoodBeverages": food,
                    "TelInternet": tel,
                    "CourierStationary": courier,
                    "RikshawBusOla": rikshaw,
                    "Remarks": remarks,
                }
                save_entry(row)
                st.success("Entry saved.")

    df = read_entries()
    filtered = df.copy()
    if name:
        filtered = filtered[filtered["Name"].astype(str) == name]
    if report_text:
        filtered = filtered[filtered["ReportText"].astype(str) == report_text]

    total = 0.0
    for col in AMOUNT_COLUMNS:
        if col in filtered.columns:
            total += filtered[col].apply(money).sum()

    st.markdown('<div class="section-card"><div class="section-title">Expense Summary</div>', unsafe_allow_html=True)
    s1, s2, s3, s4, s5 = st.columns(5)
    s1.markdown(f'<div class="small-summary"><div>Total Entries</div><div>{len(filtered)}</div></div>', unsafe_allow_html=True)
    s2.markdown(f'<div class="small-summary"><div>Toll / Parking</div><div>₹{filtered["TollParking"].apply(money).sum() if not filtered.empty else 0:.2f}</div></div>', unsafe_allow_html=True)
    s3.markdown(f'<div class="small-summary"><div>Fuel</div><div>₹{filtered["PetrolDiesel"].apply(money).sum() if not filtered.empty else 0:.2f}</div></div>', unsafe_allow_html=True)
    s4.markdown(f'<div class="small-summary"><div>Lodging + Food</div><div>₹{(filtered["LodgingBoarding"].apply(money).sum() + filtered["FoodBeverages"].apply(money).sum()) if not filtered.empty else 0:.2f}</div></div>', unsafe_allow_html=True)
    s5.markdown(f'<div class="small-summary"><div>Grand Total</div><div>₹{total:.2f}</div></div>', unsafe_allow_html=True)

    excel_bytes = generate_excel(df, employee)
    b1, b2, b3 = st.columns([1.2, 1, 3])
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
            "Export Simple CSV",
            data=filtered.to_csv(index=False).encode("utf-8"),
            file_name="travel_entries_backup.csv",
            mime="text/csv",
            use_container_width=True,
        )
    st.markdown('<div class="note">Excel download uses your original workbook template. Only header details and entry rows are filled.</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card"><div class="section-title">Saved Entries</div>', unsafe_allow_html=True)
    display_cols = [
        "Date", "From", "To", "CompanyConsultant", "ContactPerson", "InvoiceNo", "TollParking", "PetrolDiesel",
        "TwoWheelerFourWheeler", "LodgingBoarding", "FoodBeverages", "TelInternet", "CourierStationary", "RikshawBusOla", "Remarks", "EntryID"
    ]
    display_cols = [c for c in display_cols if c in filtered.columns]
    st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("Delete wrong entry"):
        st.caption("Copy EntryID from Saved Entries and paste below.")
        entry_id = st.text_input("EntryID to delete")
        if st.button("Delete Entry") and entry_id:
            new_df = df[df["EntryID"].astype(str) != str(entry_id)].copy()
            rewrite_entries(new_df)
            st.warning("Entry deleted. Refresh the app once.")

    with st.sidebar:
        st.markdown("### Storage")
        if get_google_worksheet() is not None:
            st.success("Google Sheets connected")
        else:
            st.warning("Google Sheets not connected. Local CSV fallback active.")
        st.markdown("### Important")
        st.write("Use this version for the exact original Excel columns.")


if __name__ == "__main__":
    main()
