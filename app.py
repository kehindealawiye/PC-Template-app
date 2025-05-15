import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import numbers
import pandas as pd
import io
import os
from datetime import datetime
from num2words import num2words
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

def get_gsheet_client():
    try:
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])

        # FIX: Convert escaped \\n back to actual line breaks in private key
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        gc = gspread.authorize(creds)
        st.success("Google Sheet client connected successfully.")
        return gc

    except Exception as e:
        st.error(f"Google Sheet auth failed: {e}")
        return None
        
def delete_backup_from_gsheet(user, timestamp):
    try:
        gc = get_gsheet_client()
        sheet = gc.open("PC_Backups").sheet1
        rows = sheet.get_all_records()

        # Keep only rows not matching the selected timestamp and user
        filtered_rows = [
            [row["user"], row["timestamp"], row["field_key"], row["value"]]
            for row in rows
            if not (row["user"].strip().lower() == user.strip().lower() and row["timestamp"] == timestamp)
        ]

        # Clear the entire sheet first (remove all rows)
        sheet.clear()

        # Re-insert the headers
        sheet.append_row(["user", "timestamp", "field_key", "value"])

        # Re-insert filtered rows
        for row in filtered_rows:
            sheet.append_row(row)

        return True
    except Exception as e:
        st.sidebar.error(f"Failed to delete backup: {e}")
        return False
        

def save_backup_to_gsheet(user, inputs_dict):
    gc = get_gsheet_client()
    if not gc:
        return  # If auth failed, skip saving

    try:
        sheet = gc.open("PC_Backups").sheet1
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for k, v in inputs_dict.items():
            sheet.append_row([user, timestamp, k, v])
        st.success("Backup written to Google Sheet.")
    except Exception as e:
        st.error(f"Failed to write to Google Sheet: {e}")
        

# === Page Setup ===
st.set_page_config(page_title="Prepayment Certificate App", layout="wide")
st.title("Prepayment Certificate Filler")


# === Template and Project Setup ===
project_count = st.selectbox("Number of Projects", [1, 2, 3], key="project_count_select")
template_paths = {
    1: "PC Template.xlsx",
    2: "PC Template 2.xlsx",
    3: "PC Template 3.xlsx"
}
project_columns = {
    1: {1: "B"},
    2: {1: "B", 2: "E"},
    3: {1: "B", 2: "E", 3: "H"}
}
details_sheet = "DETAILS"
naira_rows = {"10", "11", "13", "15", "18"}

def load_template(project_count):
    return load_workbook(template_paths[project_count])

# === User Identity (Prompt First) ===
if "user_confirmed" not in st.session_state:
    with st.form("user_form"):
        name_input = st.text_input("Enter Your Name to Continue:", "")
        submitted = st.form_submit_button("Enter")
        if submitted and name_input.strip():
            st.session_state["current_user"] = name_input.strip()
            st.session_state["user_confirmed"] = True
            st.rerun()
    st.stop()

# Admin check based on name
user = st.session_state["current_user"]
is_admin = (user.strip().lower() == "kehinde alawiye".lower())

# Backup folder setup
backup_root = "backups"
user_backup_dir = os.path.join(backup_root, user.replace(" ", "_"))
os.makedirs(user_backup_dir, exist_ok=True)

# === Excel Template Preview ===
st.sidebar.markdown("### Excel Template Preview")
try:
    preview_wb = load_template(project_count)
    st.sidebar.success(f"Template loaded: {template_paths[project_count]}")
except Exception as e:
    st.sidebar.error(f"Failed to load Excel template: {e}")

# === Dropdown Options ===
custom_dropdowns = {
    "Payment stage:": ["Stage Payment", "Final Payment", "Retention"],
    "Percentage of Advance payment? (as specified in the award letter)": ["0%", "25%", "40%", "50%", "60%", "70%"],
    "Is there 5% retention?": ["0%", "5%"],
    "Vat": ["0%", "7.5%"],
    "Physical Stage of Work": ["Ongoing", "Completed"],
    "Address line 1": ["The Director,", "The Chairman,", "The Permanent Secretary,", "The General Manager,", "The Honourable Commissioner,", "The Special Adviser,"]
}

# === Load Field Definitions ===
@st.cache_data
def load_field_structure():
    df = pd.read_csv("Grouped_Field_Structure_Clean.csv")
    grouped = {}
    for _, row in df.iterrows():
        group = row['Group']
        field = (str(row['Row']), row['Label'], row.get('Options', ''))
        grouped.setdefault(group, []).append(field)
    return grouped

field_structure = load_field_structure()
template_path = template_paths[project_count]
column_map = project_columns[project_count]

# === Utility Functions ===
def calculate_amount_due(inputs, proj, show_debug=False):
    def get(row):
        val = str(inputs.get(f"{row}_P{proj}", "0")).replace(",", "").replace("%", "").strip().lower()
        try:
            return float(val)
        except ValueError:
            return 0.0

    contract_sum = get("10")
    advance_payment_pct = get("12") / 100
    work_completed = get("13")
    retention_pct = get("14") / 100
    previous_payment = get("15")
    advance_refund_pct = get("16") / 100
    vat_pct = get("17") / 100

    advance_payment = contract_sum * advance_payment_pct
    retention = work_completed * retention_pct
    total_net_payment = work_completed - retention
    vat = total_net_payment * vat_pct
    total_net_amount = total_net_payment + vat
    advance_refund_amount = advance_refund_pct * advance_payment
    amount_due = total_net_amount - advance_refund_amount - previous_payment

    if show_debug:
        st.markdown(f"### 🧮 Debug for Project {proj}")
        st.write(f"Contract Sum: ₦{contract_sum:,.2f}")
        st.write(f"Advance: ₦{advance_payment:,.2f}")
        st.write(f"Work Completed: ₦{work_completed:,.2f}")
        st.write(f"Retention: ₦{retention:,.2f}")
        st.write(f"VAT: ₦{vat:,.2f}")
        st.write(f"Refund: ₦{advance_refund_amount:,.2f}")
        st.write(f"Previous Payment: ₦{previous_payment:,.2f}")

    return amount_due

def amount_in_words_naira(amount):
    try:
        naira = int(float(amount))
        kobo = int(round((float(amount) - naira) * 100))
        words = f"{num2words(naira, lang='en').capitalize()} naira"
        if kobo > 0:
            words += f", {num2words(kobo, lang='en')} kobo"
        return words.replace("-", " ")
    except:
        return "Invalid amount"

def write_to_details(ws, data_dict, column_map):
    currency_rows = {"10", "11", "13", "15", "18"}
    for proj, entries in data_dict.items():
        col = column_map[proj]
        for row_idx, value in entries.items():
            cell = ws[f"{col}{int(row_idx)}"]
            if str(row_idx) in currency_rows:
                try:
                    val = str(value).replace("\u20a6", "").replace(",", "").strip()
                    cell.value = float(val) if "." in val else int(val)
                    cell.number_format = '"\u20a6"#,##0.00'
                except:
                    cell.value = value
            else:
                cell.value = value
                
def save_data_locally(inputs, filename=None):
    inputs = dict(inputs)
    df = pd.DataFrame([inputs])
    df.to_csv("saved_form_data.csv", index=False)

    # Only name new backups if filename is not provided
    if filename:
        df.to_csv(os.path.join(user_backup_dir, filename), index=False)
    else:
        contractor = str(inputs.get("5_P1", "")).strip()
        project = str(inputs.get("7_P1", "")).strip()

        # Use 'unspecified' only if values are truly empty
        contractor_clean = re.sub(r'[^\w\-]', '_', contractor) if contractor else "unspecified_contractor"
        project_clean = re.sub(r'[^\w\-]', '_', project) if project else "unspecified_project"

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"{contractor_clean}_{project_clean}_{timestamp}.csv"
        df.to_csv(os.path.join(user_backup_dir, filename), index=False)
        
# === Restore from Backup or Load Fresh ===
if "restored_inputs" in st.session_state:
    restored = st.session_state.pop("restored_inputs")
    for k, v in restored.items():
        if pd.isna(v):
            v = ""
        else:
            v = str(v)
        st.session_state[k] = v
    all_inputs = restored.copy()
else:
    # If fields already in session_state from auto-save or fresh form
    all_inputs = {k: v for k, v in st.session_state.items() if "_P" in k}

# === Form Entry ===
for group, fields in field_structure.items():
    with st.expander(group, expanded=False):
        for row, label, _ in fields:
            for proj in range(1, project_count + 1):
                if proj > 1 and group in ["Date of Approval", "Address Line", "Signatories"]:
                    continue
                if proj > 1 and group == "Folio References" and label != "Inspection report File number":
                    continue
                    
                key = f"{row}_P{proj}"
                default = all_inputs.get(key, "")
                show_label = label if (proj == 1 or project_count == 1) else f"{label} – Project {proj}"

                if row == "18":
                    amount = calculate_amount_due(all_inputs, proj, show_debug=True)
                    amount_words = amount_in_words_naira(amount)
                    all_inputs[key] = f"{amount:,.2f}"
                    all_inputs[f"19_P{proj}"] = amount_words
                    st.session_state[f"19_P{proj}"] = amount_words
                    st.markdown(f"**{show_label}: ₦{amount:,.2f}**")
                    st.caption(f"In Words: {amount_words}")
                    continue

                elif row == "19":
                    value = all_inputs.get(key, "")
                    continue

                elif label in custom_dropdowns:
                    options = custom_dropdowns[label]
                    if key not in st.session_state or not isinstance(st.session_state[key], str):
                        st.session_state[key] = default if default in options else options[0]
                    all_inputs[key] = st.selectbox(show_label, options, key=key)

                else:
                    if key not in st.session_state or not isinstance(st.session_state[key], str):
                        st.session_state[key] = default if isinstance(default, str) else ""
                    all_inputs[key] = st.text_input(show_label, key=key)
                    
# === Save and Download Buttons ===
contractor = str(all_inputs.get("5_P1", "")).strip()
project = str(all_inputs.get("7_P1", "")).strip()
filename = st.session_state.get("loaded_filename")

if st.button("💾 Save Offline"):
    inputs_to_save = {k: v for k, v in st.session_state.items() if "_P" in k}
    save_data_locally(inputs_to_save, filename)  # existing local save
    save_backup_to_gsheet(st.session_state["current_user"], inputs_to_save)  # new!
    st.success("Form saved offline and backed up to Google Sheet.")
     
if st.button("📥 Download Excel"):
    wb = load_template(project_count)
    ws = wb[details_sheet]
    data_to_write = {p: {} for p in range(1, project_count + 1)}
    for key, value in st.session_state.items():
        if "_P" in key:
            parts = key.split("_P")
            if len(parts) == 2 and parts[1].isdigit():
                row, proj = parts
                data_to_write[int(proj)][row] = value
    write_to_details(ws, data_to_write, column_map)
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    st.download_button(
        "📂 Download Filled Excel",
        buffer,
        file_name=f"{project}_by_{contractor}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
# === Autosave to Google Sheet if form changes ===
current_inputs = {k: v for k, v in st.session_state.items() if "_P" in k}

# Detect changes since last autosave
if st.session_state.get("last_autosaved") != current_inputs:
    with st.spinner("Autosaving..."):
        save_backup_to_gsheet(st.session_state["current_user"], current_inputs)
        st.session_state["last_autosaved"] = current_inputs.copy()
        st.toast("Autosaved to Google Sheet ✅")

# === Backup Listing ===
st.sidebar.markdown("### 🔄 Manage Local Saved Backups - Temporary")
search_term = st.sidebar.text_input("🔍 Search Backups")

# Show all or per-user backups
all_user_dirs = [user_backup_dir] if not is_admin else [
    os.path.join(backup_root, u) for u in os.listdir(backup_root)
    if os.path.isdir(os.path.join(backup_root, u))
]
all_backups = []
for user_dir in all_user_dirs:
    for f in os.listdir(user_dir):
        if f.endswith(".csv"):
            path = os.path.join(user_dir, f)
            title = f"{os.path.basename(user_dir)} | {f.replace('.csv', '').replace('_', ' ')}"
            all_backups.append((path, title))

# Filter
filtered_backups = [b for b in all_backups if search_term.lower() in b[1].lower()]

for i, (path, title) in enumerate(filtered_backups):
    with st.sidebar.expander(title, expanded=False):
        col1, col2 = st.columns([2, 1])
        with col1:
            if st.button("Load", key=f"load_{i}"):
                try:
                    df = pd.read_csv(path)
                    if df.empty:
                        st.warning("Backup is empty.")
                    else:
                        st.session_state["restored_inputs"] = df.to_dict(orient="records")[0]
                        st.session_state["loaded_filename"] = os.path.basename(path)
                        st.success("Backup loaded successfully.")
                        st.rerun()
                except Exception as e:
                    st.error(f"Failed to load backup: {e}")
        with col2:
            if st.button("🗑️", key=f"delete_{i}"):
                os.remove(path)
                st.rerun()

# Reset form
if st.sidebar.button("➕ Start New Blank Form"):
    for k in list(st.session_state.keys()):
        if "_P" in k or "loaded_filename" in k or "_proj" in k:
            del st.session_state[k]
    st.rerun()

# === Restore from Google Sheets Backup ===
st.sidebar.markdown("### 🗃️ Restore from Google Sheet Backup")

try:
    gc = get_gsheet_client()
    if gc:
        sheet = gc.open("PC_Backups").sheet1
        rows = sheet.get_all_records()
        df_backups = pd.DataFrame(rows)

        current_user = st.session_state.get("current_user", "").strip()
        user_backups = df_backups[df_backups["user"].str.lower() == current_user.lower()]

        if user_backups.empty:
            st.sidebar.info("No backups found for your name.")
        else:
            timestamps = user_backups["timestamp"].unique().tolist()
            selected_timestamp = st.sidebar.selectbox("Select a Backup Time", timestamps)

            col1, col2 = st.sidebar.columns(2)

            with col1:
                if st.button("🔄 Load Backup"):
                    selected_rows = user_backups[user_backups["timestamp"] == selected_timestamp]
                    backup_dict = {row["field_key"]: row["value"] for _, row in selected_rows.iterrows()}
                    st.session_state["restored_inputs"] = backup_dict
                    st.session_state["loaded_filename"] = f"restored_from_sheet_{selected_timestamp.replace(' ', '_')}.csv"
                    st.sidebar.success("Backup loaded successfully from Google Sheet.")
                    st.rerun()

            with col2:
                if st.button("🗑️ Delete Backup"):
                    success = delete_backup_from_gsheet(current_user, selected_timestamp)
                    if success:
                        st.sidebar.success("Backup deleted successfully.")
                        st.rerun()

except Exception as e:
    st.sidebar.error(f"Google Sheet restore failed: {e}")
    

# === Summary Dashboard ===
st.markdown("---")
st.header("📊 Summary of All Projects")
total_due = 0
summary_data = []
for proj in range(1, project_count + 1):
    try:
        val = all_inputs.get(f"18_P{proj}", "0").replace(",", "")
        amt = float(val)
        total_due += amt
        summary_data.append((f"Project {proj}", amt))
    except:
        continue
 
    if summary_data:
        df_summary = pd.DataFrame(summary_data, columns=["Project", "Amount Due"])
        st.dataframe(df_summary)
        st.subheader(f"**Total Amount Due: ₦{total_due:,.2f}**")
    else:
        st.info("No amount data available to summarize.")
