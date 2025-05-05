import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from openpyxl import load_workbook
from num2words import num2words

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

custom_dropdowns = {
    "Payment stage:": ["Stage Payment", "Final Payment", "Retention Payment"],
    "Percentage of Advance payment? (as specified in the award letter)": ["0%", "25%", "40%", "50%", "60%", "70%"],
    "Is there 5% retention?": ["0%", "5%"],
    "Vat": ["0%", "7.5%"],
    "Physical Stage of Work": ["Ongoing", "Completed"],
    "Address line 1": ["The Director,", "The Chairman,", "The Permanent Secretary,", "The Honourable Commissioner,", "The Special Adviser,"]
}

def save_data_locally(all_inputs):
    df = pd.DataFrame([all_inputs])
    os.makedirs("backups", exist_ok=True)

    # Always save latest form
    df.to_csv("saved_form_data.csv", index=False)

    # Overwrite loaded file if one is being edited
    if "loaded_filename" in st.session_state:
        df.to_csv(os.path.join("backups", st.session_state["loaded_filename"]), index=False)
    else:
        contractor = all_inputs.get("7_P1", "NoContractor").replace(" ", "_").lower()
        project = all_inputs.get("5_P1", "NoProject").replace(" ", "_").lower()
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"backups/{contractor}_{project}_{timestamp}.csv"
        df.to_csv(filename, index=False)

def load_saved_data():
    try:
        return pd.read_csv("saved_form_data.csv").to_dict(orient='records')[0]
    except:
        return {}

def load_field_structure():
    df = pd.read_csv("Grouped_Field_Structure_Clean.csv")
    grouped = {}
    for _, row in df.iterrows():
        group = row['Group']
        field = (str(row['Row']), row['Label'], row.get('Options', ''))
        grouped.setdefault(group, []).append(field)
    return grouped

def load_template(project_count):
    return load_workbook(template_paths[project_count])

def write_to_details(ws, data_dict, column_map):
    for proj, entries in data_dict.items():
        col = column_map[proj]
        for row_idx, value in entries.items():
            try:
                val = float(str(value).replace(",", "").strip())
                ws[f"{col}{int(row_idx)}"] = int(val) if val.is_integer() else val
            except:
                ws[f"{col}{int(row_idx)}"] = value

def calculate_amount_due(inputs, proj, show_debug=False):
    def get(row):
        val = str(inputs.get(f"{row}_P{proj}", "0")).replace(",", "").replace("%", "").strip().lower()
        try:
            return 0.0 if val in ["", "nil"] else float(val)
        except:
            return 0.0

    contract_sum = get("10")
    revised_contract_sum = get("11")
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
        st.markdown(f"### Debug Info – Project {proj}")
        st.write(f"Contract Sum: ₦{contract_sum:,.2f}")
        st.write(f"Advance Payment %: {advance_payment_pct * 100}% → ₦{advance_payment:,.2f}")
        st.write(f"Work Completed: ₦{work_completed:,.2f}")
        st.write(f"Retention %: {retention_pct * 100}% → ₦{retention:,.2f}")
        st.write(f"Total Net Payment: ₦{total_net_payment:,.2f}")
        st.write(f"VAT %: {vat_pct * 100}% → ₦{vat:,.2f}")
        st.write(f"Total Net Amount: ₦{total_net_amount:,.2f}")
        st.write(f"Advance Refund %: {advance_refund_pct * 100}% → ₦{advance_refund_amount:,.2f}")
        st.write(f"Previous Payment: ₦{previous_payment:,.2f}")
        st.write(f"Final Amount Due: ₦{amount_due:,.2f}")

    return amount_due

def amount_in_words_naira(amount):
    try:
        naira = int(amount)
        kobo = int(round((amount - naira) * 100))
        words = f"{num2words(naira, lang='en').capitalize()} naira"
        if kobo > 0:
            words += f", {num2words(kobo, lang='en')} kobo"
        return words.replace("-", " ")
    except Exception as e:
        return f"Error: {e}"

st.set_page_config(page_title="Prepayment Form", layout="wide")
st.title("Prepayment Certificate Filler")

project_count = st.selectbox("Number of Projects", [1, 2, 3])
template_path = template_paths[project_count]
column_map = project_columns[project_count]
field_structure = load_field_structure()
if "restored_inputs" in st.session_state:
    all_inputs = st.session_state.pop("restored_inputs")
else:
    all_inputs = {}  # load a completely blank form by default

for k, v in all_inputs.items():
    if pd.isna(v):
        all_inputs[k] = ""

st.sidebar.subheader("Manage Saved Forms")

if os.path.exists("backups"):
    backup_files = sorted(
        [f for f in os.listdir("backups") if f.endswith(".csv")],
        reverse=True
    )
else:
    backup_files = []

# Preprocess all backups to extract metadata
backup_metadata = []
contractors = set()

for f in backup_files:
    try:
        data = pd.read_csv(os.path.join("backups", f)).to_dict(orient='records')[0]
        project = data.get("5_P1", "No Project").strip()
        contractor = data.get("7_P1", "No Contractor").strip()
        contractors.add(contractor)

        # Extract timestamp from filename
        parts = f.replace(".csv", "").split("_")
        if len(parts) >= 3:
            date_part = parts[-2]
            time_part = parts[-1].replace("-", ":")
            datetime_str = f"{date_part} {time_part}"
        else:
            datetime_str = "Unknown Time"

        title = f"{contractor} | {project} | {datetime_str}"
        backup_metadata.append((f, title, contractor.lower(), project.lower()))
    except:
        backup_metadata.append((f, f"(Unreadable) {f}", "", ""))

# Sidebar filters
selected_contractor = st.sidebar.selectbox("Filter by Contractor", ["All"] + sorted(contractors))
search_query = st.sidebar.text_input("Search Project or Contractor", "")

# Filter backups
filtered_files = []
for f, title, contractor_lower, project_lower in backup_metadata:
    if selected_contractor != "All" and contractor_lower != selected_contractor.lower():
        continue
    if search_query and search_query.lower() not in title.lower():
        continue
    filtered_files.append((f, title))

# Display backups
if filtered_files:
    for i, (f, title) in enumerate(filtered_files):
        with st.sidebar.expander(title, expanded=False):
            col1, col2 = st.columns([1.5, 1])
            with col1:
                if st.button(f"Load", key=f"load_{i}"):
                    try:
                        selected_data = pd.read_csv(os.path.join("backups", f)).to_dict(orient='records')[0]
                        st.session_state["restored_inputs"] = selected_data
                        st.session_state["loaded_filename"] = f
                        st.success(f"Loaded backup: {f}")
                        st.rerun()
                    except Exception as e:
                        st.warning(f"Unable to load selected backup. Error: {e}")

                try:
                    selected_data = pd.read_csv(os.path.join("backups", f)).to_dict(orient='records')[0]
                    project_count = int(selected_data.get("project_count", 1))
                    wb = load_template(project_count)
                    ws = wb[details_sheet]

                    project_data = {p: {} for p in range(1, project_count + 1)}
                    for key, value in selected_data.items():
                        if "_P" in key:
                            row, proj = key.split("_P")
                            project_data[int(proj)][row] = value
                    write_to_details(ws, project_data, project_columns[project_count])

                    excel_buffer = io.BytesIO()
                    wb.save(excel_buffer)
                    excel_buffer.seek(0)

                    contractor = selected_data.get("7_P1", "no_contractor")
                    project = selected_data.get("5_P1", "no_project")
                    file_label = f"{contractor}_{project}.xlsx".replace(" ", "_").lower()
                    st.download_button(
                        label="Download Excel",
                        data=excel_buffer,
                        file_name=file_label,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{i}"
                    )
