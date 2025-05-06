
import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
from openpyxl import load_workbook
from num2words import num2words

template_paths = {
    1: "PC Template-1.xlsx",
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
    "Payment stage:": ["Stage Payment", "Final Payment", "Retention"],
    "Percentage of Advance payment? (as specified in the award letter)": ["0%", "25%", "40%", "50%", "60%", "70%"],
    "Is there 5% retention?": ["0%", "5%"],
    "Vat": ["0%", "7.5%"],
    "Address line 1": ["The Director", "The Chairman", "The Permanent Secretary", "The Honourable Commissioner", "The Special Adviser"]
}

def save_data_locally(all_inputs):
    import re  # For safe filename cleaning
    df = pd.DataFrame([all_inputs])
    os.makedirs("backups", exist_ok=True)

    # Always save latest form
    df.to_csv("saved_form_data.csv", index=False)

    # Overwrite loaded file if one is being edited
    if "loaded_filename" in st.session_state:
        df.to_csv(os.path.join("backups", st.session_state["loaded_filename"]), index=False)
    else:
        # Safely get and clean contractor and project names
        contractor = str(all_inputs.get("7_P1", "") or "").strip()
        project = str(all_inputs.get("5_P1", "") or "").strip()

        # Replace invalid characters with underscore
        contractor = re.sub(r'[^\w\-]', '_', contractor) or "no_contractor"
        project = re.sub(r'[^\w\-]', '_', project) or "no_project"

        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"backups/{contractor.lower()}_{project.lower()}_{timestamp}.csv"
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
            ws[f"{col}{int(row_idx)}"] = value

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
    vat_pct = get(f"vat_P{proj}") / 100

    advance_payment = contract_sum * advance_payment_pct
    retention = work_completed * retention_pct
    total_net_payment = work_completed - retention
    vat = total_net_payment * vat_pct
    total_net_amount = total_net_payment + vat
    advance_refund_amount = advance_refund_pct * advance_payment
    amount_due = total_net_amount - advance_refund_amount - previous_payment

    if show_debug:
        st.markdown(f"### 💰 Summary for Project {proj}")
        st.write(f"Contract Sum: ₦{contract_sum:,.2f}")
        st.write(f"Advance Payment %: {advance_payment_pct*100}% → ₦{advance_payment:,.2f}")
        st.write(f"Work Completed: ₦{work_completed:,.2f}")
        st.write(f"Retention %: {retention_pct*100}% → ₦{retention:,.2f}")
        st.write(f"Total Net Payment: ₦{total_net_payment:,.2f}")
        st.write(f"VAT %: {vat_pct*100}% → ₦{vat:,.2f}")
        st.write(f"Total Net Amount: ₦{total_net_amount:,.2f}")
        st.write(f"Advance Refund %: {advance_refund_pct*100}% → ₦{advance_refund_amount:,.2f}")
        st.write(f"Previous Payment: ₦{previous_payment:,.2f}")
        st.write(f"Final Amount Due: ₦{amount_due:,.2f}")

    return amount_due

def amount_in_words_naira(amount):
    naira = int(amount)
    kobo = int(round((amount - naira) * 100))
    words = f"{num2words(naira, lang='en').capitalize()} naira"
    if kobo > 0:
        words += f", {num2words(kobo, lang='en')} kobo"
    return words.replace("-", " ")

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

        project = str(data.get("5_P1", "") or "").strip()
        contractor = str(data.get("7_P1", "") or "").strip()

        if not project or not contractor:
            raise ValueError("Missing key metadata.")

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

    except Exception as e:
        # fallback
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

                    # Ensure fallback defaults if critical fields are missing
                    project_count = int(selected_data.get("project_count", 1))
                    contractor = selected_data.get("7_P1", "no_contractor") or "no_contractor"
                    project = selected_data.get("5_P1", "no_project") or "no_project"

                    wb = load_template(project_count)
                    ws = wb[details_sheet]

                    project_data = {p: {} for p in range(1, project_count + 1)}
                    for key, value in selected_data.items():
                        if "_P" in key:
                            try:
                                row, proj = key.split("_P")
                                project_data[int(proj)][row] = value
                            except:
                                continue  # skip malformed keys

                    write_to_details(ws, project_data, project_columns[project_count])

                    excel_buffer = io.BytesIO()
                    wb.save(excel_buffer)
                    excel_buffer.seek(0)

                    file_label = f"{contractor}_{project}.xlsx".replace(" ", "_").lower()
                    st.download_button(
                        label="Download Excel",
                        data=excel_buffer,
                        file_name=file_label,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{i}"
                    )
                except Exception as e:
                    st.caption(f"Failed to generate Excel for download. Error: {e}")
            with col2:
                if st.button("🗑️", key=f"delete_{i}"):
                    os.remove(os.path.join("backups", f))
                    st.success(f"Deleted backup: {f}")
                    st.rerun()
else:
    st.sidebar.info("No backups found matching your filters.")

contractor = all_inputs.get("7_P1", "Contractor")
project_name = all_inputs.get("5_P1", "Project Description")

# Option to start a new form
if st.sidebar.button("Start New Blank Form"):
    st.session_state["restored_inputs"] = {}
    if "loaded_filename" in st.session_state:
        del st.session_state["loaded_filename"]
    st.rerun()

contractor = all_inputs.get("7_P1", "Contractor")
project_name = all_inputs.get("5_P1", "Project Description")

# === 🧩 FORM ENTRY BLOCK ===
for group, fields in field_structure.items():
    with st.expander(group, expanded=False):
        for row, label, _ in fields:
            for proj in range(1, project_count + 1):

                # 🛑 Skip sections for proj > 1
                if group in ["Date of Approval", "Address Line", "Signatories"] and proj > 1:
                    continue
                if group == "Folio References" and label != "Inspection report File number" and proj > 1:
                    continue

                key = f"{row}_P{proj}"
                label_suffix = f"{label} – Project {proj}" if project_count > 1 else label
                if group in ["Date of Approval", "Address Line", "Signatories", "Folio References"] and label != "Inspection report File number":
                    label_suffix = label

                if key not in st.session_state:
                    st.session_state[key] = str(all_inputs.get(key, ""))

                default = st.session_state.get(key, "")
                widget_key = f"{group}_{label}_{proj}_{row}"

                # Address Line 2 autofill
                if label == "Address line 2":
                    default_ministry = all_inputs.get(f"3_P{proj}", "")
                    st.session_state[key] = str(default_ministry)
                    all_inputs[key] = st.text_input(label_suffix, value=st.session_state[key], key=widget_key)

                # Dropdowns
                elif label in custom_dropdowns:
                    options = custom_dropdowns[label]
                    default_value = all_inputs.get(key, options[0])
                    if key not in st.session_state or st.session_state[key] not in options:
                        st.session_state[key] = default_value
                    dropdown_value = st.session_state[key]
                    if dropdown_value not in options:
                        dropdown_value = options[0]
                        st.session_state[key] = dropdown_value
                    all_inputs[key] = st.selectbox(
                        label_suffix, options, index=options.index(dropdown_value), key=widget_key
                    )

                # Row 19 – skip (auto-filled)
                elif row == "19":
                    continue

                # Default text input
                else:
                    all_inputs[key] = st.text_input(label_suffix, value=default, key=widget_key)

st.markdown("---")
st.header("💡 Amount Due Summary")

for proj in range(1, project_count + 1):
    amount = calculate_amount_due(all_inputs, proj, show_debug=True)
    amount_words = amount_in_words_naira(amount)
    all_inputs[f"18_P{proj}"] = f"{amount:,.2f}"
    all_inputs[f"19_P{proj}"] = amount_words

else:
    all_inputs[key] = st.text_input(label_suffix, value=default, key=widget_key)

# Capture some general values
contractor = all_inputs.get("7_P1", "Contractor")
project_name = all_inputs.get("5_P1", "Project title")

if st.button("Generate Excel"):
    wb = load_template(project_count)
    ws = wb[details_sheet]
    project_data = {p: {} for p in range(1, project_count + 1)}
    for key, value in all_inputs.items():
        if "_P" in key:
            row, proj = key.split("_P")
            project_data[int(proj)][row] = value
    write_to_details(ws, project_data, column_map)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    st.success("Excel file is ready.")
    st.download_button(
        label="Download Filled Excel",
        data=buffer,
        file_name=f"{project_name}_by_{contractor}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
