
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
    "Payment stage:": ["Stage Payment", "Final Payment", "Retention"],
    "Percentage of Advance payment? (as specified in the award letter)": ["0%", "25%", "40%", "50%", "60%", "70%"],
    "Is there 5% retention?": ["0%", "5%"],
    "Vat": ["0%", "7.5%"],
    "Physical Stage of Work": ["Ongoing", "Completed"],
    "Address line 1": ["The Director,", "The Chairman,", "The Permanent Secretary,", "The Honourable Commissioner,", "The Special Adviser,"]
}

def save_data_locally(all_inputs):
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"backups/form_backup_{timestamp}.csv"
    df = pd.DataFrame([all_inputs])

    # Ensure backups folder exists
    os.makedirs("backups", exist_ok=True)

    df.to_csv("saved_form_data.csv", index=False)  # still save latest form for auto-recovery
    df.to_csv(filename, index=False)  # save timestamped backup inside backups/

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
    all_inputs = load_saved_data()

for k, v in all_inputs.items():
    if pd.isna(v):
        all_inputs[k] = ""

st.sidebar.subheader("Load a Saved Form")

if os.path.exists("backups"):
    backup_files = sorted(
        [f for f in os.listdir("backups") if f.startswith("form_backup_") and f.endswith(".csv")],
        reverse=True
    )
else:
    backup_files = []

# Build preview-friendly display names
backup_titles = []
for f in backup_files:
    try:
        data = pd.read_csv(os.path.join("backups", f)).to_dict(orient='records')[0]
        project = data.get("5_P1", "No Project Name")
        contractor = data.get("7_P1", "No Contractor")
        backup_titles.append(f"{project} | {contractor} ({f})")
    except:
        backup_titles.append(f"(Unreadable) {f}")

# Let user choose based on title
if backup_titles:
    selected_title = st.sidebar.selectbox("Select backup to load", backup_titles)
    selected_file = backup_files[backup_titles.index(selected_title)]

    if st.sidebar.button("Load Selected Backup"):
        try:
            selected_data = pd.read_csv(os.path.join("backups", selected_file)).to_dict(orient='records')[0]
            st.session_state["restored_inputs"] = selected_data
            st.success(f"Loaded backup: {selected_title}")
            st.rerun()
        except Exception as e:
            st.warning(f"Unable to load selected backup. Error: {e}")
else:
    st.sidebar.info("No backups found yet. Save your form to create a backup.")

for group, fields in field_structure.items():
    with st.expander(group, expanded=False):
        for row, label, _ in fields:
            for proj in range(1, project_count + 1):
                # Skip these sections for projects > 1
                if group in ["Date of Approval", "Address Line", "Signatories"] and proj > 1:
                    continue

                # Show only 'Inspection report File number' for projects 2 & 3
                if group == "Folio References":
                    if label != "Inspection report File number" and proj > 1:
                        continue

                key = f"{row}_P{proj}"

                # Control label formatting
                if group in ["Date of Approval", "Address Line", "Signatories", "Folio References"] and label != "Inspection report File number":
                    label_suffix = label
                else:
                    label_suffix = f"{label} – Project {proj}" if project_count > 1 else label

                # Handle different field types
                if label == "Address line 2":
                    client_ministry = all_inputs.get(f"3_P{proj}", "")
                    all_inputs[key] = st.text_input(label_suffix, value=client_ministry, key=key)

                elif label in custom_dropdowns:
                    options = custom_dropdowns[label]
                    default = all_inputs.get(key, options[0]) if key in all_inputs else options[0]
                    all_inputs[key] = st.selectbox(label_suffix, options, index=options.index(default), key=key)

                elif row == "18":
                    amount = calculate_amount_due(all_inputs, proj, show_debug=True)
                    all_inputs[key] = f"{amount:,.2f}"
                    amount_words = amount_in_words_naira(amount)
                    all_inputs[f"19_P{proj}"] = amount_words
                    st.info(f"Calculated Amount Due: ₦{all_inputs[key]}")
                    st.write(f"Amount in Words: {amount_words}")

                elif row == "19":
                    continue  # skip because already handled in row 18

                else:
                # Always set default value from saved or restored inputs
                    default = all_inputs.get(key, "")
                    all_inputs[key] = st.text_input(label_suffix, value=default, key=key)

contractor = all_inputs.get("7_P1", "Contractor")
project_name = all_inputs.get("5_P1", "Project Description")

if st.button("Save My Work Offline"):
    save_data_locally(all_inputs)
    st.success("Saved successfully with timestamp and recovery file.")

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
   