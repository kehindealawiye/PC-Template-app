
import streamlit as st
import pandas as pd
import io
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

def calculate_amount_due(inputs, proj):
    def get(row):
        val = str(inputs.get(f"{row}_P{proj}", "0")).replace(",", "").replace("%", "").strip().lower()
        return 0.0 if val in ["", "nil"] else float(val)

    contract_sum = get("9")
    advance_payment_pct = get("11") / 100
    work_completed = get("12")
    retention_pct = get("13") / 100
    previous_payment = get("14")
    advance_refund_pct = get("15") / 100
    vat_label = "Vat"
    vat_val = str(inputs.get(f"{vat_label}_P{proj}", "0")).replace("%", "").strip().lower()
    vat_pct = float(vat_val) / 100 if vat_val not in ["", "nil"] else 0.0


    advance_payment = contract_sum * advance_payment_pct
    retention = work_completed * retention_pct
    total_net_payment = work_completed - retention
    vat = total_net_payment * vat_pct
    total_net_amount = total_net_payment + vat
    advance_refund_amount = advance_refund_pct * advance_payment
    amount_due = total_net_amount - advance_refund_amount - previous_payment

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
all_inputs = {}

for group, fields in field_structure.items():
    with st.expander(group, expanded=True):
        for row, label, _ in fields:
            for proj in range(1, project_count + 1):
                key = f"{row}_P{proj}"
                label_suffix = f"{label} – Project {proj}" if project_count > 1 else label
                if label == "Address line 2":
                    client_ministry = all_inputs.get(f"3_P{proj}", "")
                    all_inputs[key] = st.text_input(label_suffix, value=client_ministry, key=key)
                elif label in custom_dropdowns:
                    all_inputs[key] = st.selectbox(label_suffix, custom_dropdowns[label], key=key)
                elif row == "18":
                    amount = calculate_amount_due(all_inputs, proj)
                    all_inputs[key] = f"{amount:,.2f}"
                    st.info(f"Calculated Amount Due: ₦{all_inputs[key]}")
                elif row == "19":
                    amount = calculate_amount_due(all_inputs, proj)
                    all_inputs[key] = amount_in_words_naira(amount)
                    st.write(f"Amount in Words: {all_inputs[key]}")
                else:
                    all_inputs[key] = st.text_input(label_suffix, key=key)

for proj in range(1, project_count + 1):

contractor = all_inputs.get("4_P1", "Contractor")
project_name = all_inputs.get("1_P1", "FilledTemplate")

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
