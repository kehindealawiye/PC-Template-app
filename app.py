
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

def write_to_details(ws, data_dict, column_map, project_count):
    for proj, entries in data_dict.items():
        col = column_map[proj]

        # Filter out rows >= 35 for 2 or 3 projects
        filtered_entries = {}
        for row_idx, value in entries.items():
            try:
                row_num = int(row_idx)
                if project_count in [2, 3] and row_num >= 35:
                    continue
                filtered_entries[row_idx] = value
            except:
                continue  # skip if row index is not a number

        # Now write only filtered entries
        for row_idx, value in filtered_entries.items():
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
                    all_inputs[key] = st.selectbox(label_suffix, custom_dropdowns[label], key=key)

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
                    all_inputs[key] = st.text_input(label_suffix, key=key)

contractor = all_inputs.get("5_P1", "Contractor")
project_name = all_inputs.get("7_P1", "Project")

if st.button("Generate Excel"):
    wb = load_template(project_count)
    ws = wb[details_sheet]

    project_data = {p: {} for p in range(1, project_count + 1)}

    for key, value in all_inputs.items():
        if "_P" in key:
            row, proj = key.split("_P")
            proj = int(proj)
            row_num = int(row)

            # 🚫 Skip row 35 and downward for 2 or 3 projects
            if project_count in [2, 3] and row_num >= 46:
                continue

            project_data[proj][row] = value

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
