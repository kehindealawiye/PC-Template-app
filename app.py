
import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from docx import Document

# Template config
template_paths = {
    1: "PC Template.xlsx",
    2: "PC Template 2.xlsx",
    3: "PC Template 3.xlsx"
}

column_map = {
    1: 'B',
    2: 'E',
    3: 'H'
}

def load_labels():
    df = pd.read_excel(template_paths[1], sheet_name=0, header=None, usecols="A")
    return df[0].dropna().reset_index(drop=True)

def load_template(projects):
    return load_workbook(template_paths[projects])

def get_calculated_value(path, col_letter):
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    return ws[f"{col_letter}18"].value

def update_excel(inputs, projects):
    wb = load_template(projects)
    ws = wb.active
    col = column_map[projects]
    for idx, value in inputs.items():
        cell = f"{col}{idx + 1}"
        ws[cell] = value
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def export_to_word(sheet_names, template_path):
    doc = Document()
    for sheet in sheet_names:
        doc.add_heading(sheet, level=1)
        df_sheet = pd.read_excel(template_path, sheet_name=sheet, header=None)
        for row in df_sheet.itertuples(index=False):
            line = ' | '.join([str(cell) if pd.notna(cell) else '' for cell in row])
            doc.add_paragraph(line)
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def main():
    st.title("Multi-Project Excel Filler & Word Export")

    project_count = st.selectbox("How many projects are you working on?", [1, 2, 3])
    template_path = template_paths[project_count]
    col_letter = column_map[project_count]

    labels = load_labels()
    st.subheader("Fill in the form")

    all_inputs = {}
    for proj in range(1, project_count + 1):
        st.markdown(f"### Project {proj}")
        for i, label in enumerate(labels):
            custom_label = f"{label} – Project {proj}"
            key = f"{i}_P{proj}"
            all_inputs[key] = st.text_input(custom_label, key=key)

    contractor = all_inputs.get(f"3_P1", "Contractor")
    project_name = all_inputs.get(f"0_P1", "FilledTemplate")  # Using label 0 as project title

    if st.button("Generate Filled Excel"):
        # Flatten for target column
        excel_inputs = {}
        for i, label in enumerate(labels):
            key = f"{i}_P{project_count}"
            val = all_inputs.get(key, "")
            if val:
                excel_inputs[i] = val
        excel_file = update_excel(excel_inputs, project_count)
        st.success("Excel generated successfully")
        st.download_button(
            label="Download Excel File",
            data=excel_file,
            file_name=f"{project_name}_by_{contractor}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.subheader("Calculated Value from Template")
    calculated = get_calculated_value(template_path, col_letter)
    st.info(f"Calculated Value (Row 18): {calculated}")

    st.subheader("Export Sheets to Word")
    wb = load_template(project_count)
    sheet_names = wb.sheetnames
    selected_sheets = st.multiselect("Select sheets", sheet_names)

    if st.button("Generate Word Document") and selected_sheets:
        word_file = export_to_word(selected_sheets, template_path)
        st.download_button(
            label="Download Word File",
            data=word_file,
            file_name=f"{project_name}_by_{contractor}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

if __name__ == "__main__":
    main()
