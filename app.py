
# Prepopulate calculated fields for all projects
for proj in range(1, project_count + 1):
    amt = calculate_amount_due(all_inputs, proj)
    all_inputs[f"18_P{proj}"] = f"{amt:,.2f}"
    all_inputs[f"19_P{proj}"] = amount_in_words_naira(amt)

contractor = all_inputs.get("5_P1", "Contractor")
project_name = all_inputs.get("7_P1", "Project")

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
