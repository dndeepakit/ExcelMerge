import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Sheet Merger", layout="wide")

st.title("📊 Excel Sheet Merger Tool")
st.write("Upload multiple Excel files and merge selected sheets into one combined Excel file.")

uploaded_files = st.file_uploader(
    "Upload your Excel files here",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    selected_sheets = {}
    st.subheader("Select Sheets to Merge")

    for uploaded_file in uploaded_files:
        try:
            xls = pd.ExcelFile(uploaded_file)
            sheets = xls.sheet_names
            selected = st.multiselect(
                f"Select sheet(s) from **{uploaded_file.name}**",
                sheets,
                default=sheets[0]
            )
            selected_sheets[uploaded_file] = selected
        except Exception as e:
            st.error(f"Error reading {uploaded_file.name}: {e}")

    if st.button("🔄 Merge Selected Sheets"):
        merged_data = []

        for file, sheets in selected_sheets.items():
            for sheet in sheets:
                try:
                    df = pd.read_excel(file, sheet_name=sheet)
                    df["Source_File"] = file.name
                    df["Source_Sheet"] = sheet
                    merged_data.append(df)
                except Exception as e:
                    st.warning(f"Skipping {file.name} - {sheet}: {e}")

        if merged_data:
            result_df = pd.concat(merged_data, ignore_index=True)
            st.success(f"Merged {len(merged_data)} sheets successfully!")
            st.dataframe(result_df.head(50))

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="Merged_Data")

            st.download_button(
                label="📥 Download Merged Excel File",
                data=output.getvalue(),
                file_name="merged_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
