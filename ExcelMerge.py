import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Sheet Merger", layout="wide")

st.title("📊 Excel Sheet Merger Tool")

st.write("""
Upload multiple Excel files and choose how to merge:
- **Single sheet:** Combine all selected sheets into one.
- **Multiple sheets:** Keep each sheet separate (original names preserved).
""")

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

    st.subheader("Merge Options")
    merge_mode = st.radio(
        "Choose how to merge your selected sheets:",
        ("Single Sheet (Combine All)", "Multiple Sheets (Keep Original Names)")
    )

    output_name = st.text_input(
        "Enter output file name (without extension):",
        value="merged_output"
    )

    if st.button("🔄 Merge and Download"):
        output = BytesIO()

        # ✅ Create writer outside 'with' so we can manually close after writing
        writer = pd.ExcelWriter(output, engine="openpyxl")

        try:
            if merge_mode == "Single Sheet (Combine All)":
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
                    result_df.to_excel(writer, index=False, sheet_name="Merged_Data")
                    st.success(f"Merged {len(merged_data)} sheets into one sheet successfully!")
                    st.dataframe(result_df.head(50))
                else:
                    st.warning("No valid sheets to merge!")

            else:  # Multiple sheets mode
                for file, sheets in selected_sheets.items():
                    for sheet in sheets:
                        try:
                            df = pd.read_excel(file, sheet_name=sheet)
                            sheet_name = f"{file.name[:20]}_{sheet}"[:31]
                            df.to_excel(writer, index=False, sheet_name=sheet_name)
                        except Exception as e:
                            st.warning(f"Skipping {file.name} - {sheet}: {e}")

                st.success("All selected sheets were added as separate sheets in the output file.")
        finally:
            writer.close()  # ✅ Explicit close to finalize the workbook

        # ✅ Move buffer to start before download
        output.seek(0)

        st.download_button(
            label="📥 Download Excel File",
            data=output,
            file_name=f"{output_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
