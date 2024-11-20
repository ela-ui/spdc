import streamlit as st
import zipfile
import pandas as pd
from io import BytesIO

# Streamlit app title
st.title("SPDC DATA CONSOLIDATION")

# Step 1: Upload ZIP file
uploaded_zip = st.file_uploader("Upload a ZIP file containing Excel files", type="zip")

if uploaded_zip:
    # Step 2: Read the ZIP file
    with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
        # Step 3: Get all Excel file names within the ZIP archive
        excel_files = [f for f in zip_ref.namelist() if f.endswith('.xlsx')]

        # Initialize an empty DataFrame for combining data
        combined_df = pd.DataFrame()

        # Step 4: Process each Excel file in the ZIP
        for excel_file in excel_files:
            with zip_ref.open(excel_file) as file:
                # Read all sheets from the Excel file
                excel_df = pd.read_excel(file, sheet_name=None)

                # Flatten sheets and concatenate them into a single DataFrame
                for sheet_name, sheet_df in excel_df.items():
                    # Concatenate the current sheet's data with the combined DataFrame
                    combined_df = pd.concat([combined_df, sheet_df], ignore_index=True)

        # Step 5: Save combined DataFrame to an Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Combined Data')

        # Step 6: Prepare the download of the combined file
        st.success("Excel files have been successfully combined!")
        st.download_button(
            label="Download Combined Excel File",
            data=output.getvalue(),
            file_name="combined_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
