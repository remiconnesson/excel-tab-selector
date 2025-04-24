import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import os

st.set_page_config(
    page_title="Excel Sheet Manager",
    page_icon="📊",
    layout="centered"
)

st.title("Excel Sheet Manager")
st.write("Upload an Excel file, select which sheets to keep, and download a new file with only those sheets.")

# Function to get sheet names from an Excel file
def get_sheet_names(file):
    try:
        xls = pd.ExcelFile(file)
        return xls.sheet_names
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return []

# Function to create a new Excel file with only the selected sheets
def process_excel(input_file, selected_sheets):
    try:
        # Load the workbook with openpyxl to preserve formatting
        workbook = openpyxl.load_workbook(input_file)
        
        # Get all sheet names in the workbook
        all_sheets = workbook.sheetnames
        
        # Remove sheets that are not selected
        sheets_to_remove = [sheet for sheet in all_sheets if sheet not in selected_sheets]
        
        for sheet_name in sheets_to_remove:
            del workbook[sheet_name]
        
        # Save to BytesIO object
        output = BytesIO()
        workbook.save(output)
        output.seek(0)
        
        return output
    except Exception as e:
        st.error(f"Error processing Excel file: {e}")
        return None

# File uploader widget
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Display file info
    file_details = {"Filename": uploaded_file.name, "File size": f"{uploaded_file.size / 1024:.2f} KB"}
    st.write(file_details)
    
    # Get sheet names
    sheet_names = get_sheet_names(uploaded_file)
    
    if sheet_names:
        st.write(f"Found {len(sheet_names)} sheets in the Excel file.")
        
        # Create checkboxes for each sheet (all checked by default)
        selected_sheets = []
        
        # Using columns to display checkboxes in rows of 3
        cols = st.columns(3)
        
        for i, sheet_name in enumerate(sheet_names):
            col_idx = i % 3
            with cols[col_idx]:
                if st.checkbox(sheet_name, value=True, key=f"sheet_{i}"):
                    selected_sheets.append(sheet_name)
        
        # Input field for output filename
        default_filename = f"copy_{uploaded_file.name}"
        output_filename = st.text_input("Output filename (with .xlsx extension)", value=default_filename)
        
        if not output_filename.endswith(('.xlsx', '.xls')):
            st.warning("Please ensure the filename ends with .xlsx or .xls")
            output_filename = f"{output_filename}.xlsx"
        
        # Download button
        if st.button("Process and Download"):
            if not selected_sheets:
                st.error("Please select at least one sheet to include.")
            else:
                # Reset file pointer to beginning
                uploaded_file.seek(0)
                
                with st.spinner("Processing Excel file..."):
                    # Process the file to create a new Excel with only selected sheets
                    output = process_excel(uploaded_file, selected_sheets)
                    
                    if output:
                        # Create download button
                        st.download_button(
                            label="📥 Download Processed Excel",
                            data=output,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.success(f"Successfully created Excel file with {len(selected_sheets)} selected sheets!")
    else:
        st.error("No sheets found in the uploaded Excel file or the file format is not supported.")
else:
    st.info("👆 Upload an Excel file to begin.")

# Footer
st.markdown("---")
st.caption("This app allows you to selectively keep sheets from an Excel file while preserving all formatting.")
