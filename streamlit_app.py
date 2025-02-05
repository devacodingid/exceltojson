import pandas as pd
import json
import zipfile
import io
import streamlit as st
import openpyxl
from io import BytesIO

# Function to read data from Excel and process it based on column and ranges
def process_sheet(excel_file, sheet_name, column_name, range_size, file_prefix):
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"Error reading the sheet: {e}")
        return []

    if column_name not in df.columns:
        st.error(f"Column '{column_name}' not found in sheet '{sheet_name}'.")
        return []

    df[column_name] = pd.to_numeric(df[column_name], errors='coerce')

    max_value = df[column_name].max()

    json_files = []

    for i in range(1, int(max_value) + 1, range_size):
        df_filtered = df[(df[column_name] >= i) & (df[column_name] < i + range_size)]

        if df_filtered.empty:
            st.warning(f"No data found for range {i} to {i + range_size - 1} in column '{column_name}'.")
            continue

        file_name = f'{file_prefix}{(i - 1) // range_size + 1}.json'

        df_filtered_dict = df_filtered.to_dict(orient='records')

        json_files.append({
            'file_name': file_name,
            'data': json.dumps(df_filtered_dict, indent=4)
        })

    return json_files

def create_zip(json_files):
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for json_file in json_files:
            zip_file.writestr(json_file['file_name'], json_file['data'])

    zip_buffer.seek(0)
    return zip_buffer

def main():
    st.title("Excel to JSON Converter")

    # Initialize session state
    if 'excel_file' not in st.session_state:
        st.session_state['excel_file'] = None
    if 'json_files' not in st.session_state:
        st.session_state['json_files'] = None

    # File uploader
    excel_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xlsm"])

    if excel_file:
        st.session_state['excel_file'] = excel_file

    # Convert button
    if st.button("Convert"):
        if st.session_state['excel_file']:
            sheet1_name = 'lineitem'
            column1_name = 'loop_count'
            sheet2_name = 'changeorder'
            column2_name = 'loop'

            range_size = 5

            json_files = []
            json_files.extend(process_sheet(st.session_state['excel_file'], sheet_name=sheet1_name, column_name=column1_name, 
                                            range_size=range_size, file_prefix='lineitem'))
            json_files.extend(process_sheet(st.session_state['excel_file'], sheet_name=sheet2_name, column_name=column2_name, 
                                            range_size=range_size, file_prefix='orderfile'))

            if json_files:
                st.session_state['json_files'] = json_files
                st.success("Conversion successful! Now you can download the zip file.")

            else:
                st.warning("No data to convert. Please check the uploaded file.")

        else:
            st.warning("Please upload an Excel file before converting.")

    # Download button
    if st.session_state['json_files']:
        zip_buffer = create_zip(st.session_state['json_files'])
        st.download_button(
            label="Download JSON as Zip",
            data=zip_buffer,
            file_name="json_files.zip",
            mime="application/zip"
        )

    # Clear button (Simulates Shift + F5)
    if st.button("Clear"):
        # Clear the session state and reset the entire app
        st.session_state.clear()  # This clears all session state
        st.rerun()  # This forces the script to rerun completely

if __name__ == "__main__":
    main()
