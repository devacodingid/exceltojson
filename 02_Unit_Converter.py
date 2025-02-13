import pandas as pd
import json
import io
import zipfile
import streamlit as st

def split_data_to_json(df):
    num_rows = len(df)
    row_count = 25
    file_counter = 1
    json_files = []

    for start_row in range(0, num_rows, row_count):
        end_row = min(start_row + row_count, num_rows)
        chunk_df = df.iloc[start_row:end_row]

        data_list = []
        for _, row in chunk_df.iterrows():
            data_list.append({
                "sl_no": int(row["sl_no"]),
                "units": int(row["units"])
            })

        json_str = json.dumps(data_list, indent=4)
        json_bytes = io.BytesIO(json_str.encode('utf-8'))

        json_files.append((f"Data_{file_counter}.json", json_bytes))

        file_counter += 1

    return json_files

def create_zip(json_files):
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zip_file:
        for file_name, json_bytes in json_files:
            zip_file.writestr(file_name, json_bytes.getvalue())

    zip_buffer.seek(0)
    return zip_buffer

def excel_splitter():
    st.title("Unit Converter")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)

        if st.button("Split Data and Download ZIP"):
            json_files = split_data_to_json(df)

            zip_buffer = create_zip(json_files)

            st.download_button(
                label="Download ZIP",
                data=zip_buffer,
                file_name="json_files.zip",
                mime="application/zip"
            )

excel_splitter()
