import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import zipfile
import io
import os
import re

def generate_excel(data_lines, output_excel_path):
    """
    Generate an Excel sheet with CPM values arranged in a table with 18 columns.
    """
    rows = []
    row_labels = []
    col_labels = [i for i in range(1, 19)]
    letters = {0: "A", 1: "B", 2: "C", 3: "D", 4: "E", 5: "F", 6: "G", 7: "H"}

    for line in data_lines:
        fields = line.split()
        # Skip lines with invalid or unexpected data in the CPM field
        try:
            sample_number = int(fields[0])  # First column is the sample number
            cpm_value = int(round(float(fields[4])))  # Fifth column is the CPM value
        except (ValueError, IndexError):
        # Skip this line if there's a conversion error or missing data
            continue

        sample_number = int(fields[0])
        cpm_value = int(round(float(fields[4])))

        row_index = (sample_number - 1) // 18
        column_index = (sample_number - 1) % 18

        while len(rows) <= row_index:
            rows.append([None] * 18)
            row_labels.append(letters[len(row_labels) % 8])

        rows[row_index][column_index] = cpm_value

    df = pd.DataFrame(rows, columns=col_labels)
    df.insert(0, "", row_labels)
    df.to_excel(output_excel_path, index=False)

    workbook = load_workbook(output_excel_path)
    sheet = workbook.active

    for cell in sheet["A"]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    workbook.save(output_excel_path)

def process_lines(content, prefix_to_strip, remove_line, replacements=None, line_number_threshold=15):
    processed_lines = []
    data_lines = []
    line_counter = 0

    for line in content.splitlines():
        line_counter += 1

        if remove_line in line or "ELAPSED" in line or "NO         MIN" in line or "PAGE" in line:
            continue

        line = line.lstrip(prefix_to_strip)

        if replacements:
            for binary, replacement in replacements.items():
                line = line.replace(binary.decode("utf-8", errors="ignore"), replacement)

        if line_counter > 2:
            if line_counter > line_number_threshold and not line.strip():
                continue
            if line_counter < line_number_threshold:
                processed_lines.append(line)
            else:
                data_lines.append(line)

    formatted_data_lines = []
    for line in data_lines:
        fields = line.split()
        formatted_line = f"{fields[0]:<8} {re.sub(r"..-", "", fields[1]):<10} {fields[2]:<12} {fields[3]:<10} {fields[4]:<12} {fields[5]:<12} {fields[6]:<12} {fields[7]:<12}"
        formatted_data_lines.append(formatted_line)

    headers = f"{'Sample':<8} {'Position':<10} {'Time (min)':<12} {'H#':<10} {'CPM (3H)':<12} {'%Error (3H)':<12} {'%Lumex':<12} {'Elapsed time':<12}"
    formatted_data_lines.insert(0, headers)

    return "\n".join(processed_lines + formatted_data_lines)

st.title("Extract CPM values from LS6500 output file")

uploaded_file = st.file_uploader("Upload a RECORD.TXT file", type=["txt"])

if uploaded_file is not None:
    content_bytes = uploaded_file.read()
    original_content = content_bytes.decode("utf-8", errors="ignore")

    # File naming
    current_date = datetime.now().strftime("%Y%m%d")
    clean_file_name = f"{current_date}_RECORD_clean.txt"
    excel_file_name = f"{current_date}_RESULTS.xlsx"
    original_file_name = f"{current_date}_RECORD.txt"
    zip_file_name = f"{current_date}_RESULTS.zip"

    search_string_bytes = b"\x1bx\x00\x1bt\x00\x1b7\x1bR\x00\x1b2\x12\x1bP\x1bQP\x1bW\x00\x1bH\x1b-\x00\x1bx\x00\x1bR\x00\x1b2\x12\x1bP\x1bQP\x0c"
    remove_line = "  \x1bG  MISSING SAMPLE\x1bH"
    prefix_to_strip = "\x1bH"
    replacements = {
        b"\x1bG": "",
        b"\x1bH": "",
        b"\x1bW": "",
        b"\x1b-": "",
        b"\x1b ": "",
        b"\x00 ": "",
        b"\x01": ""
    }

    last_occurrence_index = content_bytes.rfind(search_string_bytes)

    if last_occurrence_index != -1:
        start_index = last_occurrence_index + len(search_string_bytes)
        trimmed_content_bytes = content_bytes[start_index:]
        trimmed_content = trimmed_content_bytes.decode("utf-8", errors="ignore")

        filtered_content = process_lines(
            content=trimmed_content,
            prefix_to_strip=prefix_to_strip,
            remove_line=remove_line,
            replacements=replacements
        )

        data_lines = filtered_content.splitlines()[13:]

        # Save outputs
        with open(clean_file_name, "w", encoding="utf-8") as clean_file:
            clean_file.write(filtered_content)

        generate_excel(data_lines, excel_file_name)

        # Download buttons
        st.download_button(
            label="Download Cleaned Text File",
            data=filtered_content,
            file_name=clean_file_name,
            mime="text/plain"
        )

        st.download_button(
            label="Download Results Excel File",
            data=open(excel_file_name, "rb").read(),
            file_name=excel_file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download Original File",
            data=original_content,
            file_name=original_file_name,
            mime="text/plain"
        )
        
        # Check if cleaned_output_path exists
        if not os.path.exists(clean_file_name):
            st.error(f"Cleaned file not found: {clean_file_name}")
        else:
            st.success(f"Cleaned file found: {clean_file_name}")
        
        # Check if excel_output_path exists
        if not os.path.exists(excel_file_name):
            st.error(f"Excel file not found: {excel_file_name}")
        else:
            st.success(f"Excel file found: {excel_file_name}")
        
        # Create a ZIP archive containing all outputs
        zip_buffer = io.BytesIO()  # In-memory buffer for the ZIP file
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            # Add original uploaded file to the ZIP
            zipf.writestr(f"{current_date}_RECORD.txt", uploaded_file.getvalue())
            
            # Add cleaned and Excel files only if they exist
            if os.path.exists(clean_file_name):
                zipf.write(clean_file_name, arcname=f"{current_date}_RECORD_clean.txt")
            if os.path.exists(excel_file_name):
                zipf.write(excel_file_name, arcname=f"{current_date}_RESULTS.xlsx")
        zip_buffer.seek(0)  # Rewind the buffer
        
        # Add a download button for the ZIP file
        st.download_button(
            label="Download All Outputs (ZIP)",
            data=zip_buffer,
            file_name=zip_file_name,
            mime="application/zip"
        )
#       # Create a ZIP archive containing all outputs
#       zip_buffer = io.BytesIO()  # In-memory buffer for the ZIP file
#       with zipfile.ZipFile(zip_buffer, "w") as zipf:
#           zipf.writestr(f"{current_date}_RECORD.txt", uploaded_file.getvalue())
#           zipf.write(cleaned_output_path, arcname=f"{current_date}_RECORD_clean.txt")
#           zipf.write(excel_output_path, arcname=f"{current_date}_RESULTS.xlsx")
#           zip_buffer.seek(0)  # Rewind the buffer

#       # Add a download button for the ZIP file
#       st.download_button(
#           label="Download All Outputs (ZIP)",
#           data=zip_buffer,
#           file_name=zip_file_name,
#           mime="application/zip"
#       )

    else:
        st.error("The specified string was not found in the file.")
