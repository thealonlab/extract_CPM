import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import zipfile
import io
import os
import re

def clean_and_extract_lines(content, prefix_to_strip, remove_line, replacements):
    for binary, replacement in replacements.items():
        content = content.replace(binary.decode("utf-8", errors="ignore"), replacement)

    content = re.sub(r"[^\x20-\x7E\n]+", "", content)

    processed_lines, data_lines = [], []
    for line in content.splitlines():
        if remove_line in line or "ELAPSED" in line or "NO         MIN" in line or "PAGE" in line:
            continue

        line = line.lstrip(prefix_to_strip)

        if not re.match(r"^\s*\d+", line):
            processed_lines.append(line)
            continue

        fields = line.strip().split()
        if len(fields) >= 5:
            try:
                sample_number = int(fields[0])
                sample_id = fields[1]
                time_min = fields[2]
                h_val = fields[3]
                cpm_value = int(round(float(fields[4])))

                if re.match(r"^(\*\*-\d+|\d+-\d+)$", sample_id):
                    data_lines.append(f"{sample_number:<8} {sample_id:<10} {time_min:<12} {h_val:<10} {fields[4]:<12}")
            except ValueError:
                continue

    headers = f"{'Sample':<8} {'Position':<10} {'Time (min)':<12} {'H#':<10} {'CPM (3H)':<12}"
    return "\n".join(processed_lines + [headers] + data_lines)

def generate_excel_from_clean_file(clean_file_path, output_excel_path):
    rows = []
    row_labels = []
    col_labels = list(range(1, 19))
    letters = [chr(65 + i) for i in range(8)]

    with open(clean_file_path, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f if line.strip() and not line.startswith("Sample")]

    for line in lines:
        fields = line.split()
        try:
            sample_number = int(fields[0])
            cpm_value = int(round(float(fields[4])))
        except (ValueError, IndexError):
            continue

        row_index = (sample_number - 1) // 18
        col_index = (sample_number - 1) % 18

        while len(rows) <= row_index:
            rows.append([None] * 18)
            row_labels.append(letters[len(row_labels) % 8])

        rows[row_index][col_index] = cpm_value

    df = pd.DataFrame(rows, columns=col_labels)
    df.insert(0, "", row_labels)
    df.to_excel(output_excel_path, index=False)

    workbook = load_workbook(output_excel_path)
    sheet = workbook.active
    for cell in sheet["A"]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    workbook.save(output_excel_path)

# Streamlit UI
st.title("Extract CPM values from LS6500 output file")

uploaded_file = st.file_uploader("Upload a RECORD.TXT file", type=["txt", "TXT"])

if uploaded_file:
    content_bytes = uploaded_file.read()
    original_content = content_bytes.decode("utf-8", errors="ignore")

    now = datetime.now().strftime("%Y%m%d")
    clean_file_name = os.path.join("/tmp", f"{now}_RECORD_clean.txt")
    excel_file_name = os.path.join("/tmp", f"{now}_RESULTS.xlsx")
    original_file_name = f"{now}_RECORD.txt"
    zip_file_name = f"{now}_RESULTS.zip"

    search_strings = [
        b"\x1bx \x1bt \x1b7\x1bR \x1b2\x12\x1bP\x1bQP\x1bW \x1bH\x1b- \x1bx \x1bR \x1b2\x12\x1bP\x1bQP\x0c",
        b"\x1bx\x00\x1bt\x00\x1b7\x1bR\x00\x1b2\x12\x1bP\x1bQP\x1bW\x00\x1bH\x1b-\x00\x1bx\x00\x1bR\x00\x1b2\x12\x1bP\x1bQP\x0c"
    ]

    remove_line = "  \x1bG  MISSING SAMPLE\x1bH"
    prefix_to_strip = "\x1bH"
    replacements = {
        b"\x1bG": "", b"\x1bH": "", b"\x1bW": "", b"\x1b-": "",
        b"\x1b ": "", b"\x00 ": "", b"\x01": ""
    }

    last_index, selected_string = -1, None
    for s in search_strings:
        index = content_bytes.rfind(s)
        if index > last_index:
            last_index = index
            selected_string = s

    if selected_string is None:
        st.error("None of the expected search strings were found in the file.")
    else:
        start_index = last_index + len(selected_string)
        trimmed = content_bytes[start_index:].decode("utf-8", errors="ignore")

        filtered_content = clean_and_extract_lines(trimmed, prefix_to_strip, remove_line, replacements)
        with open(clean_file_name, "w", encoding="utf-8") as f:
            f.write(filtered_content)

        generate_excel_from_clean_file(clean_file_name, excel_file_name)

        st.download_button("Download Cleaned Text File", filtered_content, clean_file_name, mime="text/plain")
        st.download_button("Download Results Excel File", open(excel_file_name, "rb").read(), os.path.basename(excel_file_name), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("Download Original File", original_content, original_file_name, mime="text/plain")

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            zipf.writestr(original_file_name, uploaded_file.getvalue())
            zipf.write(clean_file_name, arcname=os.path.basename(clean_file_name))
            zipf.write(excel_file_name, arcname=os.path.basename(excel_file_name))
        zip_buffer.seek(0)

        st.download_button("Download All Outputs (ZIP)", zip_buffer, file_name=zip_file_name, mime="application/zip")
