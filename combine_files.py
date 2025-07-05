import os
import pandas as pd

# --- ตั้งค่า ---
# ระบุโฟลเดอร์หลักของโปรเจกต์
root_folder = '.'  # '.' หมายถึงโฟลเดอร์ปัจจุบันที่สคริปต์นี้อยู่
# ระบุชื่อไฟล์ผลลัพธ์
output_filename = 'AI_Context_File.txt'
# ระบุนามสกุลไฟล์ที่ต้องการข้าม (ถ้ามี)
excluded_extensions = ['.tmp', '.pyc', '.git', '.idea']
# ----------------

def read_text_file(file_path):
    """อ่านไฟล์ข้อความธรรมดา"""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    except Exception as e:
        return f"--- Error reading {file_path}: {e} ---\n"

def read_excel_file(file_path):
    """อ่านไฟล์ Excel ทุกชีท"""
    try:
        xls = pd.ExcelFile(file_path)
        content = ""
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            content += f"--- Content from Excel sheet: '{sheet_name}' ---\n"
            content += df.to_string()
            content += "\n\n"
        return content
    except Exception as e:
        return f"--- Error reading Excel {file_path}: {e} ---\n"

# เริ่มกระบวนการ
with open(output_filename, 'w', encoding='utf-8') as outfile:
    print(f"Starting to combine files into {output_filename}...")
    for subdir, _, files in os.walk(root_folder):
        for file in files:
            file_path = os.path.join(subdir, file)
            file_extension = os.path.splitext(file)[1].lower()

            # ข้ามไฟล์ที่ไม่ต้องการ
            if file_extension in excluded_extensions or file == 'combine_files.py' or file == output_filename:
                continue

            print(f"Processing: {file_path}")

            outfile.write(f"--- START OF FILE: {file_path} ---\n\n")

            content = ""
            if file_extension in ['.xlsx', '.xls']:
                content = read_excel_file(file_path)
            else: # สำหรับไฟล์อื่นๆ ทั้งหมด ให้ถือว่าเป็น text
                content = read_text_file(file_path)

            outfile.write(content)
            outfile.write(f"\n\n--- END OF FILE: {file_path} ---\n\n")

print("Done! All files have been combined.")