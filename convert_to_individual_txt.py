import os
import sys
import pandas as pd

# --- การตั้งค่า ---
# 1. ให้โปรแกรมถามหาโฟลเดอร์ต้นทาง
source_folder = input("กรุณาวาง Path ของโฟลเดอร์ที่ต้องการแปลง: ")

# 2. ตั้งชื่อโฟลเดอร์สำหรับเก็บผลลัพธ์
output_folder = 'Converted_TXT_Files'

# ตรวจสอบว่า Path ต้นทางมีอยู่จริงหรือไม่
if not os.path.isdir(source_folder):
    print(f"\nข้อผิดพลาด: ไม่พบโฟลเดอร์ '{source_folder}'")
    sys.exit()

# สร้างโฟลเดอร์ผลลัพธ์หลัก (ถ้ายังไม่มี)
os.makedirs(output_folder, exist_ok=True)
# --- สิ้นสุดการตั้งค่า ---

def read_text_file(file_path):
    """อ่านไฟล์ข้อความธรรมดา"""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            return f.read()
    except Exception as e:
        return f"--- Error reading {os.path.basename(file_path)}: {e} ---"

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
        return f"--- Error reading Excel {os.path.basename(file_path)}: {e} ---"

# --- เริ่มกระบวนการแปลงไฟล์ ---
print(f"\nกำลังแปลงไฟล์จาก: {source_folder}")
print(f"ผลลัพธ์จะถูกเก็บไว้ที่: {output_folder}\n")

# วนลูปอ่านไฟล์ทั้งหมดในโฟลเดอร์ต้นทาง
for subdir, _, files in os.walk(source_folder):
    for file in files:
        # ข้ามตัวสคริปต์เอง
        if file == os.path.basename(__file__):
            continue

        # สร้าง Path แบบเต็มของไฟล์ต้นทาง
        source_file_path = os.path.join(subdir, file)

        # สร้าง Path ของไฟล์ผลลัพธ์ (.txt) โดยยังคงโครงสร้างโฟลเดอร์เดิม
        # 1. หา Path แบบ relative (ส่วนที่ต่อจาก source_folder)
        relative_path = os.path.relpath(source_file_path, source_folder)
        # 2. สร้าง Path ใหม่ในโฟลเดอร์ผลลัพธ์ และเติม .txt ต่อท้าย
        output_txt_path = os.path.join(output_folder, relative_path + '.txt')

        # สร้างโฟลเดอร์ย่อยในโฟลเดอร์ผลลัพธ์ (ถ้ายังไม่มี)
        output_subfolder = os.path.dirname(output_txt_path)
        os.makedirs(output_subfolder, exist_ok=True)

        print(f"กำลังแปลง: {source_file_path}  ->  {output_txt_path}")

        # อ่านเนื้อหาจากไฟล์ต้นทาง
        content = ""
        file_extension = os.path.splitext(file)[1].lower()
        if file_extension in ['.xlsx', '.xls']:
            content = read_excel_file(source_file_path)
        else: # ไฟล์ประเภทอื่นๆ ทั้งหมด
            content = read_text_file(source_file_path)

        # เขียนเนื้อหาลงในไฟล์ .txt ใหม่
        try:
            with open(output_txt_path, 'w', encoding='utf-8') as outfile:
                outfile.write(content)
        except Exception as e:
            print(f"  !! เกิดข้อผิดพลาดในการเขียนไฟล์: {e}")


print("\nแปลงไฟล์ทั้งหมดเรียบร้อยแล้ว!")