import pandas as pd

file_path = 'sample.xlsx'

print("[INFO] Starting the process to read Excel file...")

try:
    # Load all sheet names
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    print(f"[INFO] Sheets found in the Excel file: {sheet_names}\n")

    # Loop through each sheet and list column names
    for sheet in sheet_names:
        print(f"[INFO] Reading sheet: {sheet}")
        df = pd.read_excel(xls, sheet_name=sheet)
        print(f"[SUCCESS] Sheet '{sheet}' loaded successfully. Total columns: {len(df.columns)}")

        print("[INFO] Column names:")
        for col in df.columns:
            print(f"- {col}")
        print("-" * 50)

except FileNotFoundError:
    print(f"[ERROR] File not found at path: {file_path}")
except Exception as e:
    print(f"[ERROR] An unexpected error occurred: {e}")
