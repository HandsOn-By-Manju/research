import pandas as pd

# Path to your Excel file
file_path = 'sample.xlsx'

print("[INFO] Starting the process to read Excel file...")

try:
    # Load the Excel file
    print(f"[INFO] Reading the Excel file from: {file_path}")
    df = pd.read_excel(file_path)

    # Display total number of columns
    print(f"[SUCCESS] Excel file loaded successfully. Total columns found: {len(df.columns)}\n")

    # Print all column names
    print("[INFO] Listing all column names:")
    for col in df.columns:
        print(f"- {col}")

except FileNotFoundError:
    print(f"[ERROR] File not found at path: {file_path}")
except Exception as e:
    print(f"[ERROR] An unexpected error occurred: {e}")
