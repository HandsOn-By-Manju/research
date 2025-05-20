import os
import pandas as pd

# === STEP 1: Setup ===
folder_path = os.getcwd()
excel_extensions = ('.xlsx', '.xls')
excel_files = [f for f in os.listdir(folder_path) if f.endswith(excel_extensions) and not f.startswith('~$')]

print(f"📂 Current Folder: {folder_path}")
print(f"📄 Excel Files Found: {len(excel_files)}")
for f in excel_files:
    print(f"   - {f}")

if not excel_files:
    print("❌ No Excel files found. Exiting.")
    exit()

# === STEP 2: Column Validation ===
first_file = pd.read_excel(excel_files[0])
expected_columns = list(first_file.columns)
expected_col_count = len(expected_columns)

print(f"\n📌 Expected Column Structure (from first file: {excel_files[0]}):")
print(f"   Column Count: {expected_col_count}")
print(f"   Column Names: {expected_columns}")

all_match = True
for file in excel_files:
    df = pd.read_excel(file)
    current_columns = list(df.columns)
    if current_columns != expected_columns:
        print(f"❌ Column mismatch in: {file}")
        print(f"   Found Columns: {current_columns}")
        all_match = False
    else:
        print(f"✔️ Columns match in: {file}")

if all_match:
    print("\n✅ All files have matching column names, count, and sequence.")
else:
    print("\n⚠️ One or more files have mismatched columns. Please fix and retry.")
    exit()

# === STEP 3: Count by Severity and Policy ID ===
print("\n📊 Row Counts by Severity and Policy ID:")
for file in excel_files:
    df = pd.read_excel(file)
    print(f"\n📁 File: {file}")
    
    if 'Severity' in df.columns:
        print("   🔹 Severity Counts:")
        print(df['Severity'].value_counts().to_string())
    else:
        print("   ⚠️ 'Severity' column not found.")
    
    if 'Policy ID' in df.columns:
        print("   🔹 Policy ID Counts (Top 10):")
        print(df['Policy ID'].value_counts().head(10).to_string())
    else:
        print("   ⚠️ 'Policy ID' column not found.")
