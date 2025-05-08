import pandas as pd

# File paths
csv_file = 'file1.csv'
excel_file = 'file2.xlsx'

# Load files
try:
    df_csv = pd.read_csv(csv_file)
    df_excel = pd.read_excel(excel_file)
    print("✅ Files loaded successfully.")
except Exception as e:
    print(f"❌ Error loading files: {e}")
    exit()

# Function to locate column name (case-insensitive)
def get_matching_column(df, target_col_name):
    for col in df.columns:
        if col.strip().lower() == target_col_name.strip().lower():
            return col
    return None

# Detect 'Subscription ID' column in both files
csv_sub_col = get_matching_column(df_csv, 'Subscription ID')
excel_sub_col = get_matching_column(df_excel, 'Subscription ID')

# Handle missing columns with specific messages
if not csv_sub_col or not excel_sub_col:
    print("❌ Column 'Subscription ID' not found in:")
    if not csv_sub_col:
        print(f"   📄 CSV File Columns: {list(df_csv.columns)}")
    if not excel_sub_col:
        print(f"   📄 Excel File Columns: {list(df_excel.columns)}")
    exit()

print(f"🔍 Matching on column: '{csv_sub_col}' in CSV and '{excel_sub_col}' in Excel.")

# Normalize ID values (strip spaces + convert to string)
df_csv[csv_sub_col] = df_csv[csv_sub_col].astype(str).str.strip()
df_excel[excel_sub_col] = df_excel[excel_sub_col].astype(str).str.strip()

# Extract and compare Subscription IDs
ids_csv = set(df_csv[csv_sub_col])
ids_excel = set(df_excel[excel_sub_col])

common_ids = ids_csv & ids_excel
only_in_csv = ids_csv - ids_excel
only_in_excel = ids_excel - ids_csv

# Print summary
print(f"\n✅ Comparison Summary:")
print(f"   🔁 Common IDs       : {len(common_ids)}")
print(f"   📄 Only in CSV      : {len(only_in_csv)}")
print(f"   📄 Only in Excel    : {len(only_in_excel)}")

# Filter rows based on match
df_common = df_csv[df_csv[csv_sub_col].isin(common_ids)]
df_only_csv = df_csv[df_csv[csv_sub_col].isin(only_in_csv)]
df_only_excel = df_excel[df_excel[excel_sub_col].isin(only_in_excel)]

# Save to Excel
df_common.to_excel('matched_subscriptions.xlsx', index=False)
df_only_csv.to_excel('only_in_csv.xlsx', index=False)
df_only_excel.to_excel('only_in_excel.xlsx', index=False)

print("\n📁 Output files created successfully:")
print("   ✅ matched_subscriptions.xlsx")
print("   ✅ only_in_csv.xlsx")
print("   ✅ only_in_excel.xlsx")
