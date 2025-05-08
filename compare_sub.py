import pandas as pd

# File paths
csv_file = 'file1.csv'
excel_file = 'file2.xlsx'

# Read the files
df_csv = pd.read_csv(csv_file)
df_excel = pd.read_excel(excel_file)

# Helper to find actual column name ignoring case
def find_column(df, target):
    return next((col for col in df.columns if col.strip().lower() == target.lower()), None)

# Locate column names case-insensitively
csv_col = find_column(df_csv, 'Subscription ID')
excel_col = find_column(df_excel, 'Subscription ID')

# Validate both columns exist
if csv_col and excel_col:
    # Clean and normalize values
    df_csv[csv_col] = df_csv[csv_col].astype(str).str.strip()
    df_excel[excel_col] = df_excel[excel_col].astype(str).str.strip()

    # Create ID sets
    ids_csv = set(df_csv[csv_col])
    ids_excel = set(df_excel[excel_col])

    # Compare
    common_ids = ids_csv & ids_excel
    only_in_csv = ids_csv - ids_excel
    only_in_excel = ids_excel - ids_csv

    # Filter DataFrames
    df_common = df_csv[df_csv[csv_col].isin(common_ids)]
    df_only_csv = df_csv[df_csv[csv_col].isin(only_in_csv)]
    df_only_excel = df_excel[df_excel[excel_col].isin(only_in_excel)]

    # Save output
    df_common.to_excel('matched_subscriptions.xlsx', index=False)
    df_only_csv.to_excel('only_in_csv.xlsx', index=False)
    df_only_excel.to_excel('only_in_excel.xlsx', index=False)

    # Output summary
    print(f"✅ Common: {len(common_ids)} | Only in CSV: {len(only_in_csv)} | Only in Excel: {len(only_in_excel)}")
else:
    print("❌ 'Subscription ID' column not found in one or both files.")
