import pandas as pd
import time
import os

# === CONFIGURATION ===

# File paths
csv_file_path = 'input_data.csv'
excel_output_path = 'output_data.xlsx'
reference_excel_path = 'reference.xlsx'

# Rename columns: {old_name: new_name}
columns_to_rename = {
    'EmpName': 'Employee Name',
    'Dept': 'Department'
}

# Add new columns with default values
columns_to_add = {
    'Reviewed': 'No',
    'Reviewer': ''
}

# Split columns: {column_name: {'delimiter': str, 'new_columns': [col1, col2]}}
columns_to_split = {
    'Location': {
        'delimiter': ',',
        'new_columns': ['City', 'State']
    },
    'FullName': {
        'delimiter': ' ',
        'new_columns': ['FirstName', 'LastName']
    }
}

# Remove columns
columns_to_remove = ['UnwantedCol1', 'UnwantedCol2']

# Filter rows: {column_name: [values_to_remove]}
rows_to_filter_out = {
    'Department': ['HR', 'Finance']
}

# Reference match + value copy: match column, and {reference_column: destination_column_in_df}
reference_match_column = 'Employee ID'
input_column_to_match = 'EmpID'
reference_fill_map = {
    'Manager': 'Manager'  # Copy Manager from reference to Manager column in df
}

# Desired final column order
desired_column_order = [
    'Employee Name', 'EmpID', 'Department', 'City', 'State',
    'FirstName', 'LastName', 'Reviewed', 'Reviewer', 'Manager'
]

# === EXECUTION START ===
start_time = time.time()

# === FILE CHECK ===
if not os.path.exists(csv_file_path):
    raise FileNotFoundError(f"Input CSV not found: {csv_file_path}")
if not os.path.exists(reference_excel_path):
    raise FileNotFoundError(f"Reference Excel not found: {reference_excel_path}")

# === LOAD FILES ===
df = pd.read_csv(csv_file_path)
ref_df = pd.read_excel(reference_excel_path)

print("\n📋 Columns in Input File:")
for col in df.columns:
    print(f" - {col}")

# === RENAME COLUMNS ===
df.rename(columns=columns_to_rename, inplace=True)

# === ADD NEW COLUMNS ===
for col, default_value in columns_to_add.items():
    df[col] = default_value

# === SPLIT COLUMNS ===
for col, cfg in columns_to_split.items():
    if col in df.columns:
        splits = df[col].astype(str).str.split(cfg['delimiter'], n=1, expand=True)
        if len(splits.columns) < 2:
            splits[1] = ''
        splits.columns = cfg['new_columns']
        df = pd.concat([df, splits], axis=1)
    else:
        print(f"⚠️ Column '{col}' not found for splitting")

# === REMOVE COLUMNS ===
df.drop(columns=[col for col in columns_to_remove if col in df.columns], inplace=True)

# === FILTER OUT ROWS ===
for col, values in rows_to_filter_out.items():
    if col in df.columns:
        before = len(df)
        df = df[~df[col].isin(values)]
        after = len(df)
        print(f"🚫 Removed {before - after} rows where '{col}' in {values}")
    else:
        print(f"⚠️ Column '{col}' not found for filtering")

# === ENRICH EXISTING COLUMNS USING REFERENCE FILE ===
if reference_match_column in ref_df.columns and input_column_to_match in df.columns:
    ref_subset = ref_df[[reference_match_column] + list(reference_fill_map.keys())]
    df = df.merge(ref_subset, how='left', left_on=input_column_to_match, right_on=reference_match_column)

    # Fill mapped values and drop the reference column
    for ref_col, target_col in reference_fill_map.items():
        if ref_col in df.columns:
            df[target_col] = df[ref_col]
            if ref_col != target_col:
                df.drop(columns=[ref_col], inplace=True)
    if reference_match_column in df.columns:
        df.drop(columns=[reference_match_column], inplace=True)
else:
    print("⚠️ Enrichment skipped: Matching columns not found")

# === REORDER COLUMNS ===
reordered_cols = [col for col in desired_column_order if col in df.columns]
remaining_cols = [col for col in df.columns if col not in reordered_cols]
df = df[reordered_cols + remaining_cols]

print("\n📐 Final Column Order:")
for col in df.columns:
    print(f" - {col}")

# === SAVE TO EXCEL ===
df.to_excel(excel_output_path, index=False)
print(f"\n✅ Excel saved: {excel_output_path}")

# === EXECUTION TIME ===
end_time = time.time()
duration = end_time - start_time

print("\n⏱️ Execution Time:")
if duration < 60:
    print(f" - {duration:.2f} seconds")
elif duration < 3600:
    print(f" - {int(duration // 60)} minutes {duration % 60:.2f} seconds")
else:
    h = int(duration // 3600)
    m = int((duration % 3600) // 60)
    s = duration % 60
    print(f" - {h} hours {m} minutes {s:.2f} seconds")
