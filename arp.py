import pandas as pd
import time
import os
from datetime import datetime
from pathlib import Path

# === CONFIGURATION ===
csv_file_path = 'input_data.csv'
reference_excel_path = 'reference.xlsx'
history_file = 'issue_history.xlsx'
today_str = datetime.today().strftime('%Y-%m-%d')
output_excel_path = f'tracked_report_{today_str}.xlsx'

columns_to_rename = {'EmpName': 'Employee Name', 'Dept': 'Department'}
columns_to_add = {'Reviewed': 'No', 'Reviewer': ''}
columns_to_split = {
    'Location': {'delimiter': ',', 'new_columns': ['City', 'State']},
    'FullName': {'delimiter': ' ', 'new_columns': ['FirstName', 'LastName']}
}
columns_to_remove = ['UnwantedCol1', 'UnwantedCol2']
rows_to_filter_out = {'Department': ['HR', 'Finance']}
policy_id_column = 'Policy ID'
reference_fields_to_enrich = ['Policy Statement', 'Policy Remediation']
desired_column_order = [
    'Employee Name', 'EmpID', 'Department', 'City', 'State',
    'FirstName', 'LastName', 'Policy ID', 'Policy Statement',
    'Policy Remediation', 'Reviewed', 'Reviewer',
    'Issue Creation Date', 'Status', 'Closed Date'
]
key_columns = ['EmpID', 'Policy ID']

print("=== Script started ===")
start_time = time.time()

# === LOAD FILES ===
print(f"Reading input file: {csv_file_path}")
df = pd.read_csv(csv_file_path)

print(f"Reading reference file: {reference_excel_path}")
ref_df = pd.read_excel(reference_excel_path)

# === TYPE NORMALIZATION ===
print("Normalizing types for merge columns...")
df[policy_id_column] = df[policy_id_column].astype(str)
ref_df[policy_id_column] = ref_df[policy_id_column].astype(str)
for key in key_columns:
    df[key] = df[key].astype(str)

# === RENAME COLUMNS ===
print("Renaming columns...")
df.rename(columns=columns_to_rename, inplace=True)

# === ADD NEW COLUMNS ===
print("Adding new columns...")
for col, val in columns_to_add.items():
    df[col] = val

# === SPLIT COLUMNS ===
print("Splitting configured columns...")
for col, cfg in columns_to_split.items():
    if col in df.columns:
        print(f" - Splitting '{col}' into {cfg['new_columns']}")
        split_df = df[col].astype(str).str.split(cfg['delimiter'], n=1, expand=True)
        if len(split_df.columns) < 2:
            split_df[1] = ''
        split_df.columns = cfg['new_columns']
        df = pd.concat([df, split_df], axis=1)

# === REMOVE UNNEEDED COLUMNS ===
print("Removing unwanted columns...")
df.drop(columns=[col for col in columns_to_remove if col in df.columns], inplace=True)

# === FILTER ROWS ===
print("Filtering rows based on configured conditions...")
for col, values in rows_to_filter_out.items():
    if col in df.columns:
        original_len = len(df)
        df = df[~df[col].isin(values)]
        print(f" - Removed {original_len - len(df)} rows from column '{col}'")

# === ENRICH DATA ===
print("Enriching data from reference file...")
if policy_id_column in df.columns and policy_id_column in ref_df.columns:
    enrichment_df = ref_df[[policy_id_column] + reference_fields_to_enrich]
    df = df.merge(enrichment_df, how='left', on=policy_id_column)

# === TRACKING COLUMNS ===
print("Adding tracking columns...")
df['Issue Creation Date'] = today_str
df['Status'] = 'Open'
df['Closed Date'] = ''

# === HANDLE ISSUE HISTORY ===
if Path(history_file).exists():
    print("Existing issue history found. Comparing...")
    history_df = pd.read_excel(history_file)
    for key in key_columns:
        history_df[key] = history_df[key].astype(str)

    merged = history_df.merge(df, how='outer', on=key_columns, suffixes=('_old', ''))
    rows = []

    for _, row in merged.iterrows():
        if pd.notna(row.get('Status_old')) and pd.isna(row.get('Status')):
            for col in merged.columns:
                if col.endswith('_old'):
                    base_col = col.replace('_old', '')
                    if pd.isna(row.get(base_col)) and pd.notna(row.get(col)):
                        row[base_col] = row[col]
            row['Status'] = 'Closed'
            row['Closed Date'] = today_str
            row['Issue Creation Date'] = row['Issue Creation Date_old']
            rows.append(row.to_dict())

        elif pd.isna(row.get('Status_old')) and pd.notna(row.get('Status')):
            rows.append(row.to_dict())

        elif pd.notna(row.get('Status_old')) and pd.notna(row.get('Status')):
            row['Issue Creation Date'] = row['Issue Creation Date_old']
            row['Status'] = 'Open'
            row['Closed Date'] = ''
            rows.append(row.to_dict())

    final_df = pd.DataFrame(rows)
    final_df.drop(columns=[col for col in final_df.columns if col.endswith('_old')], inplace=True)
else:
    print("No previous issue history. This is the first run.")
    final_df = df.copy()

# === REORDER COLUMNS ===
print("Reordering columns...")
ordered = [col for col in desired_column_order if col in final_df.columns]
remaining = [col for col in final_df.columns if col not in ordered]
final_df = final_df[ordered + remaining]

# === SAVE OUTPUT ===
print(f"Saving final report to: {output_excel_path}")
final_df.to_excel(output_excel_path, index=False)

print(f"Updating issue history file: {history_file}")
final_df.to_excel(history_file, index=False)

# === EXECUTION TIME ===
end_time = time.time()
duration = end_time - start_time
if duration < 60:
    print(f"Execution time: {duration:.2f} seconds")
elif duration < 3600:
    print(f"Execution time: {int(duration // 60)} min {duration % 60:.2f} sec")
else:
    h = int(duration // 3600)
    m = int((duration % 3600) // 60)
    s = duration % 60
    print(f"Execution time: {h} hr {m} min {s:.2f} sec")

print("=== Script completed successfully ===")
