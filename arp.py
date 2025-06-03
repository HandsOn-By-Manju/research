import pandas as pd
import time
import os
from datetime import datetime
from pathlib import Path

# === CONFIGURATION ===

csv_file_path = 'input_data.csv'             # Main input report (daily)
reference_excel_path = 'reference.xlsx'      # Reference file for enrichment
history_file = 'issue_history.xlsx'          # History tracker for comparison
today_str = datetime.today().strftime('%Y-%m-%d')
output_excel_path = f'tracked_report_{today_str}.xlsx'

# Define how columns should be renamed, added, split, removed
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
key_columns = ['EmpID', 'Policy ID']  # Unique keys to track an issue

# === TRACK EXECUTION TIME ===
start_time = time.time()

# === LOAD FILES ===
if not os.path.exists(csv_file_path):
    raise FileNotFoundError(f"Missing input file: {csv_file_path}")
if not os.path.exists(reference_excel_path):
    raise FileNotFoundError(f"Missing reference file: {reference_excel_path}")

df = pd.read_csv(csv_file_path)
ref_df = pd.read_excel(reference_excel_path)

# === TYPE NORMALIZATION FOR MERGE KEYS ===
df[policy_id_column] = df[policy_id_column].astype(str)
ref_df[policy_id_column] = ref_df[policy_id_column].astype(str)
for key in key_columns:
    df[key] = df[key].astype(str)

# === RENAME COLUMNS ===
df.rename(columns=columns_to_rename, inplace=True)

# === ADD DEFAULT COLUMNS ===
for col, val in columns_to_add.items():
    df[col] = val

# === SPLIT COLUMNS ===
for col, cfg in columns_to_split.items():
    if col in df.columns:
        split_df = df[col].astype(str).str.split(cfg['delimiter'], n=1, expand=True)
        if len(split_df.columns) < 2:
            split_df[1] = ''
        split_df.columns = cfg['new_columns']
        df = pd.concat([df, split_df], axis=1)

# === REMOVE UNNECESSARY COLUMNS ===
df.drop(columns=[col for col in columns_to_remove if col in df.columns], inplace=True)

# === FILTER ROWS ===
for col, values in rows_to_filter_out.items():
    if col in df.columns:
        df = df[~df[col].isin(values)]

# === ENRICH USING POLICY REFERENCE ===
if policy_id_column in df.columns and policy_id_column in ref_df.columns:
    enrichment_df = ref_df[[policy_id_column] + reference_fields_to_enrich]
    df = df.merge(enrichment_df, how='left', on=policy_id_column)

# === TRACKING COLUMNS FOR ISSUE LIFECYCLE ===
df['Issue Creation Date'] = today_str
df['Status'] = 'Open'
df['Closed Date'] = ''

# === COMPARE WITH ISSUE HISTORY ===
if Path(history_file).exists():
    history_df = pd.read_excel(history_file)
    for key in key_columns:
        history_df[key] = history_df[key].astype(str)

    # Merge on keys to compare new and old data
    merged = history_df.merge(df, how='outer', on=key_columns, suffixes=('_old', ''))
    rows = []

    # Determine what changed
    for _, row in merged.iterrows():
        if pd.notna(row.get('Status_old')) and pd.isna(row.get('Status')):
            # Found in history but not today → Closed
            row['Status'] = 'Closed'
            row['Closed Date'] = today_str
            row['Issue Creation Date'] = row['Issue Creation Date_old']
            rows.append(row.to_dict())
        elif pd.isna(row.get('Status_old')) and pd.notna(row.get('Status')):
            # New issue today
            rows.append(row.to_dict())
        elif pd.notna(row.get('Status_old')) and pd.notna(row.get('Status')):
            # Still exists → keep status open
            row['Issue Creation Date'] = row['Issue Creation Date_old']
            row['Status'] = 'Open'
            row['Closed Date'] = ''
            rows.append(row.to_dict())

    final_df = pd.DataFrame(rows)
    final_df.drop(columns=[col for col in final_df.columns if col.endswith('_old')], inplace=True)
else:
    # First time run — treat all as new issues
    final_df = df.copy()

# === REORDER COLUMNS ===
ordered = [col for col in desired_column_order if col in final_df.columns]
remaining = [col for col in final_df.columns if col not in ordered]
final_df = final_df[ordered + remaining]

# === SAVE TO EXCEL ===
final_df.to_excel(output_excel_path, index=False)
final_df.to_excel(history_file, index=False)

# === CALCULATE EXECUTION TIME ===
end_time = time.time()
duration = end_time - start_time

# === DISPLAY EXECUTION TIME ===
if duration < 60:
    print(f"Execution time: {duration:.2f} seconds")
elif duration < 3600:
    print(f"Execution time: {int(duration // 60)} minutes {duration % 60:.2f} seconds")
else:
    h = int(duration // 3600)
    m = int((duration % 3600) // 60)
    s = duration % 60
    print(f"Execution time: {h} hours {m} minutes {s:.2f} seconds")
