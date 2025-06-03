import pandas as pd
import time
import os
from datetime import datetime
from pathlib import Path

# === CONFIGURATION SECTION ===

# Paths to the input files (can be modified as needed)
csv_file_path = 'input_data.csv'             # Input data (daily report)
reference_excel_path = 'reference.xlsx'      # Reference policy data
history_file = 'issue_history.xlsx'          # Persistent file to track issue lifecycle

# Today's date for tagging and filenames
today_str = datetime.today().strftime('%Y-%m-%d')
output_excel_path = f'tracked_report_{today_str}.xlsx'

# Configurable column mappings
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
key_columns = ['EmpID', 'Policy ID']  # Unique identifier for tracking

# === STEP 1: START TIMER ===
start_time = time.time()

# === STEP 2: LOAD INPUT FILES ===
if not os.path.exists(csv_file_path):
    raise FileNotFoundError(f"Input CSV not found: {csv_file_path}")
if not os.path.exists(reference_excel_path):
    raise FileNotFoundError(f"Reference Excel not found: {reference_excel_path}")

df = pd.read_csv(csv_file_path)
ref_df = pd.read_excel(reference_excel_path)

# === STEP 3: RENAME COLUMNS ===
df.rename(columns=columns_to_rename, inplace=True)

# === STEP 4: ADD CONFIGURED COLUMNS WITH DEFAULT VALUES ===
for col, val in columns_to_add.items():
    df[col] = val

# === STEP 5: SPLIT COLUMNS ===
for col, cfg in columns_to_split.items():
    if col in df.columns:
        split_df = df[col].astype(str).str.split(cfg['delimiter'], n=1, expand=True)
        if len(split_df.columns) < 2:
            split_df[1] = ''
        split_df.columns = cfg['new_columns']
        df = pd.concat([df, split_df], axis=1)

# === STEP 6: REMOVE UNWANTED COLUMNS ===
df.drop(columns=[col for col in columns_to_remove if col in df.columns], inplace=True)

# === STEP 7: FILTER OUT UNWANTED ROWS ===
for col, values in rows_to_filter_out.items():
    if col in df.columns:
        df = df[~df[col].isin(values)]

# === STEP 8: ENRICH DATA USING REFERENCE FILE ===
if policy_id_column in df.columns and policy_id_column in ref_df.columns:
    enrichment_df = ref_df[[policy_id_column] + reference_fields_to_enrich]
    df = df.merge(enrichment_df, how='left', on=policy_id_column)

# === STEP 9: INITIALIZE ISSUE TRACKING FIELDS ===
df['Issue Creation Date'] = today_str
df['Status'] = 'Open'
df['Closed Date'] = ''

# === STEP 10: COMPARE WITH HISTORY FILE (IF EXISTS) ===
if Path(history_file).exists():
    history_df = pd.read_excel(history_file)
    merged = history_df.merge(df, how='outer', on=key_columns, suffixes=('_old', ''))

    rows = []
    for _, row in merged.iterrows():
        if pd.notna(row.get('Status_old')) and pd.isna(row.get('Status')):
            # Issue disappeared today → mark as closed
            row['Status'] = 'Closed'
            row['Closed Date'] = today_str
            row['Issue Creation Date'] = row['Issue Creation Date_old']
            rows.append(row)
        elif pd.isna(row.get('Status_old')) and pd.notna(row.get('Status')):
            # New issue today
            rows.append(row)
        elif pd.notna(row.get('Status_old')) and pd.notna(row.get('Status')):
            # Still open → keep original creation date
            row['Issue Creation Date'] = row['Issue Creation Date_old']
            row['Status'] = 'Open'
            row['Closed Date'] = ''
            rows.append(row)

    final_df = pd.DataFrame(rows)
    final_df.drop(columns=[col for col in final_df.columns if col.endswith('_old')], inplace=True)
else:
    # First-time run → all issues are new
    final_df = df.copy()

# === STEP 11: REORDER COLUMNS IF CONFIGURED ===
ordered = [col for col in desired_column_order if col in final_df.columns]
remaining = [col for col in final_df.columns if col not in ordered]
final_df = final_df[ordered + remaining]

# === STEP 12: SAVE OUTPUT REPORT & ISSUE HISTORY ===
final_df.to_excel(output_excel_path, index=False)
final_df.to_excel(history_file, index=False)

# === STEP 13: PRINT EXECUTION TIME ===
end_time = time.time()
duration = end_time - start_time
if duration < 60:
    exec_time_str = f"{duration:.2f} seconds"
elif duration < 3600:
    exec_time_str = f"{int(duration // 60)} minutes {duration % 60:.2f} seconds"
else:
    h = int(duration // 3600)
    m = int((duration % 3600) // 60)
    s = duration % 60
    exec_time_str = f"{h} hours {m} minutes {s:.2f} seconds"

print(f"\n✅ Script completed in {exec_time_str}")
print(f"📁 Output saved to: {output_excel_path}")
print(f"🧾 Issue history updated: {history_file}")
