import pandas as pd
import time
from datetime import datetime
from pathlib import Path

# === CONFIGURATION ===
csv_file_path = 'input_data.csv'
reference_excel_path = 'reference.xlsx'
history_file = 'issue_history.xlsx'
today_str = datetime.today().strftime('%Y-%m-%d')
output_excel_path = f'tracked_report_{today_str}.xlsx'

# Key columns to uniquely identify an issue
key_columns = ['EmpID', 'Policy ID']

# Optional transformations
columns_to_rename = {'EmpName': 'Employee Name', 'Dept': 'Department'}
columns_to_add = {'Reviewed': 'No', 'Reviewer': ''}
columns_to_split = {
    'Location': {'delimiter': ',', 'new_columns': ['City', 'State']},
    'FullName': {'delimiter': ' ', 'new_columns': ['FirstName', 'LastName']}
}
columns_to_remove = ['UnwantedCol1', 'UnwantedCol2']
rows_to_filter_out = {'Department': ['HR', 'Finance']}
reference_fields_to_enrich = ['Policy Statement', 'Policy Remediation']
policy_id_column = 'Policy ID'
desired_column_order = [
    'Employee Name', 'EmpID', 'Department', 'City', 'State',
    'FirstName', 'LastName', 'Policy ID', 'Policy Statement',
    'Policy Remediation', 'Reviewed', 'Reviewer',
    'Issue Creation Date', 'Status', 'Closed Date'
]

# === START TIMER ===
print("=== Script started ===")
start_time = time.time()

# === Load and preprocess today's report ===
print(f"Reading today's input CSV: {csv_file_path}")
df = pd.read_csv(csv_file_path)
df.rename(columns=columns_to_rename, inplace=True)
for col, val in columns_to_add.items():
    df[col] = val

for col, cfg in columns_to_split.items():
    if col in df.columns:
        split_df = df[col].astype(str).str.split(cfg['delimiter'], n=1, expand=True)
        split_df.columns = cfg['new_columns']
        df = pd.concat([df, split_df], axis=1)

df.drop(columns=[col for col in columns_to_remove if col in df.columns], inplace=True)

for col, values in rows_to_filter_out.items():
    if col in df.columns:
        original_len = len(df)
        df = df[~df[col].isin(values)]
        print(f"Filtered {original_len - len(df)} rows from column '{col}'")

for key in key_columns:
    df[key] = df[key].astype(str)

# === Enrich from reference file ===
print(f"Enriching from reference file: {reference_excel_path}")
ref_df = pd.read_excel(reference_excel_path)
ref_df[policy_id_column] = ref_df[policy_id_column].astype(str)
df[policy_id_column] = df[policy_id_column].astype(str)
df = df.merge(ref_df[[policy_id_column] + reference_fields_to_enrich], on=policy_id_column, how='left')

# === Add tracking info ===
df['Issue Creation Date'] = today_str
df['Status'] = 'Open'
df['Closed Date'] = ''

# === Load or initialize issue history ===
if Path(history_file).exists():
    print("Reading existing issue history...")
    history_df = pd.read_excel(history_file)
    for key in key_columns:
        history_df[key] = history_df[key].astype(str)
else:
    print("No issue history found. Initializing new one.")
    history_df = pd.DataFrame(columns=df.columns)

# === Create key sets for fast comparison ===
def get_key(row):
    return tuple(row[k] for k in key_columns)

history_keys = set(history_df[key_columns].astype(str).apply(tuple, axis=1))
today_keys = set(df[key_columns].astype(str).apply(tuple, axis=1))

# === Update existing issues ===
print("Updating status of existing issues...")
history_df['Status'] = history_df.apply(
    lambda row: 'Open' if get_key(row) in today_keys else 'Closed',
    axis=1
)
history_df['Closed Date'] = history_df.apply(
    lambda row: '' if row['Status'] == 'Open' else today_str,
    axis=1
)

# === Add new issues ===
print("Adding new issues to history...")
new_issues = df[df.apply(lambda row: get_key(row) not in history_keys, axis=1)]
history_df = pd.concat([history_df, new_issues], ignore_index=True)

# === Reorder columns ===
ordered = [col for col in desired_column_order if col in history_df.columns]
remaining = [col for col in history_df.columns if col not in ordered]
final_df = history_df[ordered + remaining]

# === Save output files ===
print(f"Saving today's tracked report: {output_excel_path}")
final_df.to_excel(output_excel_path, index=False)

print(f"Updating issue history file: {history_file}")
final_df.to_excel(history_file, index=False)

# === Stop timer ===
end_time = time.time()
duration = end_time - start_time
if duration < 60:
    print(f"Script completed in {duration:.2f} seconds")
elif duration < 3600:
    print(f"Script completed in {int(duration // 60)} min {duration % 60:.2f} sec")
else:
    h = int(duration // 3600)
    m = int((duration % 3600) // 60)
    s = duration % 60
    print(f"Script completed in {h} hr {m} min {s:.2f} sec")

print("=== Done ===")
