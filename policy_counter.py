import pandas as pd

# === Configuration ===
excel_file = 'your_file.xlsx'           # Input Excel file
sheet_name = 0                          # Sheet name or index
policy_id_col = 'Policy ID'             # Column name for Policy ID
policy_name_col = 'Policy Name'         # Column name for Policy Name
business_unit_col = 'Business Unit'     # Column name for Business Unit
output_file = 'policy_id_summary.xlsx'  # Output Excel file name

# === Load Excel File ===
print("[INFO] Reading Excel file...")
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# === Drop rows where Policy ID is missing ===
df = df.dropna(subset=[policy_id_col])

# === Get unique Policy ID + Policy Name mapping ===
print("[INFO] Fetching Policy Name mapping...")
policy_mapping = df[[policy_id_col, policy_name_col]].drop_duplicates()

# === Count total occurrences of each Policy ID ===
print("[INFO] Calculating total Policy ID counts...")
total_counts = df[policy_id_col].value_counts().rename_axis(policy_id_col).reset_index(name='Total Count')

# === Count occurrences of each Policy ID per Business Unit ===
print("[INFO] Calculating BU-wise Policy ID counts...")
grouped = df.groupby([policy_id_col, business_unit_col]).size().unstack(fill_value=0)

# === Merge all data: Total Count + Policy Name + BU-wise counts ===
print("[INFO] Merging total, names and BU-wise counts...")
merged = pd.merge(total_counts, policy_mapping, on=policy_id_col)
summary_df = pd.merge(merged, grouped, on=policy_id_col)

# === Reorder columns: Policy ID | Policy Name | Total Count | BU Columns... ===
cols = [policy_id_col, policy_name_col, 'Total Count'] + [col for col in summary_df.columns if col not in [policy_id_col, policy_name_col, 'Total Count']]
summary_df = summary_df[cols]

# === Save to Excel ===
summary_df.to_excel(output_file, index=False)
print(f"[SUCCESS] Summary written to {output_file}")
