import pandas as pd

# === Configuration ===
excel_file = 'your_file.xlsx'          # Replace with your file path
sheet_name = 0                         # You can also use sheet name like 'Sheet1'
policy_id_col = 'Policy ID'
business_unit_col = 'Business Unit'    # Adjust if your BU column has a different name

# === Load Excel File ===
print("[INFO] Reading Excel file...")
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# === Drop rows where Policy ID is missing ===
df = df.dropna(subset=[policy_id_col])

# === Count total occurrences of each unique Policy ID ===
print("[INFO] Calculating total Policy ID counts...")
total_counts = df[policy_id_col].value_counts().rename_axis(policy_id_col).reset_index(name='Total Count')

# === Count occurrences of each Policy ID per Business Unit ===
print("[INFO] Calculating BU-wise Policy ID counts...")
grouped = df.groupby([policy_id_col, business_unit_col]).size().unstack(fill_value=0)

# === Merge both into one summary table ===
print("[INFO] Merging total and BU-wise counts...")
summary_df = pd.merge(total_counts, grouped, on=policy_id_col)

# === Save to Excel ===
output_file = 'policy_id_summary.xlsx'
summary_df.to_excel(output_file, index=False)
print(f"[SUCCESS] Summary written to {output_file}")
