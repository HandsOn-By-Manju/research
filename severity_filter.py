import pandas as pd

# --- Config ---
input_file = "input.xlsx"  # Change to your actual file path
sheet_name = "Sheet1"
column_name = "Severity"
filter_value = "Critical"

# --- Read Excel ---
df = pd.read_excel(input_file, sheet_name=sheet_name)

# --- Filter only 'Critical' rows ---
df_filtered = df[df[column_name].str.strip().str.lower() == filter_value.lower()]

# --- Save back to the same file ---
df_filtered.to_excel(input_file, index=False)

print(f"Filtered and saved only '{filter_value}' severity rows to: {input_file}")
