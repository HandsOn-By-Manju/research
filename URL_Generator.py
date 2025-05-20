import pandas as pd

# === Configuration ===
excel_file_path = "your_excel_file.xlsx"               # <- Input Excel file
sheet_name = 0                                          # <- Sheet index or name
filter_column_name = "Severity"                        # <- Column to filter by
target_column_name = "Subscription ID"                 # <- Column to extract

# Per-item prefix/suffix
prefix = "prefix_data_for_each_subscriptions"
suffix = "suffix_data_for_each_subscriptions"

# Final base URL and suffix
final_base_url = "https://example.com/trigger/"
final_suffix = "_FINAL_SUFFIX"

# Output text file
output_text_file = "Generated_URL_Report.txt"

# === Step 1: Load Excel ===
print("[INFO] Loading Excel file...")
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# === Step 2: Validate Columns ===
missing = [col for col in [filter_column_name, target_column_name] if col not in df.columns]
if missing:
    print(f"[ERROR] Missing columns in input file: {', '.join(missing)}")
    exit()

# === Step 3: Process Each Unique Filter Value ===
unique_values = df[filter_column_name].dropna().astype(str).unique().tolist()
print(f"[INFO] Found {len(unique_values)} unique values in '{filter_column_name}' column.\n")

output_lines = []

for filter_val in unique_values:
    print(f"[SECTION] Processing: {filter_column_name} = '{filter_val}'")

    # Filter the dataframe
    filtered_df = df[df[filter_column_name].astype(str).str.strip().str.lower() ==
                     filter_val.strip().lower()]

    # Extract unique target values
    values = filtered_df[target_column_name].dropna().astype(str).unique().tolist()
    count = len(values)

    if count == 0:
        print(f"  [WARN] No '{target_column_name}' values found for '{filter_val}'")
        continue

    # Build URL string
    formatted_items = []
    for i, val in enumerate(values):
        item = f"{prefix}{val}"
        if i < len(values) - 1:
            item += suffix
        formatted_items.append(item)

    combined_string = "".join(formatted_items)
    full_url = f"{final_base_url}{combined_string}{final_suffix}"

    # Console output
    print(f"  [INFO] Unique '{target_column_name}' values: {count}")
    print(f"  [RESULT] URL: {full_url}\n")

    # Add to text output
    output_lines.append(f"[Filter Value: {filter_val}]")
    output_lines.append(f"Subscription Count: {count}")
    output_lines.append(f"URL: {full_url}")
    output_lines.append("")  # Blank line between blocks

# === Step 4: Write to File ===
with open(output_text_file, "w", encoding="utf-8") as f:
    f.write("\n".join(output_lines))

print(f"[SUCCESS] All results written to: {output_text_file}")
