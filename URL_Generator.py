import pandas as pd

# === Configuration ===
excel_file_path = "your_excel_file.xlsx"               # <- Update your Excel file path
sheet_name = 0                                          # <- Sheet index or name
filter_column_name = "Severity"                        # <- Column to filter
filter_value = "Critical"                              # <- Value to filter on
target_column_name = "Subscription ID"                 # <- Column from which to extract values

# Per-item prefix/suffix
prefix = "prefix_data_for_each_subscriptions"
suffix = "suffix_data_for_each_subscriptions"

# Final combined string prefix/suffix
final_prefix = "FINAL_PREFIX_"
final_suffix = "_FINAL_SUFFIX"

# === Step 1: Load Excel File ===
print("[INFO] Loading Excel file...")
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# === Step 2: Check Required Columns ===
missing_cols = [col for col in [filter_column_name, target_column_name] if col not in df.columns]
if missing_cols:
    print(f"[ERROR] Column(s) not found in Excel: {', '.join(missing_cols)}")
else:
    # === Step 3: Filter Rows ===
    print(f"[INFO] Filtering rows where '{filter_column_name}' == '{filter_value}'...")
    filtered_df = df[df[filter_column_name].astype(str).str.strip().str.lower() == filter_value.strip().lower()]

    if filtered_df.empty:
        print(f"[WARN] No rows found with {filter_column_name} = '{filter_value}'")
    else:
        # === Step 4: Extract Target Values ===
        values = filtered_df[target_column_name].dropna().astype(str).unique().tolist()
        count = len(values)
        print(f"[INFO] Found {count} unique values in '{target_column_name}' after filtering.")

        # === Step 5: Format Each Item ===
        formatted_items = []
        for i, val in enumerate(values):
            item = f"{prefix}{val}"
            if i < count - 1:
                item += suffix
            formatted_items.append(item)

        combined_string = "".join(formatted_items)

        # === Step 6: Add Final Prefix/Suffix ===
        final_output = f"{final_prefix}{combined_string}{final_suffix}"

        # === Step 7: Output Result ===
        print("\n[RESULT] Final formatted string:\n")
        print(final_output)
        print(f"\n[INFO] Total count of unique '{target_column_name}': {count}")
