import pandas as pd

# === Configuration ===
excel_file_path = "your_excel_file.xlsx"         # <- Update this path
sheet_name = 0                                    # <- Sheet index or name
filter_column_name = "Severity"                  # <- Column to filter by
filter_value = "Critical"                        # <- Value to match
target_column_name = "Subscription ID"           # <- Column to extract values from
prefix = "prefix_data_for_each_subscriptions"    # <- Prefix
suffix = "suffix_data_for_each_subscriptions"    # <- Suffix

# === Step 1: Load Excel File ===
print("[INFO] Loading Excel file...")
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# === Step 2: Check if Columns Exist ===
missing_cols = [col for col in [filter_column_name, target_column_name] if col not in df.columns]
if missing_cols:
    print(f"[ERROR] Column(s) not found in Excel: {', '.join(missing_cols)}")
else:
    # === Step 3: Apply Filtering ===
    print(f"[INFO] Filtering rows where '{filter_column_name}' == '{filter_value}'...")
    filtered_df = df[df[filter_column_name].astype(str).str.strip().str.lower() == filter_value.strip().lower()]

    if filtered_df.empty:
        print(f"[WARN] No rows found with {filter_column_name} = '{filter_value}'")
    else:
        # === Step 4: Extract and Process Target Column ===
        values = filtered_df[target_column_name].dropna().astype(str).unique().tolist()
        count = len(values)
        print(f"[INFO] Found {count} unique '{target_column_name}' values after filtering.")

        formatted_values = []
        for i, val in enumerate(values):
            item = f"{prefix}{val}"
            if i < count - 1:
                item += suffix
            formatted_values.append(item)

        # === Step 5: Final Output ===
        final_string = "".join(formatted_values)
        print("\n[RESULT] Final formatted string:\n")
        print(final_string)
