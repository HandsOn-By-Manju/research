import pandas as pd

# === Configuration ===
excel_file_path = "your_excel_file.xlsx"               # <- Update your Excel file path
sheet_name = 0                                          # <- Sheet index or name
filter_column_name = "Severity"                        # <- Column to filter
target_column_name = "Subscription ID"                 # <- Column from which to extract values

# Per-item prefix/suffix
prefix = "prefix_data_for_each_subscriptions"
suffix = "suffix_data_for_each_subscriptions"

# Final prefix is the BASE URL
final_base_url = "https://example.com/trigger/"         # <- Final prefix becomes base URL
final_suffix = "_FINAL_SUFFIX"                          # <- Optional suffix for full string

# === Step 1: Load Excel File ===
print("[INFO] Loading Excel file...")
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# === Step 2: Check Required Columns ===
missing_cols = [col for col in [filter_column_name, target_column_name] if col not in df.columns]
if missing_cols:
    print(f"[ERROR] Column(s) not found in Excel: {', '.join(missing_cols)}")
else:
    # === Step 3: Get All Unique Values in Filter Column ===
    unique_filter_values = df[filter_column_name].dropna().astype(str).unique().tolist()
    print(f"[INFO] Found {len(unique_filter_values)} unique values in '{filter_column_name}' column.\n")

    for filter_val in unique_filter_values:
        print(f"[INFO] Processing for '{filter_column_name}' = '{filter_val}'...")

        # Filter rows for this filter_val
        filtered_df = df[df[filter_column_name].astype(str).str.strip().str.lower() == filter_val.strip().lower()]

        if filtered_df.empty:
            print(f"[WARN] No rows found for value: {filter_val}")
            continue

        # Extract target column values
        values = filtered_df[target_column_name].dropna().astype(str).unique().tolist()
        count = len(values)
        print(f"  -> Found {count} unique '{target_column_name}' values.")

        # Format each value
        formatted_items = []
        for i, val in enumerate(values):
            item = f"{prefix}{val}"
            if i < count - 1:
                item += suffix
            formatted_items.append(item)

        combined_string = "".join(formatted_items)

        # Build final clickable URL
        final_url = f"{final_base_url}{combined_string}{final_suffix}"

        # Print final output
        print(f"  -> Final URL:\n     {final_url}\n")
