import pandas as pd

# === Configuration ===
excel_file_path = "your_excel_file.xlsx"       # <- Update your file path
sheet_name = 0                                  # <- Change if specific sheet needed
filter_column_name = "Subscription ID"          # <- Configurable column name
prefix = "prefix_data_for_each_subscriptions"   # <- Configurable prefix
suffix = "suffix_data_for_each_subscriptions"   # <- Configurable suffix

# === Step 1: Load Excel File ===
print("[INFO] Loading Excel file...")
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# === Step 2: Check if Column Exists ===
if filter_column_name not in df.columns:
    print(f"[ERROR] Column '{filter_column_name}' not found in the Excel file.")
else:
    # === Step 3: Extract and Clean Values ===
    values = df[filter_column_name].dropna().astype(str).unique().tolist()
    count = len(values)
    print(f"[INFO] Found {count} unique values in column '{filter_column_name}'.")

    # === Step 4: Apply Prefix/Suffix Logic ===
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
