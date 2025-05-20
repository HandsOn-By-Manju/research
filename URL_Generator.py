import pandas as pd

# === Configuration ===
excel_file_path = "your_excel_file.xlsx"  # Update your Excel file path
sheet_name = 0  # Change if you're using a specific sheet by name or index
subscription_column = "Subscription ID"
prefix = "prefix_data_for_each_subscriptions"
suffix = "suffix_data_for_each_subscriptions"

# === Step 1: Load Excel File ===
print("[INFO] Loading Excel file...")
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# === Step 2: Check Column Existence ===
if subscription_column not in df.columns:
    print(f"[ERROR] Column '{subscription_column}' not found in the file.")
else:
    # === Step 3: Extract and Clean Subscription IDs ===
    subscription_ids = df[subscription_column].dropna().astype(str).unique().tolist()
    count = len(subscription_ids)
    print(f"[INFO] Found {count} unique Subscription IDs.")

    # === Step 4: Apply Prefix and Suffix ===
    print("[INFO] Applying prefix and suffix formatting...")
    formatted_subs = []
    for i, sub_id in enumerate(subscription_ids):
        full_sub = f"{prefix}{sub_id}"
        if i < count - 1:
            full_sub += suffix
        formatted_subs.append(full_sub)

    # === Step 5: Print Final String ===
    final_string = "".join(formatted_subs)
    print("\n[RESULT] Final combined string:\n")
    print(final_string)
