import pandas as pd
import time

# Start timer
start_time = time.time()

# File paths
input_file_path = "input_file.csv"
remediation_file_path = "remediation_file.xlsx"
subscription_details_path = "subscription_details.xlsx"
ownership_file_path = "ownership_file.xlsx"
output_excel_path = "output_file.xlsx"

unmatched_policy_log_path = "unmatched_policy_ids.txt"
unmatched_subscription_log_path = "unmatched_subscription_ids.txt"
unmatched_primary_contact_log_path = "unmatched_primary_contacts.txt"

# Load input files
print("📥 Loading input files...")
df = pd.read_csv(input_file_path)
remediation_df = pd.read_excel(remediation_file_path)
subscription_df = pd.read_excel(subscription_details_path)
ownership_df = pd.read_excel(ownership_file_path)

print(f"🧾 Columns in input file: {list(df.columns)}")

# Step 1: Rename 'Policy statement' to 'Description'
if 'Policy statement' in df.columns:
    df.rename(columns={'Policy statement': 'Description'}, inplace=True)
    print("✅ Renamed 'Policy statement' to 'Description'.")

# Step 2: Extract 'Subscription ID' and 'Subscription Name'
if 'Account' in df.columns:
    print("🔍 Extracting 'Subscription ID' and 'Subscription Name' from 'Account'...")
    extracted = df['Account'].str.extract(r'([^()]+)\s*\(\s*([^()]+)\s*\)')
    extracted.columns = ['Subscription ID', 'Subscription Name']
    extracted['Subscription ID'] = extracted['Subscription ID'].str.replace(' ', '', regex=False)
    extracted['Subscription Name'] = extracted['Subscription Name'].str.replace(' ', '', regex=False)

    account_index = df.columns.get_loc('Account')
    for i, col in enumerate(extracted.columns):
        df.insert(account_index + 1 + i, col, extracted[col])

    desired_order = ['Account', 'Subscription ID', 'Subscription Name'] + [
        col for col in df.columns if col not in ['Account', 'Subscription ID', 'Subscription Name']
    ]
    df = df[desired_order]
    print("✅ Added 'Subscription ID' and 'Subscription Name'.")

# Step 3: Clean 'Resource ID'
if 'Resource ID' in df.columns:
    print("🧹 Cleaning 'Resource ID' column...")
    df['Resource ID'] = df['Resource ID'].astype(str).apply(
        lambda x: x.rstrip('/').rsplit('/', 1)[-1] if '/' in x.rstrip('/') else x
    )
    print("✅ Cleaned 'Resource ID'.")

# Step 4: Validate and merge remediation data
print("🔍 Validating and merging 'Policy ID'...")
if 'Policy ID' in df.columns and 'Policy ID' in remediation_df.columns:
    input_policy_ids = set(df['Policy ID'].dropna().unique())
    remediation_policy_ids = set(remediation_df['Policy ID'].dropna().unique())
    unmatched_policy_ids = sorted(list(input_policy_ids - remediation_policy_ids))

    if unmatched_policy_ids:
        with open(unmatched_policy_log_path, 'w') as f:
            f.write("Unmatched Policy IDs:\n")
            for pid in unmatched_policy_ids:
                f.write(str(pid) + "\n")
        print(f"⚠️ {len(unmatched_policy_ids)} unmatched Policy IDs logged to {unmatched_policy_log_path}")

    remediation_subset = remediation_df[['Policy ID', 'Policy Statement', 'Policy Remediation']]
    df = df.merge(remediation_subset, on='Policy ID', how='left')
    print("✅ Merged 'Policy Statement' and 'Policy Remediation'.")

# Step 5: Validate and merge subscription details
print("🔍 Validating and merging 'Subscription ID'...")
if 'Subscription ID' in df.columns and 'Subscription ID' in subscription_df.columns:
    input_sub_ids = set(df['Subscription ID'].dropna().unique())
    subscription_sub_ids = set(subscription_df['Subscription ID'].dropna().unique())
    unmatched_sub_ids = sorted(list(input_sub_ids - subscription_sub_ids))

    if unmatched_sub_ids:
        with open(unmatched_subscription_log_path, 'w') as f:
            f.write("Unmatched Subscription IDs:\n")
            for sid in unmatched_sub_ids:
                f.write(str(sid) + "\n")
        print(f"⚠️ {len(unmatched_sub_ids)} unmatched Subscription IDs logged to {unmatched_subscription_log_path}")

    subscription_subset = subscription_df[['Subscription ID', 'Environment', 'BU', 'Primary Contact']]
    df = df.merge(subscription_subset, on='Subscription ID', how='left')
    print("✅ Merged 'Environment', 'BU', and 'Primary Contact'.")

# Step 6: Validate and merge ownership details
print("🔍 Validating and merging 'Primary Contact'...")
if 'Primary Contact' in df.columns and 'Primary Contact' in ownership_df.columns:
    input_contacts = set(df['Primary Contact'].dropna().unique())
    ownership_contacts = set(ownership_df['Primary Contact'].dropna().unique())
    unmatched_contacts = sorted(list(input_contacts - ownership_contacts))

    if unmatched_contacts:
        with open(unmatched_primary_contact_log_path, 'w') as f:
            f.write("Unmatched Primary Contacts:\n")
            for contact in unmatched_contacts:
                f.write(str(contact) + "\n")
        print(f"⚠️ {len(unmatched_contacts)} unmatched Primary Contacts logged to {unmatched_primary_contact_log_path}")

    ownership_cols = [
        'Primary Contact',
        'Manager / Sr Manager / Director / Sr Director',
        'Sr Director / VP',
        'VP / SVP / CVP'
    ]
    df = df.merge(ownership_df[ownership_cols], on='Primary Contact', how='left')
    print("✅ Merged ownership hierarchy columns.")

# Step 7: Remove unwanted columns
columns_to_remove = ['Column1', 'Column2', 'Column3']  # ⬅️ Replace with actual columns to remove
print(f"🧽 Dropping columns if present: {columns_to_remove}")
existing_columns_to_remove = [col for col in columns_to_remove if col in df.columns]
df.drop(columns=existing_columns_to_remove, inplace=True)
print(f"✅ Dropped columns: {existing_columns_to_remove}")

# Final column list
print(f"\n📋 Final columns in the Excel file:\n{list(df.columns)}")

# Step 8: Save Excel output
df.to_excel(output_excel_path, index=False)
print(f"💾 Saved final Excel file to: {output_excel_path}")

# Timer
end_time = time.time()
elapsed = end_time - start_time

# Execution time
if elapsed < 60:
    print(f"\n⏱️ Execution Time: {elapsed:.2f} seconds")
elif elapsed < 3600:
    print(f"\n⏱️ Execution Time: {int(elapsed // 60)} minutes {elapsed % 60:.2f} seconds")
else:
    hours = int(elapsed // 3600)
    minutes = int((elapsed % 3600) // 60)
    seconds = elapsed % 60
    print(f"\n⏱️ Execution Time: {hours} hours {minutes} minutes {seconds:.2f} seconds")
