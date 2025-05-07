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

# Rename 'Policy statement' to 'Description'
if 'Policy statement' in df.columns:
    df.rename(columns={'Policy statement': 'Description'}, inplace=True)
    print("✅ Renamed 'Policy statement' to 'Description'.")

# Extract Subscription ID and Name
if 'Account' in df.columns:
    print("🔍 Extracting 'Subscription ID' and 'Subscription Name' from 'Account'...")
    extracted = df['Account'].str.extract(r'([^()]+)\s*\(\s*([^()]+)\s*\)')
    extracted.columns = ['Subscription ID', 'Subscription Name']
    extracted['Subscription ID'] = extracted['Subscription ID'].str.replace(' ', '', regex=False)
    extracted['Subscription Name'] = extracted['Subscription Name'].str.replace(' ', '', regex=False)
    account_index = df.columns.get_loc('Account')
    for i, col in enumerate(extracted.columns):
        df.insert(account_index + 1 + i, col, extracted[col])
    print("✅ Added 'Subscription ID' and 'Subscription Name'.")

# Clean 'Resource ID'
if 'Resource ID' in df.columns:
    print("🧹 Cleaning 'Resource ID' column...")
    df['Resource ID'] = df['Resource ID'].astype(str).apply(
        lambda x: x.rstrip('/').rsplit('/', 1)[-1] if '/' in x.rstrip('/') else x
    )
    print("✅ Cleaned 'Resource ID'.")

# Merge remediation data
print("🔍 Validating and merging 'Policy ID'...")
if 'Policy ID' in df.columns and 'Policy ID' in remediation_df.columns:
    unmatched = set(df['Policy ID'].dropna()) - set(remediation_df['Policy ID'].dropna())
    if unmatched:
        with open(unmatched_policy_log_path, 'w') as f:
            for item in unmatched:
                f.write(f"{item}\n")
        print(f"⚠️ Unmatched Policy IDs written to {unmatched_policy_log_path}")
    df = df.merge(remediation_df[['Policy ID', 'Policy Statement', 'Policy Remediation']], on='Policy ID', how='left')
    print("✅ Merged remediation columns.")

# Merge subscription details
print("🔍 Validating and merging 'Subscription ID'...")
if 'Subscription ID' in df.columns and 'Subscription ID' in subscription_df.columns:
    unmatched = set(df['Subscription ID'].dropna()) - set(subscription_df['Subscription ID'].dropna())
    if unmatched:
        with open(unmatched_subscription_log_path, 'w') as f:
            for item in unmatched:
                f.write(f"{item}\n")
        print(f"⚠️ Unmatched Subscription IDs written to {unmatched_subscription_log_path}")
    df = df.merge(subscription_df[['Subscription ID', 'Environment', 'BU', 'Primary Contact']], on='Subscription ID', how='left')
    print("✅ Merged subscription details.")

# Merge ownership data
print("🔍 Validating and merging 'Primary Contact'...")
if 'Primary Contact' in df.columns and 'Primary Contact' in ownership_df.columns:
    unmatched = set(df['Primary Contact'].dropna()) - set(ownership_df['Primary Contact'].dropna())
    if unmatched:
        with open(unmatched_primary_contact_log_path, 'w') as f:
            for item in unmatched:
                f.write(f"{item}\n")
        print(f"⚠️ Unmatched Primary Contacts written to {unmatched_primary_contact_log_path}")
    df = df.merge(
        ownership_df[
            ['Primary Contact', 
             'Manager / Sr Manager / Director / Sr Director', 
             'Sr Director / VP', 
             'VP / SVP / CVP']
        ], on='Primary Contact', how='left'
    )
    print("✅ Merged ownership data.")

# Remove unwanted columns
columns_to_remove = ['Column1', 'Column2', 'Column3']  # 🔁 Update as needed
existing = [col for col in columns_to_remove if col in df.columns]
df.drop(columns=existing, inplace=True)
print(f"🧽 Dropped columns: {existing}")

# Rename columns
df.rename(columns={
    'Cloud provider': 'Cloud Provider',
    'Environment_y': 'Environment'
}, inplace=True)
print("🔤 Renamed columns: 'Cloud provider' → 'Cloud Provider', 'Environment_y' → 'Environment'")

# Rearrange columns
desired_order = [
    'Cloud Provider', 'Subscription ID', 'Subscription Name', 'Region', 'Service', 'Resource ID',
    'Policy ID', 'Description', 'Severity', 'Policy Statement', 'Policy Remediation', 'Finding',
    'Environment', 'Primary Contact',
    'Manager / Sr Manager / Director / Sr Director', 'Sr Director / VP', 'VP / SVP / CVP', 'BU'
]
missing = [col for col in desired_order if col not in df.columns]
if missing:
    print(f"⚠️ Missing columns not included in output: {missing}")
df = df[[col for col in desired_order if col in df.columns]]
print("📐 Columns rearranged in specified order.")

# Save final Excel
df.to_excel(output_excel_path, index=False)
print(f"💾 Saved final Excel file to: {output_excel_path}")

# Show execution time
elapsed = time.time() - start_time
if elapsed < 60:
    print(f"\n⏱️ Execution Time: {elapsed:.2f} seconds")
elif elapsed < 3600:
    print(f"\n⏱️ Execution Time: {elapsed//60:.0f} minutes {elapsed%60:.2f} seconds")
else:
    h, rem = divmod(elapsed, 3600)
    m, s = divmod(rem, 60)
    print(f"\n⏱️ Execution Time: {int(h)}h {int(m)}m {s:.2f}s")
