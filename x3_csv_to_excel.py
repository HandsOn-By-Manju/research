import pandas as pd
import time

# Start timer
start_time = time.time()

# File paths
input_file_path = "input_file.csv"
remediation_file_path = "remediation_file.xlsx"
subscription_details_path = "subscription_details.xlsx"
output_excel_path = "output_file.xlsx"
unmatched_policy_log_path = "unmatched_policy_ids.txt"
unmatched_subscription_log_path = "unmatched_subscription_ids.txt"

# Load files
df = pd.read_csv(input_file_path)
remediation_df = pd.read_excel(remediation_file_path)
subscription_df = pd.read_excel(subscription_details_path)

# Step 1: Rename 'Policy statement' to 'Description' in input
if 'Policy statement' in df.columns:
    df.rename(columns={'Policy statement': 'Description'}, inplace=True)

# Step 2: Process 'Account' column
if 'Account' in df.columns:
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

# Step 3: Clean 'Resource ID'
if 'Resource ID' in df.columns:
    df['Resource ID'] = df['Resource ID'].astype(str).apply(
        lambda x: x.rstrip('/').rsplit('/', 1)[-1] if '/' in x.rstrip('/') else x
    )

# Step 4: Validate Policy ID and merge remediation file
if 'Policy ID' in df.columns and 'Policy ID' in remediation_df.columns:
    input_policy_ids = set(df['Policy ID'].dropna().unique())
    remediation_policy_ids = set(remediation_df['Policy ID'].dropna().unique())
    unmatched_policy_ids = sorted(list(input_policy_ids - remediation_policy_ids))

    if unmatched_policy_ids:
        print(f"⚠️ {len(unmatched_policy_ids)} unmatched Policy IDs logged to {unmatched_policy_log_path}")
        with open(unmatched_policy_log_path, 'w') as f:
            f.write("Unmatched Policy IDs:\n")
            for pid in unmatched_policy_ids:
                f.write(str(pid) + "\n")
    else:
        print("✅ All Policy IDs matched.")

    remediation_subset = remediation_df[['Policy ID', 'Policy Statement', 'Policy Remediation']]
    df = df.merge(remediation_subset, on='Policy ID', how='left')

# Step 5: Validate Subscription ID and merge subscription details
if 'Subscription ID' in df.columns and 'Subscription ID' in subscription_df.columns:
    input_sub_ids = set(df['Subscription ID'].dropna().unique())
    subscription_sub_ids = set(subscription_df['Subscription ID'].dropna().unique())
    unmatched_sub_ids = sorted(list(input_sub_ids - subscription_sub_ids))

    if unmatched_sub_ids:
        print(f"⚠️ {len(unmatched_sub_ids)} unmatched Subscription IDs logged to {unmatched_subscription_log_path}")
        with open(unmatched_subscription_log_path, 'w') as f:
            f.write("Unmatched Subscription IDs:\n")
            for sid in unmatched_sub_ids:
                f.write(str(sid) + "\n")
    else:
        print("✅ All Subscription IDs matched.")

    subscription_subset = subscription_df[['Subscription ID', 'Environment', 'BU', 'Primary Contact']]
    df = df.merge(subscription_subset, on='Subscription ID', how='left')

# Step 6: Save final Excel file
df.to_excel(output_excel_path, index=False)

# Execution timer
end_time = time.time()
elapsed = end_time - start_time

# Time display
if elapsed < 60:
    print(f"\n⏱️ Execution Time: {elapsed:.2f} seconds")
elif elapsed < 3600:
    print(f"\n⏱️ Execution Time: {int(elapsed // 60)} minutes {elapsed % 60:.2f} seconds")
else:
    hours = int(elapsed // 3600)
    minutes = int((elapsed % 3600) // 60)
    seconds = elapsed % 60
    print(f"\n⏱️ Execution Time: {hours} hours {minutes} minutes {seconds:.2f} seconds")

# Completion messages
print(f"✅ Final Excel file saved to: {output_excel_path}")
if 'unmatched_policy_ids' in locals() and unmatched_policy_ids:
    print(f"📄 Unmatched Policy IDs in: {unmatched_policy_log_path}")
if 'unmatched_sub_ids' in locals() and unmatched_sub_ids:
    print(f"📄 Unmatched Subscription IDs in: {unmatched_subscription_log_path}")
print(f"📊 Final columns: {list(df.columns)}")
