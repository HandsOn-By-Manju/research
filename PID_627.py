import os
import pandas as pd
import time
from azure.identity import AzureCliCredential
from azure.core.exceptions import ClientAuthenticationError
from azure.mgmt.storage import StorageManagementClient

# ---------------------
# 📥 Config
# ---------------------
INPUT_FILE = "storage_input.xlsx"
SHEET_NAME = "Sheet1"
POLICY_FILTER_VALUE = "123456"
PARTIAL_OUTPUT_FILE = "storage_pe_missing_partial.xlsx"
FINAL_OUTPUT_FILE = "storage_pe_missing_output.xlsx"
SAVE_EVERY = 100

# ---------------------
# 🔐 Azure CLI Login Check
# ---------------------
try:
    credential = AzureCliCredential()
    credential.get_token("https://management.azure.com/.default")
except ClientAuthenticationError:
    print("⚠️ Azure CLI session expired or not logged in.")
    print("👉 Please run 'az login' and rerun this script.")
    exit(1)

# ---------------------
# ⏱️ Timer Start
# ---------------------
start_time = time.time()

# ---------------------
# 🧾 Load and Filter Input
# ---------------------
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
df.columns = df.columns.str.strip()
df['Policy ID'] = df['Policy ID'].apply(lambda x: str(x).strip())
filtered_df = df[df['Policy ID'] == POLICY_FILTER_VALUE.strip()]

print(f"\n📊 Total entries: {len(df)}")
print(f"🔎 Matching Policy ID '{POLICY_FILTER_VALUE}': {len(filtered_df)}")

if filtered_df.empty:
    print("⚠️ No matching rows. Exiting.")
    exit(1)

# ---------------------
# 🔁 Resume Support
# ---------------------
processed_pairs = set()
results = []

if os.path.exists(PARTIAL_OUTPUT_FILE):
    partial_df = pd.read_excel(PARTIAL_OUTPUT_FILE)
    results = partial_df.to_dict(orient='records')
    for _, row in partial_df.iterrows():
        key = (row['Storage Account Name'].strip().lower(), row['Subscription ID'].strip().lower())
        processed_pairs.add(key)
    print(f"🔁 Resuming from {len(processed_pairs)} processed storage accounts.")

# ---------------------
# 🚀 Process Each Storage Account
# ---------------------
filtered_df = filtered_df.reset_index(drop=True)
total = len(filtered_df)
processed_count = len(processed_pairs)

for idx, (_, row) in enumerate(filtered_df.iterrows(), start=1):
    sub_id = str(row['Subscription ID']).strip()
    sa_name = str(row['Storage Account Name']).strip()
    pair_key = (sa_name.lower(), sub_id.lower())

    if pair_key in processed_pairs:
        continue

    print(f"\n🔍 [Processing {processed_count + 1} of {total}] '{sa_name}' in subscription '{sub_id}'")

    entry = {
        "Index": processed_count + 1,
        "Subscription ID": sub_id,
        "Storage Account Name": sa_name,
        "Resource Group": "",
        "Private Endpoint Configured?": "",
        "Private Endpoint Names": "",
        "Status": "",
        "Message": ""
    }

    try:
        client = StorageManagementClient(credential, sub_id)
        accounts = list(client.storage_accounts.list())

        sa = next((a for a in accounts if a.name.lower() == sa_name.lower()), None)
        if not sa:
            entry["Status"] = "Failed"
            entry["Message"] = "Storage Account not found"
            print(f"❌ {entry['Message']}")
        else:
            rg_name = sa.id.split("/")[sa.id.split("/").index("resourceGroups") + 1]
            entry["Resource Group"] = rg_name

            pe_connections = client.private_endpoint_connections.list(rg_name, sa_name)
            pe_list = list(pe_connections)
            pe_names = [pe.name for pe in pe_list]

            if pe_list:
                entry["Private Endpoint Configured?"] = "Yes 🔒"
                entry["Private Endpoint Names"] = ", ".join(pe_names)
            else:
                entry["Private Endpoint Configured?"] = "No 🌐"
                entry["Private Endpoint Names"] = ""

            entry["Status"] = "Success"
            entry["Message"] = "Processed successfully"
            print(f"✅ {entry['Message']}")

    except ClientAuthenticationError:
        print(f"\n⛔ CLI session expired at row {processed_count + 1} → '{sa_name}'")
        pd.DataFrame(results).to_excel(PARTIAL_OUTPUT_FILE, index=False)
        print(f"💾 Partial saved → {PARTIAL_OUTPUT_FILE}")
        print("👉 Run 'az login' and rerun this script.")
        exit(1)
    except Exception as e:
        entry["Status"] = "Failed"
        entry["Message"] = str(e)
        print(f"❗ Error: {e}")

    results.append(entry)
    processed_pairs.add(pair_key)
    processed_count += 1

    if processed_count % SAVE_EVERY == 0:
        pd.DataFrame(results).to_excel(PARTIAL_OUTPUT_FILE, index=False)
        print(f"💾 Saved {processed_count} rows → {PARTIAL_OUTPUT_FILE}")

# ---------------------
# ✅ Final Save
# ---------------------
pd.DataFrame(results).to_excel(FINAL_OUTPUT_FILE, index=False)
print(f"\n📁 Final output saved to: {FINAL_OUTPUT_FILE}")

# ---------------------
# ⏱️ Execution Time
# ---------------------
elapsed = time.time() - start_time
h, m, s = int(elapsed // 3600), int((elapsed % 3600) // 60), round(elapsed % 60, 2)
print(f"\n⏱️ Execution time: {h} hours, {m} minutes, {s} seconds")
print("✅ All storage accounts processed.")
