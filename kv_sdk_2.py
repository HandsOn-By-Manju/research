import os
import pandas as pd
import time
from azure.identity import AzureCliCredential
from azure.core.exceptions import ClientAuthenticationError
from azure.mgmt.keyvault import KeyVaultManagementClient

# ---------------------
# 📥 Config
# ---------------------
INPUT_FILE = "keyvault_input.xlsx"
SHEET_NAME = "Sheet1"
POLICY_FILTER_VALUE = "123456"
PARTIAL_OUTPUT_FILE = "keyvault_output_partial.xlsx"
FINAL_OUTPUT_FILE = "keyvault_output.xlsx"
SAVE_EVERY = 100  # Save progress every N rows

# ---------------------
# 🔐 Azure CLI Login Check
# ---------------------
try:
    credential = AzureCliCredential()
    credential.get_token("https://management.azure.com/.default")  # proactively test login
except ClientAuthenticationError:
    print("⚠️ Azure CLI session expired or not logged in.")
    print("👉 Please run 'az login' and rerun this script.")
    exit(1)

# ---------------------
# ⏱️ Start Timer
# ---------------------
start_time = time.time()

# ---------------------
# 🧾 Load Input File and Filter
# ---------------------
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
df.columns = df.columns.str.strip()
df['Policy ID'] = df['Policy ID'].apply(lambda x: str(x).strip())
filtered_df = df[df['Policy ID'] == POLICY_FILTER_VALUE.strip()]

print(f"\n📊 Total entries in file: {len(df)}")
print(f"🔎 Filtered rows with Policy ID = '{POLICY_FILTER_VALUE}': {len(filtered_df)}")

if filtered_df.empty:
    print("⚠️ No matching rows found. Exiting.")
    exit()

# ---------------------
# 🔁 Resume Support: Load Partial Output
# ---------------------
processed_pairs = set()
results = []

if os.path.exists(PARTIAL_OUTPUT_FILE):
    partial_df = pd.read_excel(PARTIAL_OUTPUT_FILE)
    results = partial_df.to_dict(orient='records')
    for _, row in partial_df.iterrows():
        key = (row['Key Vault Name'].strip().lower(), row['Subscription ID'].strip().lower())
        processed_pairs.add(key)
    print(f"🔁 Resuming: {len(processed_pairs)} Key Vaults already processed (from {PARTIAL_OUTPUT_FILE})")

# ---------------------
# 🚀 Process Remaining Rows
# ---------------------
filtered_df = filtered_df.reset_index(drop=True)
total = len(filtered_df)
processed_count = len(processed_pairs)

for idx, (_, row) in enumerate(filtered_df.iterrows(), start=1):
    subscription_id = str(row['Subscription ID']).strip()
    keyvault_name = str(row['Key Vault Name']).strip()
    pair_key = (keyvault_name.lower(), subscription_id.lower())

    if pair_key in processed_pairs:
        continue

    print(f"\n🔍 [Processing {processed_count + 1} of {total}] Key Vault '{keyvault_name}' in subscription '{subscription_id}'")

    entry = {
        "Index": processed_count + 1,
        "Subscription ID": subscription_id,
        "Key Vault Name": keyvault_name,
        "Resource Group": "",
        "Network ACLs": "",
        "Private Endpoints": "",
        "Public Network Access": "",
        "Status": "",
        "Message": ""
    }

    try:
        client = KeyVaultManagementClient(credential, subscription_id)
        vault_found = next((v for v in client.vaults.list() if v.name.lower() == keyvault_name.lower()), None)

        if not vault_found:
            entry["Status"] = "Failed"
            entry["Message"] = "Key Vault not found"
            print(f"❌ {entry['Message']}")
        else:
            rg_parts = vault_found.id.split('/')
            resource_group_name = rg_parts[rg_parts.index('resourceGroups') + 1]

            vault_details = client.vaults.get(resource_group_name, keyvault_name)

            entry["Resource Group"] = resource_group_name
            entry["Network ACLs"] = str(vault_details.properties.network_acls.as_dict() if vault_details.properties.network_acls else "")
            entry["Private Endpoints"] = str([pe.as_dict() for pe in (vault_details.properties.private_endpoint_connections or [])])
            entry["Public Network Access"] = str(vault_details.properties.public_network_access)
            entry["Status"] = "Success"
            entry["Message"] = "Processed successfully"
            print(f"✅ {entry['Message']}")

    except ClientAuthenticationError:
        print(f"\n⛔ Azure CLI session expired during row {processed_count + 1} → Key Vault: '{keyvault_name}'")
        print("👉 Please run 'az login' and rerun the script.")
        pd.DataFrame(results).to_excel(PARTIAL_OUTPUT_FILE, index=False)
        print(f"💾 Progress saved to: {PARTIAL_OUTPUT_FILE}")
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
        print(f"💾 Partial output saved at row {processed_count} → {PARTIAL_OUTPUT_FILE}")

# ---------------------
# 📝 Final Save
# ---------------------
pd.DataFrame(results).to_excel(FINAL_OUTPUT_FILE, index=False)
print(f"\n📁 Final output saved to: {FINAL_OUTPUT_FILE}")

# ---------------------
# ⏱️ Execution Time
# ---------------------
elapsed = time.time() - start_time
hours = int(elapsed // 3600)
minutes = int((elapsed % 3600) // 60)
seconds = round(elapsed % 60, 2)
print(f"\n⏱️ Execution time: {hours} hours, {minutes} minutes, {seconds} seconds")
print("\n✅ Finished processing all remaining Key Vaults.")
