import pandas as pd
import time
from azure.identity import AzureCliCredential
from azure.mgmt.keyvault import KeyVaultManagementClient

# ---------------------
# 📥 Input Config
# ---------------------
INPUT_FILE = "keyvault_input.xlsx"
SHEET_NAME = "Sheet1"
POLICY_FILTER_VALUE = "123456"
OUTPUT_FILE = "keyvault_output.xlsx"

# ---------------------
# 🔐 Azure CLI Auth
# ---------------------
credential = AzureCliCredential()

# ---------------------
# ⏱️ Start Timer
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
print(f"🔎 Filtered rows with Policy ID = '{POLICY_FILTER_VALUE}': {len(filtered_df)}")

if filtered_df.empty:
    print("⚠️ No matching rows found. Exiting.")
    exit()

# ---------------------
# 📦 Results Storage
# ---------------------
results = []

# ---------------------
# 🚀 Process Each Entry
# ---------------------
for count, (_, row) in enumerate(filtered_df.iterrows(), start=1):
    subscription_id = str(row['Subscription ID']).strip()
    keyvault_name = str(row['Key Vault Name']).strip()

    print(f"\n🔍 [{count}] Processing Key Vault '{keyvault_name}' in subscription '{subscription_id}'...")
    
    entry = {
        "Index": count,
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
            results.append(entry)
            continue

        # Extract Resource Group
        rg_parts = vault_found.id.split('/')
        resource_group_name = rg_parts[rg_parts.index('resourceGroups') + 1]

        # Get full details
        vault_details = client.vaults.get(resource_group_name, keyvault_name)

        entry["Resource Group"] = resource_group_name
        entry["Network ACLs"] = str(vault_details.properties.network_acls.as_dict() if vault_details.properties.network_acls else "")
        entry["Private Endpoints"] = str([pe.as_dict() for pe in (vault_details.properties.private_endpoint_connections or [])])
        entry["Public Network Access"] = str(vault_details.properties.public_network_access)
        entry["Status"] = "Success"
        entry["Message"] = "Processed successfully"

        print(f"✅ {entry['Message']}")
    except Exception as e:
        entry["Status"] = "Failed"
        entry["Message"] = str(e)
        print(f"❗ Error: {e}")

    results.append(entry)

# ---------------------
# 📝 Write to Excel
# ---------------------
output_df = pd.DataFrame(results)
output_df.to_excel(OUTPUT_FILE, index=False)
print(f"\n📁 Output saved to: {OUTPUT_FILE}")

# ---------------------
# ⏱️ Execution Time
# ---------------------
elapsed = time.time() - start_time
if elapsed < 60:
    print(f"⏱️ Execution time: {round(elapsed, 2)} seconds.")
else:
    mins = int(elapsed // 60)
    secs = round(elapsed % 60, 2)
    print(f"⏱️ Execution time: {mins} minutes, {secs} seconds.")

print("\n✅ Finished processing all filtered Key Vaults.")
