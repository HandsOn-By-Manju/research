import pandas as pd
import time
from azure.identity import AzureCliCredential
from azure.mgmt.keyvault import KeyVaultManagementClient

# ---------------------
# 📥 Input Config
# ---------------------
INPUT_FILE = "keyvault_input.xlsx"  # or CSV
SHEET_NAME = "Sheet1"
POLICY_FILTER_VALUE = "abc"

# ---------------------
# 🔐 Azure CLI Auth
# ---------------------
credential = AzureCliCredential()

# ---------------------
# 🧾 Load and Filter Input File
# ---------------------
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)  # or use read_csv()
filtered_df = df[df["Policy ID"] == POLICY_FILTER_VALUE]

print(f"\n📊 Total entries in file: {len(df)}")
print(f"🔎 Filtered rows with Policy ID = '{POLICY_FILTER_VALUE}': {len(filtered_df)}")

if filtered_df.empty:
    print("⚠️ No rows match the given Policy ID filter. Exiting.")
    exit()

# ---------------------
# 🚀 Process Each Filtered Row
# ---------------------
for idx, row in filtered_df.iterrows():
    subscription_id = row['Subscription ID']
    keyvault_name = row['Key Vault Name']

    print(f"\n🔍 [{idx+1}] Processing Key Vault '{keyvault_name}' in subscription '{subscription_id}'...")

    try:
        client = KeyVaultManagementClient(credential, subscription_id)
        vault_found = None

        # Search Key Vault by name
        for vault in client.vaults.list():
            if vault.name.lower() == keyvault_name.lower():
                vault_found = vault
                break

        if not vault_found:
            print(f"❌ Key Vault '{keyvault_name}' not found.")
            continue

        # Extract resource group from resource ID
        rg_index = vault_found.id.split('/').index('resourceGroups') + 1
        resource_group_name = vault_found.id.split('/')[rg_index]

        # Get full vault details
        vault_details = client.vaults.get(resource_group_name, keyvault_name)

        network_acls = vault_details.properties.network_acls.as_dict() if vault_details.properties.network_acls else None
        private_endpoints = [pe.as_dict() for pe in (vault_details.properties.private_endpoint_connections or [])]
        public_network_access = vault_details.properties.public_network_access

        # Print results
        print(f"✅ Resource Group: {resource_group_name}")
        print("🌐 Network ACLs:", network_acls)
        print("🔗 Private Endpoint Connections:", private_endpoints)
        print("🌍 Public Network Access:", public_network_access)

    except Exception as e:
        print(f"❗ Error processing '{keyvault_name}': {e}")

print("\n✅ Finished processing all filtered Key Vaults.")
