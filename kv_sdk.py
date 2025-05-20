import pandas as pd
import time
from azure.identity import AzureCliCredential
from azure.mgmt.keyvault import KeyVaultManagementClient

# ---------------------
# 📥 Input Config
# ---------------------
INPUT_FILE = "keyvault_input.xlsx"   # or CSV
SHEET_NAME = "Sheet1"
POLICY_FILTER_VALUE = "123456"       # <- Numeric Policy ID as string

# ---------------------
# 🔐 Azure CLI Auth
# ---------------------
credential = AzureCliCredential()

# ---------------------
# 🧾 Load and Filter Data
# ---------------------
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
df.columns = df.columns.str.strip()

# Clean and convert Policy ID to string for safe comparison
df['Policy ID'] = df['Policy ID'].apply(lambda x: str(x).strip())

# Filter by exact match (as string)
filtered_df = df[df['Policy ID'] == POLICY_FILTER_VALUE.strip()]

# Display counts
print(f"\n📊 Total entries in file: {len(df)}")
print(f"🔎 Filtered rows with Policy ID = '{POLICY_FILTER_VALUE}': {len(filtered_df)}")

# Exit if no rows match
if filtered_df.empty:
    print("⚠️ No matching rows found. Exiting.")
    exit()

# ---------------------
# 🚀 Process Each Filtered Row
# ---------------------
for idx, row in filtered_df.iterrows():
    subscription_id = str(row['Subscription ID']).strip()
    keyvault_name = str(row['Key Vault Name']).strip()

    print(f"\n🔍 [{idx+1}] Processing Key Vault '{keyvault_name}' in subscription '{subscription_id}'...")

    try:
        client = KeyVaultManagementClient(credential, subscription_id)

        # Find the Key Vault by name
        vault_found = next((v for v in client.vaults.list() if v.name.lower() == keyvault_name.lower()), None)

        if not vault_found:
            print(f"❌ Key Vault '{keyvault_name}' not found.")
            continue

        # Extract resource group from resource ID
        rg_parts = vault_found.id.split('/')
        resource_group_name = rg_parts[rg_parts.index('resourceGroups') + 1]

        # Get full vault details
        vault_details = client.vaults.get(resource_group_name, keyvault_name)

        network_acls = vault_details.properties.network_acls.as_dict() if vault_details.properties.network_acls else None
        private_endpoints = [pe.as_dict() for pe in (vault_details.properties.private_endpoint_connections or [])]
        public_network_access = vault_details.properties.public_network_access

        # Output results
        print(f"✅ Resource Group: {resource_group_name}")
        print("🌐 Network ACLs:", network_acls)
        print("🔗 Private Endpoint Connections:", private_endpoints)
        print("🌍 Public Network Access:", public_network_access)

    except Exception as e:
        print(f"❗ Error processing '{keyvault_name}': {e}")

print("\n✅ Finished processing all filtered Key Vaults.")
