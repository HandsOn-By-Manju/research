import pandas as pd
import time
from azure.identity import AzureCliCredential
from azure.mgmt.keyvault import KeyVaultManagementClient

# ---------------------
# 📥 Input Config
# ---------------------
INPUT_FILE = "keyvault_input.xlsx"
SHEET_NAME = "Sheet1"
POLICY_FILTER_VALUE = "123456"  # Replace with your numeric/string Policy ID

# ---------------------
# 🔐 Azure CLI Auth
# ---------------------
credential = AzureCliCredential()

# ---------------------
# ⏱️ Start Timer
# ---------------------
start_time = time.time()

# ---------------------
# 🧾 Load and Filter Data
# ---------------------
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
df.columns = df.columns.str.strip()

# Normalize Policy ID to string and strip
df['Policy ID'] = df['Policy ID'].apply(lambda x: str(x).strip())

# Apply filter
filtered_df = df[df['Policy ID'] == POLICY_FILTER_VALUE.strip()]

# Display counts
print(f"\n📊 Total entries in file: {len(df)}")
print(f"🔎 Filtered rows with Policy ID = '{POLICY_FILTER_VALUE}': {len(filtered_df)}")

# Exit if no matches
if filtered_df.empty:
    print("⚠️ No matching rows found. Exiting.")
    exit()

# ---------------------
# 🚀 Process Each Filtered Row
# ---------------------
for count, (_, row) in enumerate(filtered_df.iterrows(), start=1):
    subscription_id = str(row['Subscription ID']).strip()
    keyvault_name = str(row['Key Vault Name']).strip()

    print(f"\n🔍 [{count}] Processing Key Vault '{keyvault_name}' in subscription '{subscription_id}'...")

    try:
        client = KeyVaultManagementClient(credential, subscription_id)

        # Find the Key Vault by name
        vault_found = next((v for v in client.vaults.list() if v.name.lower() == keyvault_name.lower()), None)

        if not vault_found:
            print(f"❌ Key Vault '{keyvault_name}' not found.")
            continue

        # Extract resource group name from resource ID
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

# ---------------------
# ⏱️ End Timer and Display Duration
# ---------------------
end_time = time.time()
elapsed = end_time - start_time
if elapsed < 60:
    print(f"\n⏱️ Execution time: {round(elapsed, 2)} seconds.")
else:
    mins = int(elapsed // 60)
    secs = round(elapsed % 60, 2)
    print(f"\n⏱️ Execution time: {mins} minutes, {secs} seconds.")

print("\n✅ Finished processing all filtered Key Vaults.")
