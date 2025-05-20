import time
from azure.identity import AzureCliCredential
from azure.mgmt.keyvault import KeyVaultManagementClient

# ---------------------
# 🔧 Configuration
# ---------------------
SUBSCRIPTION_ID = "your-subscription-id"   # Replace with your Subscription ID
KEYVAULT_NAME = "your-keyvault-name"       # Replace with your Key Vault name

# ---------------------
# 🔐 Authenticate using Azure CLI Login
# ---------------------
credential = AzureCliCredential()
client = KeyVaultManagementClient(credential, SUBSCRIPTION_ID)

# ---------------------
# 🔍 Search for the Key Vault by name
# ---------------------
print(f"🔎 Searching for Key Vault '{KEYVAULT_NAME}' in subscription '{SUBSCRIPTION_ID}'...")
start_time = time.time()

keyvault_found = None
for vault in client.vaults.list():
    if vault.name.lower() == KEYVAULT_NAME.lower():
        keyvault_found = vault
        break

if not keyvault_found:
    print(f"❌ Key Vault '{KEYVAULT_NAME}' not found in subscription '{SUBSCRIPTION_ID}'.")
else:
    # Extract resource group from the ID
    resource_id_parts = keyvault_found.id.split('/')
    rg_index = resource_id_parts.index('resourceGroups') + 1
    resource_group_name = resource_id_parts[rg_index]

    print(f"✅ Found Key Vault in resource group: {resource_group_name}")
    
    # ---------------------
    # 📥 Fetch Full Vault Details
    # ---------------------
    vault_details = client.vaults.get(resource_group_name, KEYVAULT_NAME)

    network_acls = vault_details.properties.network_acls.as_dict() if vault_details.properties.network_acls else None
    private_endpoints = [pe.as_dict() for pe in (vault_details.properties.private_endpoint_connections or [])]
    public_network_access = vault_details.properties.public_network_access

    # ---------------------
    # 🖨️ Output
    # ---------------------
    print("\n🌐 Network ACLs:")
    print(network_acls)

    print("\n🔗 Private Endpoint Connections:")
    print(private_endpoints)

    print("\n🌍 Public Network Access:")
    print(public_network_access)

    print(f"\n⏱️ Completed in {round(time.time() - start_time, 2)} seconds.")
