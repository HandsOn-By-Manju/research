import time
from azure.identity import AzureCliCredential
from azure.mgmt.keyvault import KeyVaultManagementClient

# ---------------------
# 🔧 Configuration
# ---------------------
SUBSCRIPTION_ID = "abc"                 # Replace with your subscription ID
KEYVAULT_NAME = "your-keyvault-name"    # Replace with your Key Vault name

# ---------------------
# 🔐 Authenticate
# ---------------------
credential = AzureCliCredential()
client = KeyVaultManagementClient(credential, SUBSCRIPTION_ID)

# ---------------------
# 🔍 Find Key Vault by Name
# ---------------------
start_time = time.time()
print(f"🔍 Searching for Key Vault: {KEYVAULT_NAME} in subscription {SUBSCRIPTION_ID}...")

keyvault_found = None

for vault in client.vaults.list():
    if vault.name.lower() == KEYVAULT_NAME.lower():
        keyvault_found = vault
        break

if not keyvault_found:
    print(f"❌ Key Vault '{KEYVAULT_NAME}' not found in subscription '{SUBSCRIPTION_ID}'.")
else:
    # ---------------------
    # 📥 Extract and Display Details
    # ---------------------
    network_acls = keyvault_found.properties.network_acls.as_dict() if keyvault_found.properties.network_acls else None
    private_endpoints = [pe.as_dict() for pe in (keyvault_found.properties.private_endpoint_connections or [])]
    public_network_access = keyvault_found.properties.public_network_access

    print("\n🌐 Network ACLs:")
    print(network_acls)

    print("\n🔗 Private Endpoint Connections:")
    print(private_endpoints)

    print("\n🌍 Public Network Access:")
    print(public_network_access)

    print(f"\n✅ Found in resource group: {keyvault_found.resource_group_name}")
    print(f"✅ Done in {round(time.time() - start_time, 2)} seconds.")
