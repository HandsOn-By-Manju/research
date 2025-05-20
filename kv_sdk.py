import pandas as pd
import time
import json

from azure.identity import AzureCliCredential
from azure.mgmt.keyvault import KeyVaultManagementClient
from azure.core.exceptions import ResourceNotFoundError
from pandas.io.excel import ExcelWriter

# ===== CONFIGURATION =====
input_file = "keyvault_input.xlsx"
output_file = "keyvault_filtered_network_access_report_sdk.xlsx"
sheet_name = "Sheet1"
policy_id_filter = ["KV-PublicAccess", "KV-OpenToAll"]  # Edit as needed
# ==========================

start_time = time.time()

# Step 1: Load Excel input
df = pd.read_excel(input_file, sheet_name=sheet_name)

# Normalize column names
df.columns = df.columns.str.strip().str.title()

# Normalize Policy ID values
df["Policy Id"] = df["Policy Id"].astype(str).str.strip().str.upper()
normalized_filter = [pid.strip().upper() for pid in policy_id_filter]
filtered_df = df[df["Policy Id"].isin(normalized_filter)]

print(f"📄 Loaded {len(filtered_df)} filtered Key Vault rows matching Policy IDs: {policy_id_filter}")

# Step 2: Azure SDK client using Azure CLI auth
credential = AzureCliCredential()
results = []
success_count = 0
error_count = 0

# Step 3: Loop through each filtered row
for idx, row in enumerate(filtered_df.itertuples(index=False), start=1):
    row_dict = row._asdict()
    sub_id = row_dict["Subscription Id"]
    vault_name = row_dict["Key Vault Name"]
    policy_id = row_dict["Policy Id"]

    print(f"\n🔄 [{idx}/{len(filtered_df)}] Checking Key Vault: {vault_name} in Subscription: {sub_id}")

    try:
        # Get KeyVault client for this subscription
        kv_client = KeyVaultManagementClient(credential, sub_id)

        # List all vaults and find match by name
        vaults = kv_client.vaults.list()
        match = next((v for v in vaults if v.name.lower() == vault_name.lower()), None)

        if not match:
            raise ResourceNotFoundError(f"Key Vault '{vault_name}' not found in subscription {sub_id}")

        # Extract fields
        props = match.properties
        network_acls = props.network_acls.as_dict() if props.network_acls else {}
        public_access = props.public_network_access or "Unknown"
        pe_connections = props.private_endpoint_connections or []
        pe_summary = f"{len(pe_connections)} endpoint(s)" if pe_connections else "None"

        print(f"✅ Public: {public_access}, Private: {pe_summary}")
        success_count += 1

        results.append({
            "Policy ID": policy_id,
            "Subscription ID": sub_id,
            "Key Vault Name": vault_name,
            "Public Network Access": public_access,
            "Private Endpoints": pe_summary,
            "Network ACLs (Raw JSON)": json.dumps(network_acls, indent=2)
        })

    except Exception as ex:
        print(f"❌ Error: {ex}")
        error_count += 1
        results.append({
            "Policy ID": policy_id,
            "Subscription ID": sub_id,
            "Key Vault Name": vault_name,
            "Public Network Access": "Error",
            "Private Endpoints": "Error",
            "Network ACLs (Raw JSON)": f"Error: {str(ex)}"
        })

# Step 4: Export to Excel
result_df = pd.DataFrame(results)

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    result_df.to_excel(writer, index=False, sheet_name="Access Report")
    workbook = writer.book
    worksheet = writer.sheets["Access Report"]

    # Header style
    header_format = workbook.add_format({'bold': True, 'bg_color': '#87CEEB', 'border': 1})
    for col_num, col_name in enumerate(result_df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        worksheet.set_column(col_num, col_num, 50)
    worksheet.freeze_panes(1, 0)

# Final summary
elapsed = round(time.time() - start_time, 2)
print("\n📊 Summary:")
print(f"✅ Successful checks: {success_count}")
print(f"❌ Errors: {error_count}")
print(f"📁 Excel saved to: {output_file}")
print(f"⏱️ Time taken: {elapsed} seconds ({elapsed / 60:.2f} mins)")
