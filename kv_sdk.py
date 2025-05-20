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
policy_id_filter = ["KV-PublicAccess", "KV-OpenToAll"]
# =========================

start_time = time.time()

# Step 1: Load Excel
df = pd.read_excel(input_file, sheet_name=sheet_name)
df["Policy ID"] = df["Policy ID"].astype(str).str.strip().str.upper()
normalized_filter = [pid.strip().upper() for pid in policy_id_filter]
filtered_df = df[df["Policy ID"].isin(normalized_filter)]

print(f"📄 Loaded {len(filtered_df)} filtered Key Vaults matching policies: {policy_id_filter}")

# Step 2: Prepare SDK
credential = AzureCliCredential()
results = []
success_count = 0
error_count = 0

# Step 3: Iterate through filtered rows
for idx, row in enumerate(filtered_df.itertuples(), start=1):
    sub_id = getattr(row, "Subscription_ID")
    vault_name = getattr(row, "Key_Vault_Name")
    policy_id = getattr(row, "Policy_ID")

    print(f"\n🔄 [{idx}/{len(filtered_df)}] Checking vault: {vault_name} in subscription: {sub_id}")

    try:
        kv_client = KeyVaultManagementClient(credential, sub_id)
        vaults = kv_client.vaults.list()
        match = next((v for v in vaults if v.name.lower() == vault_name.lower()), None)

        if not match:
            raise ResourceNotFoundError(f"Key Vault '{vault_name}' not found.")

        props = match.properties
        network_acls = props.network_acls.as_dict() if props.network_acls else {}
        public_access = props.public_network_access or "Unknown"
        pe_connections = props.private_endpoint_connections or []
        pe_summary = f"{len(pe_connections)} endpoint(s)" if pe_connections else "None"

        results.append({
            "Policy ID": policy_id,
            "Subscription ID": sub_id,
            "Key Vault Name": vault_name,
            "Public Network Access": public_access,
            "Private Endpoints": pe_summary,
            "Network ACLs (Raw JSON)": json.dumps(network_acls, indent=2)
        })
        print(f"✅ Public={public_access}, Private={pe_summary}")
        success_count += 1

    except Exception as ex:
        print(f"❌ Error: {ex}")
        results.append({
            "Policy ID": policy_id,
            "Subscription ID": sub_id,
            "Key Vault Name": vault_name,
            "Public Network Access": "Error",
            "Private Endpoints": "Error",
            "Network ACLs (Raw JSON)": f"Error: {str(ex)}"
        })
        error_count += 1

# Step 4: Write to Excel
result_df = pd.DataFrame(results)

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    result_df.to_excel(writer, index=False, sheet_name="Access Report")
    workbook = writer.book
    worksheet = writer.sheets["Access Report"]

    # Style
    header_format = workbook.add_format({'bold': True, 'bg_color': '#87CEEB', 'border': 1})
    for col, col_name in enumerate(result_df.columns):
        worksheet.write(0, col, col_name, header_format)
        worksheet.set_column(col, col, 50)
    worksheet.freeze_panes(1, 0)

elapsed = round(time.time() - start_time, 2)
print("\n📊 Summary:")
print(f"✅ Success: {success_count}")
print(f"❌ Errors: {error_count}")
print(f"📁 Excel saved to: {output_file}")
print(f"⏱️ Time taken: {elapsed} seconds ({elapsed/60:.2f} mins)")
