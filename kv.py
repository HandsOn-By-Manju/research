import pandas as pd
import os
import json
import time

from pandas.io.excel import ExcelWriter

# ====== CONFIGURATION ======
input_file = "keyvault_input.xlsx"
output_file = "keyvault_filtered_network_access_report.xlsx"
sheet_name = "Sheet1"
policy_id_filter = ["KV-PublicAccess", "KV-OpenToAll"]  # Change as needed
# ===========================

start = time.time()

# Load Excel
df = pd.read_excel(input_file, sheet_name=sheet_name)
print(f"📄 Loaded {len(df)} rows from {input_file}")

# Normalize 'Policy ID' for filtering
df["Policy ID"] = df["Policy ID"].astype(str).str.strip().str.upper()
normalized_policy_filter = [pid.strip().upper() for pid in policy_id_filter]

# Filter rows by Policy ID
filtered_df = df[df["Policy ID"].isin(normalized_policy_filter)]
print(f"🔍 Found {len(filtered_df)} matching rows for Policy ID(s): {policy_id_filter}")

results = []
success_count = 0
error_count = 0
total = len(filtered_df)

for idx, row in filtered_df.iterrows():
    subscription_id = str(row["Subscription ID"]).strip()
    kv_name = str(row["Key Vault Name"]).strip()
    policy_id = str(row["Policy ID"]).strip()

    print(f"\n🔄 [{idx + 1}/{total}] Checking Key Vault: {kv_name} (Policy: {policy_id})")

    try:
        # Set subscription
        os.system(f'az account set --subscription "{subscription_id}"')

        # Run az CLI and get output
        az_cmd = f'az keyvault show --name "{kv_name}" --subscription "{subscription_id}" --output json'
        with os.popen(az_cmd) as stream:
            raw_output = stream.read()

        kv_data = json.loads(raw_output)

        # Extract desired fields
        network_acls = json.dumps(kv_data.get("properties", {}).get("networkAcls", {}), indent=2)
        public_network_access = kv_data.get("properties", {}).get("publicNetworkAccess", "Unknown")
        private_endpoint_connections = kv_data.get("properties", {}).get("privateEndpointConnections", [])
        private_endpoint_summary = f"{len(private_endpoint_connections)} endpoint(s)" if private_endpoint_connections else "None"

        print(f"✅ Success: Public={public_network_access}, Private={private_endpoint_summary}")
        success_count += 1

        results.append({
            "Policy ID": policy_id,
            "Subscription ID": subscription_id,
            "Key Vault Name": kv_name,
            "Public Network Access": public_network_access,
            "Private Endpoints": private_endpoint_summary,
            "Network ACLs (Raw JSON)": network_acls
        })

    except Exception as e:
        print(f"❌ Error: Could not retrieve Key Vault '{kv_name}' → {e}")
        error_count += 1

        results.append({
            "Policy ID": policy_id,
            "Subscription ID": subscription_id,
            "Key Vault Name": kv_name,
            "Public Network Access": "Error",
            "Private Endpoints": "Error",
            "Network ACLs (Raw JSON)": f"Error: {str(e)}"
        })

# Write to Excel
result_df = pd.DataFrame(results)

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    result_df.to_excel(writer, index=False, sheet_name="Filtered Access Report")
    workbook = writer.book
    worksheet = writer.sheets["Filtered Access Report"]

    # Format headers
    header_format = workbook.add_format({'bold': True, 'bg_color': '#87CEEB', 'border': 1})
    for col_num, col_name in enumerate(result_df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        worksheet.set_column(col_num, col_num, 50)

    worksheet.freeze_panes(1, 0)

end = time.time()
elapsed = round(end - start, 2)
minutes = round(elapsed / 60, 2)

print("\n📊 Summary:")
print(f"✅ Successful checks: {success_count}")
print(f"❌ Errors: {error_count}")
print(f"📁 Excel report saved to: {output_file}")
print(f"⏱️ Total time: {elapsed} seconds ({minutes} mins)")
