import pandas as pd
import subprocess
import time
import json

from pandas.io.excel import ExcelWriter

# ====== CONFIGURATION ======
input_file = "keyvault_input.xlsx"
output_file = "keyvault_filtered_network_access_report.xlsx"
sheet_name = "Sheet1"
policy_id_filter = ["KV-PublicAccess", "KV-OpenToAll"]  # <--- Modify this as needed
# ===========================

start = time.time()

# Load Excel
df = pd.read_excel(input_file, sheet_name=sheet_name)
print(f"📄 Loaded {len(df)} rows from {input_file}")

# Clean and normalize 'Policy ID' for filtering
df["Policy ID"] = df["Policy ID"].astype(str).str.strip().str.upper()
normalized_policy_filter = [pid.strip().upper() for pid in policy_id_filter]

# Filter rows by Policy ID
filtered_df = df[df["Policy ID"].isin(normalized_policy_filter)]
print(f"🔍 Found {len(filtered_df)} matching rows for Policy ID(s): {policy_id_filter}")

results = []

for idx, row in filtered_df.iterrows():
    subscription_id = str(row["Subscription ID"]).strip()
    kv_name = str(row["Key Vault Name"]).strip()
    policy_id = str(row["Policy ID"]).strip()

    print(f"\n🔄 Checking: KeyVault={kv_name}, Subscription={subscription_id}, Policy={policy_id}")

    try:
        subprocess.run(["az", "account", "set", "--subscription", subscription_id], check=True)

        cmd = [
            "az", "keyvault", "show",
            "--name", kv_name,
            "--subscription", subscription_id,
            "--output", "json"
        ]
        raw_output = subprocess.check_output(cmd, text=True)
        kv_data = json.loads(raw_output)

        network_acls = json.dumps(kv_data.get("properties", {}).get("networkAcls", {}), indent=2)
        public_network_access = kv_data.get("properties", {}).get("publicNetworkAccess", "Unknown")
        private_endpoint_connections = kv_data.get("properties", {}).get("privateEndpointConnections", [])
        private_endpoint_summary = f"{len(private_endpoint_connections)} endpoint(s)" if private_endpoint_connections else "None"

        print(f"✅ Access: Public={public_network_access}, Private={private_endpoint_summary}")

        results.append({
            "Policy ID": policy_id,
            "Subscription ID": subscription_id,
            "Key Vault Name": kv_name,
            "Public Network Access": public_network_access,
            "Private Endpoints": private_endpoint_summary,
            "Network ACLs (Raw JSON)": network_acls
        })

    except subprocess.CalledProcessError as e:
        print(f"⚠️ Error for {kv_name}: {e}")
        results.append({
            "Policy ID": policy_id,
            "Subscription ID": subscription_id,
            "Key Vault Name": kv_name,
            "Public Network Access": "Error",
            "Private Endpoints": "Error",
            "Network ACLs (Raw JSON)": f"Error: {str(e)}"
        })

# Save to Excel
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

print(f"\n✅ Filtered report saved to: {output_file}")
print(f"⏱️ Completed in {round(time.time() - start, 2)} seconds.")
