import os
import pandas as pd
import time
from azure.identity import AzureCliCredential
from azure.core.exceptions import ClientAuthenticationError
from azure.mgmt.cosmosdb import CosmosDBManagementClient

# ---------------------
# 📥 Config
# ---------------------
INPUT_FILE = "cosmosdb_input.xlsx"
SHEET_NAME = "Sheet1"
POLICY_FILTER_VALUE = "123456"
PARTIAL_OUTPUT_FILE = "cosmosdb_public_access_partial.xlsx"
FINAL_OUTPUT_FILE = "cosmosdb_public_access_output.xlsx"
SAVE_EVERY = 100

# ---------------------
# 🔐 Azure CLI Login Check
# ---------------------
try:
    credential = AzureCliCredential()
    credential.get_token("https://management.azure.com/.default")
except ClientAuthenticationError:
    print("⚠️ Azure CLI session expired or not logged in.")
    print("👉 Please run 'az login' and rerun this script.")
    exit(1)

# ---------------------
# ⏱️ Timer Start
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
print(f"🔎 Matching Policy ID '{POLICY_FILTER_VALUE}': {len(filtered_df)}")

if filtered_df.empty:
    print("⚠️ No matching rows. Exiting.")
    exit(1)

# ---------------------
# 🔁 Resume Support
# ---------------------
processed_pairs = set()
results = []

if os.path.exists(PARTIAL_OUTPUT_FILE):
    partial_df = pd.read_excel(PARTIAL_OUTPUT_FILE)
    results = partial_df.to_dict(orient='records')
    for _, row in partial_df.iterrows():
        key = (row['Cosmos DB Name'].strip().lower(), row['Subscription ID'].strip().lower())
        processed_pairs.add(key)
    print(f"🔁 Resuming from {len(processed_pairs)} processed Cosmos DB accounts.")

# ---------------------
# 🚀 Process Each Cosmos DB Account
# ---------------------
filtered_df = filtered_df.reset_index(drop=True)
total = len(filtered_df)
processed_count = len(processed_pairs)

def extract_vnet_subnet_name(vnet_rule_id):
    try:
        parts = vnet_rule_id.split('/')
        vnet = parts[parts.index('virtualNetworks') + 1]
        subnet = parts[parts.index('subnets') + 1]
        return f"{vnet}/{subnet}"
    except Exception:
        return vnet_rule_id  # fallback to full ID if parsing fails

for idx, (_, row) in enumerate(filtered_df.iterrows(), start=1):
    sub_id = str(row['Subscription ID']).strip()
    acc_name = str(row['Cosmos DB Name']).strip()
    pair_key = (acc_name.lower(), sub_id.lower())

    if pair_key in processed_pairs:
        continue

    print(f"\n🔍 [Processing {processed_count + 1} of {total}] '{acc_name}' in subscription '{sub_id}'")

    entry = {
        "Index": processed_count + 1,
        "Subscription ID": sub_id,
        "Cosmos DB Name": acc_name,
        "Resource Group": "",
        "Public Network Access": "",
        "IP Rules Count": "",
        "IP Rule Details": "",
        "VNet Rules Count": "",
        "VNet Rule Details": "",
        "Exposed to All Networks?": "",
        "Status": "",
        "Message": ""
    }

    try:
        client = CosmosDBManagementClient(credential, sub_id)
        accounts = list(client.database_accounts.list())

        acc = next((a for a in accounts if a.name.lower() == acc_name.lower()), None)
        if not acc:
            entry["Status"] = "Failed"
            entry["Message"] = "Cosmos DB account not found"
            print(f"❌ {entry['Message']}")
        else:
            rg_name = acc.id.split("/")[acc.id.split("/").index("resourceGroups") + 1]
            entry["Resource Group"] = rg_name

            props = client.database_accounts.get(rg_name, acc_name)

            public_access = props.public_network_access or "Enabled"
            ip_rules = props.ip_rules or []
            vnet_rules = props.virtual_network_rules or []

            entry["Public Network Access"] = public_access
            entry["IP Rules Count"] = len(ip_rules)
            entry["IP Rule Details"] = ", ".join([r.ip_address_or_range for r in ip_rules])
            entry["VNet Rules Count"] = len(vnet_rules)
            entry["VNet Rule Details"] = ", ".join([extract_vnet_subnet_name(r.id) for r in vnet_rules if r.id])

            if public_access == "Enabled" and len(ip_rules) == 0 and len(vnet_rules) == 0:
                entry["Exposed to All Networks?"] = "Yes 🌐"
            else:
                entry["Exposed to All Networks?"] = "No 🔒"

            entry["Status"] = "Success"
            entry["Message"] = "Processed successfully"
            print(f"✅ {entry['Message']}")

    except ClientAuthenticationError:
        print(f"\n⛔ CLI session expired at row {processed_count + 1} → '{acc_name}'")
        pd.DataFrame(results).to_excel(PARTIAL_OUTPUT_FILE, index=False)
        print(f"💾 Partial saved → {PARTIAL_OUTPUT_FILE}")
        print("👉 Run 'az login' and rerun this script.")
        exit(1)
    except Exception as e:
        entry["Status"] = "Failed"
        entry["Message"] = str(e)
        print(f"❗ Error: {e}")

    results.append(entry)
    processed_pairs.add(pair_key)
    processed_count += 1

    if processed_count % SAVE_EVERY == 0:
        pd.DataFrame(results).to_excel(PARTIAL_OUTPUT_FILE, index=False)
        print(f"💾 Saved {processed_count} rows → {PARTIAL_OUTPUT_FILE}")

# ---------------------
# ✅ Final Save
# ---------------------
pd.DataFrame(results).to_excel(FINAL_OUTPUT_FILE, index=False)
print(f"\n📁 Final output saved to: {FINAL_OUTPUT_FILE}")

# ---------------------
# ⏱️ Execution Time
# ---------------------
elapsed = time.time() - start_time
h, m, s = int(elapsed // 3600), int((elapsed % 3600) // 60), round(elapsed % 60, 2)
print(f"\n⏱️ Execution time: {h} hours, {m} minutes, {s} seconds")
print("✅ All Cosmos DB accounts processed.")
