import os
import pandas as pd
import time
from azure.identity import AzureCliCredential
from azure.core.exceptions import ClientAuthenticationError
from azure.mgmt.keyvault import KeyVaultManagementClient

# ---------------------
# 📥 Config
# ---------------------
INPUT_FILE = "keyvault_input.xlsx"
SHEET_NAME = "Sheet1"
POLICY_FILTER_VALUE = "123456"
PARTIAL_OUTPUT_FILE = "keyvault_rbac_output_partial.xlsx"
FINAL_OUTPUT_FILE = "keyvault_rbac_output.xlsx"
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
        key = (row['Key Vault Name'].strip().lower(), row['Subscription ID'].strip().lower())
        processed_pairs.add(key)
    print(f"🔁 Resuming from {len(processed_pairs)} processed vaults.")

# ---------------------
# 📦 Vault Caching
# ---------------------
vault_cache = {}

def get_vault_from_cache(client, subscription_id, kv_name):
    if subscription_id not in vault_cache:
        print(f"📦 Caching Key Vaults in subscription: {subscription_id}")
        vault_cache[subscription_id] = list(client.vaults.list())
    return next((v for v in vault_cache[subscription_id] if v.name.lower() == kv_name.lower()), None)

# ---------------------
# 🚀 Process Each Vault
# ---------------------
filtered_df = filtered_df.reset_index(drop=True)
total = len(filtered_df)
processed_count = len(processed_pairs)

for idx, (_, row) in enumerate(filtered_df.iterrows(), start=1):
    sub_id = str(row['Subscription ID']).strip()
    kv_name = str(row['Key Vault Name']).strip()
    pair_key = (kv_name.lower(), sub_id.lower())

    if pair_key in processed_pairs:
        continue

    print(f"\n🔍 [Processing {processed_count + 1} of {total}] '{kv_name}' in subscription '{sub_id}'")

    entry = {
        "Index": processed_count + 1,
        "Subscription ID": sub_id,
        "Key Vault Name": kv_name,
        "Resource Group": "",
        "RBAC Enabled?": "",
        "Access Control Type": "",
        "Access Policy Count": "",
        "Access Policy Object IDs": "",
        "Access Policy Permissions": "",
        "Status": "",
        "Message": ""
    }

    try:
        client = KeyVaultManagementClient(credential, sub_id)
        vault_ref = get_vault_from_cache(client, sub_id, kv_name)

        if not vault_ref:
            entry["Status"] = "Failed"
            entry["Message"] = "Key Vault not found"
            print(f"❌ {entry['Message']}")
        else:
            rg_parts = vault_ref.id.split('/')
            rg_name = rg_parts[rg_parts.index('resourceGroups') + 1]
            vault = client.vaults.get(rg_name, kv_name)

            props = vault.properties
            rbac_enabled = props.enable_rbac_authorization if props.enable_rbac_authorization else False

            entry["Resource Group"] = rg_name
            entry["RBAC Enabled?"] = "Yes ✅" if rbac_enabled else "No ❌"
            entry["Access Control Type"] = "RBAC" if rbac_enabled else "Access Policy"

            if not rbac_enabled:
                policies = props.access_policies or []
                entry["Access Policy Count"] = len(policies)
                object_ids = [p.object_id for p in policies if p.object_id]
                perms = []

                for p in policies:
                    if hasattr(p.permissions, 'keys'):
                        perms += p.permissions.keys or []
                    if hasattr(p.permissions, 'secrets'):
                        perms += p.permissions.secrets or []
                    if hasattr(p.permissions, 'certificates'):
                        perms += p.permissions.certificates or []

                entry["Access Policy Object IDs"] = ", ".join(object_ids)
                entry["Access Policy Permissions"] = ", ".join(sorted(set(perms)))

            entry["Status"] = "Success"
            entry["Message"] = "Processed successfully"
            print(f"✅ {entry['Message']}")

    except ClientAuthenticationError:
        print(f"\n⛔ CLI session expired at row {processed_count + 1} → '{kv_name}'")
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
print("✅ All vaults processed.")
