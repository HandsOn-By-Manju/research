import pandas as pd
import os
import time
from azure.identity import AzureCliCredential
from azure.core.exceptions import ClientAuthenticationError, HttpResponseError
from azure.mgmt.resource import SubscriptionClient, ResourceManagementClient

# ---------------------
# 📥 Config
# ---------------------
PARTIAL_FILE = "azure_tags_partial.xlsx"
FINAL_FILE = "azure_subscription_and_rg_tags.xlsx"

# ---------------------
# ⏱️ Timer Start
# ---------------------
start_time = time.time()

# ---------------------
# 🔐 Azure CLI Login Check
# ---------------------
try:
    credential = AzureCliCredential()
    credential.get_token("https://management.azure.com/.default")
except ClientAuthenticationError:
    raise SystemExit("⚠️ Azure CLI session expired. Please run 'az login' and rerun this script.")

# ---------------------
# 🔁 Load Partial Output if Exists
# ---------------------
if os.path.exists(PARTIAL_FILE):
    partial_df = pd.read_excel(PARTIAL_FILE, sheet_name=None)
    subscription_data = partial_df.get("Subscription Tags", pd.DataFrame()).to_dict("records")
    rg_tag_data = partial_df.get("Resource Group Tags", pd.DataFrame()).to_dict("records")
    processed_subs = {row["Subscription ID"] for row in subscription_data}
    print(f"🔁 Resuming from {len(processed_subs)} subscriptions")
else:
    subscription_data = []
    rg_tag_data = []
    processed_subs = set()

# ---------------------
# 🔁 Retry Helper
# ---------------------
def retry_call(func, retries=3, delay=5):
    for attempt in range(retries):
        try:
            return func()
        except HttpResponseError as e:
            print(f"⏳ Retry {attempt+1}/{retries} - {e}")
            time.sleep(delay)
    raise

# ---------------------
# 📦 Process Subscriptions and RG Tags
# ---------------------
sub_client = SubscriptionClient(credential)
all_subs = list(sub_client.subscriptions.list())
total_subs = len(all_subs)
total_rgs = 0

for idx, sub in enumerate(all_subs, start=1):
    sub_id = sub.subscription_id
    sub_name = sub.display_name

    if sub_id in processed_subs:
        continue

    print(f"\n🔍 [{idx} of {total_subs}] Processing subscription: {sub_name} ({sub_id})")

    sub_entry = {
        "Subscription ID": sub_id,
        "Subscription Name": sub_name,
        "Tags": "",
        "Message": ""
    }

    try:
        sub_details = retry_call(lambda: sub_client.subscriptions.get(sub_id))
        if sub_details.tags:
            sub_entry["Tags"] = ", ".join(f"{k}={v}" for k, v in sub_details.tags.items())
        sub_entry["Message"] = "Success"
    except Exception as e:
        sub_entry["Tags"] = "ERROR"
        sub_entry["Message"] = str(e)

    subscription_data.append(sub_entry)

    try:
        rg_client = ResourceManagementClient(credential, sub_id)
        rgs = list(rg_client.resource_groups.list())
        print(f"📁   Found {len(rgs)} resource groups in {sub_name}")
        for rg_idx, rg in enumerate(rgs, start=1):
            rg_tags = rg.tags or {}
            print(f"📦     - [{rg_idx} of {len(rgs)}] {rg.name}")
            rg_tag_data.append({
                "Subscription ID": sub_id,
                "Subscription Name": sub_name,
                "Resource Group": rg.name,
                "Location": rg.location,
                "Tags": ", ".join(f"{k}={v}" for k, v in rg_tags.items()) if rg_tags else "",
                "Message": "Success"
            })
        total_rgs += len(rgs)
    except Exception as e:
        rg_tag_data.append({
            "Subscription ID": sub_id,
            "Subscription Name": sub_name,
            "Resource Group": "ERROR",
            "Location": "",
            "Tags": "",
            "Message": str(e)
        })

    with pd.ExcelWriter(PARTIAL_FILE, engine="xlsxwriter") as writer:
        pd.DataFrame(subscription_data).to_excel(writer, sheet_name="Subscription Tags", index=False)
        pd.DataFrame(rg_tag_data).to_excel(writer, sheet_name="Resource Group Tags", index=False)
    print(f"💾 Partial saved after: {sub_name}")

# ---------------------
# 📤 Final Save
# ---------------------
with pd.ExcelWriter(FINAL_FILE, engine="xlsxwriter") as writer:
    pd.DataFrame(subscription_data).to_excel(writer, sheet_name="Subscription Tags", index=False)
    pd.DataFrame(rg_tag_data).to_excel(writer, sheet_name="Resource Group Tags", index=False)

# ---------------------
# ⏱️ Execution Time
# ---------------------
elapsed = time.time() - start_time
h, m, s = int(elapsed // 3600), int((elapsed % 3600) // 60), round(elapsed % 60, 2)
print(f"\n✅ Completed {len(subscription_data)} subscriptions and {total_rgs} resource groups")
print(f"📁 Final output saved to: {FINAL_FILE}")
print(f"⏱️ Completed in {h}h {m}m {s}s")
