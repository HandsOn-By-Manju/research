import os
import time
from datetime import datetime
from azure.identity import AzureCliCredential
from azure.mgmt.resource import SubscriptionClient, ResourceManagementClient
from openpyxl import Workbook

# Start execution timer
start_time = time.time()

# Authenticate using Azure CLI
print("[INFO] Authenticating using Azure CLI credentials...")
credential = AzureCliCredential()
sub_client = SubscriptionClient(credential)

# Data holders
subscriptions_data = []
resource_groups_data = []
subscription_tag_keys = set()
rg_tag_keys = set()

print("[INFO] Starting tag collection from all subscriptions...\n")

# Loop through subscriptions
for sub in sub_client.subscriptions.list():
    sub_id = sub.subscription_id
    sub_name = sub.display_name
    print(f"🔍 Processing Subscription: {sub_name} ({sub_id})")

    # Subscription tags
    sub_details = sub_client.subscriptions.get(sub_id)
    sub_tags = sub_details.tags or {}
    if sub_tags:
        print(f"    ✅ Subscription has {len(sub_tags)} tag(s)")
    else:
        print(f"    ⚠️  No tags found on this subscription")

    subscription_tag_keys.update(sub_tags.keys())
    subscriptions_data.append({
        "Subscription Name": sub_name,
        "Subscription ID": sub_id,
        **sub_tags
    })

    # Initialize Resource Client
    resource_client = ResourceManagementClient(credential, sub_id)

    # Loop through Resource Groups
    rg_list = list(resource_client.resource_groups.list())
    if not rg_list:
        print("    ℹ️  No resource groups found.")
        continue

    print(f"    ➕ Found {len(rg_list)} resource group(s)")
    for rg in rg_list:
        rg_name = rg.name
        rg_location = rg.location
        rg_details = resource_client.resource_groups.get(rg_name)
        rg_tags = rg_details.tags or {}
        if rg_tags:
            print(f"        ✅ RG '{rg_name}' has {len(rg_tags)} tag(s)")
        else:
            print(f"        ⚠️  RG '{rg_name}' has no tags")

        rg_tag_keys.update(rg_tags.keys())

        resource_groups_data.append({
            "Subscription Name": sub_name,
            "Subscription ID": sub_id,
            "Resource Group": rg_name,
            "Location": rg_location,
            **rg_tags
        })

# Function to write data to a sheet
def write_sheet(wb, sheet_name, data, extra_columns, tag_keys):
    ws = wb.create_sheet(sheet_name)
    tag_keys_sorted = sorted(tag_keys)
    headers = extra_columns + tag_keys_sorted
    ws.append(headers)
    for item in data:
        row = [item.get(col, "") for col in headers]
        ws.append(row)
    print(f"[EXCEL] ✅ Sheet '{sheet_name}' written with {len(data)} row(s) and {len(headers)} column(s)")

# Create Excel workbook
wb = Workbook()
wb.remove(wb.active)  # remove default sheet

print("\n[INFO] Writing data to Excel workbook...")
write_sheet(wb, "Subscription Tags", subscriptions_data, ["Subscription Name", "Subscription ID"], subscription_tag_keys)
write_sheet(wb, "ResourceGroup Tags", resource_groups_data, ["Subscription Name", "Subscription ID", "Resource Group", "Location"], rg_tag_keys)

# Save Excel file
filename = f"Azure_Subscription_RG_Tags_{datetime.now().strftime('%Y%m%d')}.xlsx"
wb.save(filename)
print(f"\n✅ [SUCCESS] Report saved to file: {filename}")

# End execution timer
end_time = time.time()
elapsed_time = end_time - start_time
print(f"⏱️ [INFO] Total execution time: {elapsed_time:.2f} seconds")
