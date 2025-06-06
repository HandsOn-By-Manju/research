import subprocess
import json
import pandas as pd

# ------------------ CONFIGURATION ------------------
CONFIG = {
    "subscription_id": "your-subscription-id",  # 🔁 Replace with your Azure Subscription ID
    "output_excel": "Azure_VM_Details.xlsx"
}
# ----------------------------------------------------

def fetch_vm_data(subscription_id):
    print(f"\n🔍 Fetching VM data from subscription: {subscription_id}")
    try:
        cmd = [
            "az", "vm", "list",
            "--subscription", subscription_id,
            "--show-details", "--output", "json"
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, check=True)
        return json.loads(result.stdout)
    except subprocess.CalledProcessError as e:
        print(f"❌ Azure CLI Error:\n{e.stderr}")
        return []
    except Exception as ex:
        print(f"❌ Unexpected error: {str(ex)}")
        return []

def extract_vm_info(vm_list, subscription_id):
    print(f"📦 Extracting details for {len(vm_list)} VM(s)...")
    extracted = []
    for vm in vm_list:
        os_disk = vm.get("storageProfile", {}).get("osDisk", {})
        image = vm.get("storageProfile", {}).get("imageReference", {})
        identity = vm.get("identity", {})
        os_profile = vm.get("osProfile", {})
        linux_config = os_profile.get("linuxConfiguration", {})
        windows_config = os_profile.get("windowsConfiguration", {})

        extracted.append({
            "Subscription ID": subscription_id,
            "VM Name": vm.get("name"),
            "Resource Group": vm.get("resourceGroup"),
            "Region": vm.get("location"),
            "Availability Zone": ",".join(vm.get("zones", [])) if vm.get("zones") else "",
            "VM Size": vm.get("hardwareProfile", {}).get("vmSize"),
            "OS Type": os_disk.get("osType"),
            "Computer Name": os_profile.get("computerName"),
            "Admin Username": os_profile.get("adminUsername"),
            "OS Disk Name": os_disk.get("name"),
            "OS Disk Size (GB)": os_disk.get("diskSizeGb"),
            "Image Publisher": image.get("publisher"),
            "Image Offer": image.get("offer"),
            "Image SKU": image.get("sku"),
            "Image Version": image.get("version"),
            "Data Disks Count": len(vm.get("storageProfile", {}).get("dataDisks", [])),
            "Private IP": vm.get("privateIps"),
            "Public IP": vm.get("publicIps"),
            "NICs": ", ".join([nic.get("id", "").split("/")[-1] for nic in vm.get("networkProfile", {}).get("networkInterfaces", [])]),
            "Power State": vm.get("powerState"),
            "Provisioning State": vm.get("provisioningState"),
            "Identity Type": identity.get("type"),
            "User Assigned Identities": ", ".join(identity.get("userAssignedIdentities", {}).keys()) if identity.get("userAssignedIdentities") else "",
            "Linux Configuration": json.dumps(linux_config) if linux_config else "",
            "Windows Configuration": json.dumps(windows_config) if windows_config else "",
            "Tags": json.dumps(vm.get("tags", {}))
        })
    return extracted

def save_to_excel(data, output_file):
    print(f"💾 Saving results to Excel: {output_file}")
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)
    print(f"✅ Export complete! {len(df)} VM entries written.\n")

def main():
    subscription_id = CONFIG["subscription_id"]
    output_file = CONFIG["output_excel"]

    print("🚀 Azure VM Full Inventory Script Started")
    vm_data = fetch_vm_data(subscription_id)

    if not vm_data:
        print("⚠️ No VM data found or failed to retrieve data.")
        return

    vm_info = extract_vm_info(vm_data, subscription_id)
    save_to_excel(vm_info, output_file)
    print("🏁 Script finished.\n")

if __name__ == "__main__":
    main()
