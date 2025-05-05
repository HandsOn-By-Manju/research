import pandas as pd
import time
import difflib
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

# === Configurable Inputs ===
input_csv = "input_file.csv"
output_excel = "output_step5.xlsx"
anex_file = "Report_Anex.xlsx"

account_column_name = "Account"
resource_column_name = "Affected Resource"
parse_account_column = True

anex1_sheet = "Anex1_Remediation_Sheet"
anex2_sheet = "Anex2_Sub_Sheet"
anex3_sheet = "Anex3_Contact_Sheet"

primary_contact_column = "Primary Contact"
manager_columns = [
    "Manager / Sr Manager / Director / Sr Director",
    "Sr Director / VP",
    "VP / SVP / CVP",
    "BU"
]

columns_to_remove = [
    "DummyColumn1", "DummyColumn2", "DummyColumn3",
    "DummyColumn4", "DummyColumn5", "DummyColumn6",
    "DummyColumn7", "DummyColumn8", "DummyColumn9"
]

columns_to_add = [
    "Description",
    "Remediation Steps",
    "Environment",
    primary_contact_column,
] + manager_columns

final_columns = [
    "Cloud Provider", "Subscription ID", "Subscription Name",
    "Policy ID", "Policy Statement", "Affected Resource",
    "Severity", "Description", "Remediation Steps",
    "Region", "Service", "Environment", "Primary Contact",
    "Manager / Sr Manager / Director / Sr Director", "Sr Director / VP",
    "VP / SVP / CVP", "BU", "Account", "Finding"
]

validation_checks = [
    {"name": "Subscription ID", "sheet": anex2_sheet, "join_column": "Subscription ID"},
    {"name": "Policy ID", "sheet": anex1_sheet, "join_column": "Policy ID"}
]

# === Start Execution ===
start_time = time.time()
print("\n🚀 Starting preprocessing and validation...")

# Step 1: Load CSV
df = pd.read_csv(input_csv)
print(f"✅ Loaded input file: {input_csv}")

# Step 1.5: Rename columns
df.rename(columns={
    "Cloud provider": "Cloud Provider",
    "Policy statement": "Policy Statement",
    "Resource ID": "Affected Resource"
}, inplace=True)

# Step 2: Extract Subscription ID & Name
if parse_account_column and account_column_name in df.columns:
    print(f"🔧 Extracting Subscription ID and Name from '{account_column_name}'")
    df["Subscription ID"] = df[account_column_name].str.extract(r"^(\S+)\s*\(")[0].str.replace(r"\s+", "", regex=True)
    df["Subscription Name"] = df[account_column_name].str.extract(r"\((.*?)\)")[0].str.replace(r"\s+", "", regex=True)

# Step 3: Clean Affected Resource
if resource_column_name in df.columns:
    print(f"🔧 Cleaning '{resource_column_name}' to extract filename")
    df[resource_column_name] = df[resource_column_name].apply(lambda x: str(x).split("/")[-1])

# Step 4: Drop unwanted columns
existing_to_drop = [col for col in columns_to_remove if col in df.columns]
df.drop(columns=existing_to_drop, inplace=True)
print(f"🧹 Dropped columns: {existing_to_drop if existing_to_drop else 'None'}")

# Step 5: Add empty columns
print(f"➕ Adding columns: {columns_to_add}")
for col in columns_to_add:
    if col not in df.columns:
        df[col] = ""

# Step 6: Validation
for check in validation_checks:
    try:
        print(f"\n🔍 Validating {check['name']} using sheet '{check['sheet']}'")
        df_anex = pd.read_excel(anex_file, sheet_name=check["sheet"])
        df_anex[check["join_column"]] = df_anex[check["join_column"]].astype(str).str.strip()
        df[check["join_column"]] = df[check["join_column"]].astype(str).str.strip()
        unmatched = sorted(set(df[check["join_column"]]) - set(df_anex[check["join_column"]]))
        if unmatched:
            with open(f"unmatched_{check['join_column'].replace(' ', '_').lower()}.txt", "w") as f:
                for val in unmatched:
                    f.write(f"{val}\n")
            print(f"❌ {len(unmatched)} unmatched {check['name']} entries saved.")
        else:
            print(f"✅ All {check['name']} entries matched.")
    except Exception as e:
        print(f"❌ Error validating {check['name']}: {e}")

# Step 7: Map Description & Remediation
try:
    df_remed = pd.read_excel(anex_file, sheet_name=anex1_sheet)
    df_remed["Policy ID"] = df_remed["Policy ID"].astype(str).str.strip()
    df = df.merge(df_remed[["Policy ID", "Policy Statement", "Policy Remediation"]],
                  on="Policy ID", how="left")
    df["Description"] = df["Policy Statement"].fillna("Policy details not available")
    df["Remediation Steps"] = df["Policy Remediation"].fillna("Remediation steps not available")
    df.drop(columns=["Policy Statement", "Policy Remediation"], inplace=True, errors="ignore")
    print("✅ Mapped Description and Remediation Steps.")
except Exception as e:
    print(f"❌ Error mapping remediation: {e}")

# Step 8: Map Environment & Primary Contact
try:
    df_env = pd.read_excel(anex_file, sheet_name=anex2_sheet)
    df_env["Subscription ID"] = df_env["Subscription ID"].astype(str).str.strip()
    df = df.merge(df_env[["Subscription ID", "Environment", primary_contact_column]],
                  on="Subscription ID", how="left")
    df["Environment"] = df["Environment"].fillna("Environment not available")
    df[primary_contact_column] = df[primary_contact_column].fillna("Primary contact not available")
    print("✅ Mapped Environment and Primary Contact.")
except Exception as e:
    print(f"❌ Error mapping environment/contact: {e}")

# Step 9: Map Manager Hierarchy
try:
    df_contact = pd.read_excel(anex_file, sheet_name=anex3_sheet)
    df_contact.columns = df_contact.columns.str.strip()
    missing_columns = [col for col in manager_columns if col not in df_contact.columns]
    if missing_columns:
        print(f"❌ Missing columns in contact sheet: {missing_columns}")
        raise Exception("Missing manager columns.")
    df = df.merge(df_contact[[primary_contact_column] + manager_columns],
                  on=primary_contact_column, how="left")
    print("✅ Mapped Manager Hierarchy and BU.")
except Exception as e:
    print(f"❌ Error mapping contact hierarchy: {e}")

# Step 10: Reorder columns
existing_final_cols = [col for col in final_columns if col in df.columns]
df = df[existing_final_cols + [col for col in df.columns if col not in existing_final_cols]]

# Step 11: Save to Excel with formatting
print("\n💾 Saving Excel file with formatting...")
try:
    df.to_excel(output_excel, index=False)
    wb = load_workbook(output_excel)
    ws = wb.active

    # Formatting
    align = Alignment(horizontal="left", vertical="top", wrap_text=True)
    header_fill = PatternFill(start_color="B7DEE8", end_color="B7DEE8", fill_type="solid")
    header_font = Font(bold=True)

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = align

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font

    ws.freeze_panes = "A2"
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

    wb.save(output_excel)
    print(f"✅ Final file saved with formatting: {output_excel}")
except Exception as e:
    print(f"❌ Error formatting/saving Excel: {e}")

print(f"⏱️ Total time: {time.time() - start_time:.2f} seconds")
