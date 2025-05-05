import pandas as pd
import time

# === File Inputs ===
input_csv = "input_file.csv"
output_excel = "output_step5.xlsx"
anex_file = "Report_Anex.xlsx"

# === Preprocessing Settings ===
columns_to_remove = [
    "DummyColumn1", "DummyColumn2", "DummyColumn3",
    "DummyColumn4", "DummyColumn5", "DummyColumn6",
    "DummyColumn7", "DummyColumn8", "DummyColumn9"
]

columns_to_add = [
    "Description",
    "Remediation Steps",
    "Environment",
    "Primary Contact",
    "Manager / Sr Manager / Director / Sr Director",
    "Sr Director / VP",
    "VP / SVP / CVP",
    "BU"
]

parse_account_column = True
account_column_name = "Account"
resource_column_name = "Resource ID"

# === Validation Details ===
validation_checks = [
    {
        "name": "Subscription ID",
        "sheet": "Anex2_Sub_Sheet",
        "join_column": "Subscription ID"
    },
    {
        "name": "Policy ID",
        "sheet": "Anex1_Remediation_Sheet",
        "join_column": "Policy ID"
    }
]

# === Start Execution ===
start_time = time.time()
print("\n🚀 Starting preprocessing and validation...")

# Step 1: Load Input CSV
df = pd.read_csv(input_csv)
print(f"✅ Loaded input file: {input_csv}")

# Step 2: Parse Subscription ID and Name from Account
if parse_account_column and account_column_name in df.columns:
    print(f"🔧 Parsing 'Subscription ID' and 'Subscription Name' from '{account_column_name}'")
    df["Subscription ID"] = df[account_column_name].str.extract(r"^(\S+)\s*\(")[0].str.replace(r"\s+", "", regex=True)
    df["Subscription Name"] = df[account_column_name].str.extract(r"\((.*?)\)")[0].str.replace(r"\s+", "", regex=True)

# Step 3: Clean Resource ID to extract filename
if resource_column_name in df.columns:
    print(f"🔧 Cleaning '{resource_column_name}' to extract filename")
    df[resource_column_name] = df[resource_column_name].apply(lambda x: str(x).split("/")[-1])

# Step 4: Drop unwanted columns
existing_to_drop = [col for col in columns_to_remove if col in df.columns]
df.drop(columns=existing_to_drop, inplace=True)
print(f"🧹 Dropped columns: {existing_to_drop if existing_to_drop else 'None'}")

# Step 5: Add empty columns
print(f"➕ Adding empty columns: {columns_to_add}")
for col in columns_to_add:
    df[col] = ""

# Step 6: Validate Subscription ID and Policy ID
for check in validation_checks:
    name = check["name"]
    sheet = check["sheet"]
    join_column = check["join_column"]

    print(f"\n🔍 Validating {name} using sheet '{sheet}'")
    try:
        df[join_column] = df[join_column].astype(str).str.strip()
        df_anex = pd.read_excel(anex_file, sheet_name=sheet)
        df_anex[join_column] = df_anex[join_column].astype(str).str.strip()

        input_ids = set(df[join_column].dropna())
        anex_ids = set(df_anex[join_column].dropna())

        unmatched_ids = sorted(input_ids - anex_ids)
        matched_ids = sorted(input_ids & anex_ids)

        total = len(input_ids)
        matched = len(matched_ids)
        unmatched = len(unmatched_ids)
        match_percent = round((matched / total) * 100, 2) if total else 0

        # Save unmatched
        if unmatched:
            file_unmatched = f"unmatched_{join_column.replace(' ', '_').lower()}.txt"
            with open(file_unmatched, "w") as f:
                for val in unmatched_ids:
                    f.write(f"{val}\n")
            print(f"❌ {unmatched} unmatched {name}(s) saved to {file_unmatched}")
        else:
            print(f"✅ All {name}s matched!")

        # Save matched
        file_matched = f"matched_{join_column.replace(' ', '_').lower()}.txt"
        with open(file_matched, "w") as f:
            for val in matched_ids:
                f.write(f"{val}\n")
        print(f"✅ {matched} matched {name}(s) saved to {file_matched}")

        # Print summary
        print(f"📊 {name} Summary:")
        print(f"   Total     : {total}")
        print(f"   Matched   : {matched}")
        print(f"   Unmatched : {unmatched}")
        print(f"   Match %   : {match_percent}%")

    except Exception as e:
        print(f"❌ Error validating {name}: {e}")

# Step 7: Map Description and Remediation Steps using Policy ID
print("\n🧩 Mapping 'Description' and 'Remediation Steps' using Policy ID...")
try:
    df_remed = pd.read_excel(anex_file, sheet_name="Anex1_Remediation_Sheet")
    df_remed["Policy ID"] = df_remed["Policy ID"].astype(str).str.strip()

    df = df.merge(df_remed[["Policy ID", "Policy Statement", "Policy Remediation"]],
                  on="Policy ID", how="left")

    df["Description"] = df["Policy Statement"].fillna("Policy details not available")
    df["Remediation Steps"] = df["Policy Remediation"].fillna("Remediation steps not available")

    df.drop(columns=["Policy Statement", "Policy Remediation"], inplace=True)
    print("✅ Remediation fields filled successfully.")

except Exception as e:
    print(f"❌ Error during remediation mapping: {e}")

# Step 8: Map Environment and Primary Contact using Subscription ID
print("\n📌 Mapping 'Environment' and 'Primary Contact' using Subscription ID...")
try:
    df_env = pd.read_excel(anex_file, sheet_name="Anex2_Sub_Sheet")
    df_env["Subscription ID"] = df_env["Subscription ID"].astype(str).str.strip()

    # Drop placeholders before merge to avoid _env suffix
    df.drop(columns=["Environment", "Primary Contact"], inplace=True, errors="ignore")

    df = df.merge(df_env[["Subscription ID", "Environment", "Primary Contact"]],
                  on="Subscription ID", how="left")

    df["Environment"] = df["Environment"].fillna("Environment not available")
    df["Primary Contact"] = df["Primary Contact"].fillna("Primary contact not available")

    print("✅ Environment and Primary Contact fields filled successfully.")

except Exception as e:
    print(f"❌ Error during environment/contact mapping: {e}")

# Step 9: Save Output
df.to_excel(output_excel, index=False)
print(f"\n✅ Final file saved to: {output_excel}")
print(f"⏱️ Time taken: {time.time() - start_time:.2f} seconds")
