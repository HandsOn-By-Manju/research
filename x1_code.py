import pandas as pd
import time

# === Input and reference files ===
input_csv = "input_file.csv"
output_excel = "output_step5.xlsx"
anex_file = "Report_Anex.xlsx"

# === Preprocessing settings ===
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

# === Validation checks ===
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

# === Start execution ===
start_time = time.time()
print("\n🚀 Starting preprocessing and validation...")

# Step 1: Load CSV
df = pd.read_csv(input_csv)
print(f"✅ Loaded input file: {input_csv}")

# Step 2: Extract Subscription ID and Name
if parse_account_column and account_column_name in df.columns:
    print(f"🔧 Extracting Subscription ID and Name from '{account_column_name}'")
    df["Subscription ID"] = df[account_column_name].str.extract(r"^(\S+)\s*\(")[0].str.replace(r"\s+", "", regex=True)
    df["Subscription Name"] = df[account_column_name].str.extract(r"\((.*?)\)")[0].str.replace(r"\s+", "", regex=True)

# Step 3: Clean Resource ID
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

    print(f"\n🔍 Validating {name} against '{sheet}' in {anex_file}")
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
            unmatched_file = f"unmatched_{join_column.replace(' ', '_').lower()}.txt"
            with open(unmatched_file, "w") as f:
                for uid in unmatched_ids:
                    f.write(f"{uid}\n")
            print(f"❌ {unmatched} unmatched {name}(s) saved to {unmatched_file}")
        else:
            print(f"✅ All {name}s matched!")

        # Save matched
        matched_file = f"matched_{join_column.replace(' ', '_').lower()}.txt"
        with open(matched_file, "w") as f:
            for mid in matched_ids:
                f.write(f"{mid}\n")
        print(f"✅ {matched} matched {name}(s) saved to {matched_file}")

        # Print summary
        print(f"📊 Summary for {name}:")
        print(f"   Total     : {total}")
        print(f"   Matched   : {matched}")
        print(f"   Unmatched : {unmatched}")
        print(f"   Match %   : {match_percent}%")

    except Exception as e:
        print(f"❌ Error during validation of {name}: {e}")

# Step 7: Save preprocessed output
df.to_excel(output_excel, index=False)
print(f"\n✅ Preprocessed file saved to '{output_excel}'")
print(f"⏱️ Total time taken: {time.time() - start_time:.2f} seconds")
