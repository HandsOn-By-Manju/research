import pandas as pd
import time

# === User-defined variables ===
input_csv = "input_file.csv"
output_excel = "output_step5.xlsx"

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

# === Mapping definitions ===
mappings = [
    {
        "name": "Remediation Mapping",
        "file": "Report_Anex.xlsx",
        "sheet": "Anex1_Remediation_Sheet",
        "join_column": "Policy ID",
        "source_to_target": {
            "Policy Statement": "Description",
            "Policy Remediation": "Remediation Steps"
        }
    },
    {
        "name": "Subscription Mapping",
        "file": "Report_Anex.xlsx",
        "sheet": "Anex2_Sub_Sheet",
        "join_column": "Subscription ID",
        "source_to_target": {
            "Environment": "Environment",
            "Primary Contact": "Primary Contact"
        }
    }
]

start_time = time.time()
print("\n🚀 Starting processing...")

# Step 1: Load CSV
df = pd.read_csv(input_csv)
print(f"✅ Loaded input file: {input_csv}")
print("🔎 Columns:", df.columns.tolist())

# Step 2: Extract Subscription ID and Name from 'Account'
if parse_account_column and account_column_name in df.columns:
    print(f"🔧 Extracting 'Subscription ID' and 'Subscription Name' from '{account_column_name}'")
    df["Subscription ID"] = df[account_column_name].str.extract(r"^(\S+)\s*\(")[0].str.replace(r"\s+", "", regex=True)
    df["Subscription Name"] = df[account_column_name].str.extract(r"\((.*?)\)")[0].str.replace(r"\s+", "", regex=True)
else:
    print(f"⚠️ Skipping Account parsing – column not found or disabled.")

# Step 3: Clean 'Resource ID'
if resource_column_name in df.columns:
    print(f"🔧 Extracting filename from '{resource_column_name}'")
    df[resource_column_name] = df[resource_column_name].apply(lambda x: str(x).split('/')[-1])
else:
    print(f"⚠️ Resource column '{resource_column_name}' not found — skipping.")

# Step 4: Remove unwanted columns
existing_to_drop = [col for col in columns_to_remove if col in df.columns]
df.drop(columns=existing_to_drop, inplace=True)
print(f"🧹 Dropped columns: {existing_to_drop if existing_to_drop else 'None'}")

# Step 5: Add empty columns
print(f"➕ Adding blank columns: {columns_to_add}")
for col in columns_to_add:
    df[col] = ''

# Step 6: Perform mappings
for mapping in mappings:
    try:
        print(f"\n🔄 Performing mapping: {mapping['name']}")
        map_df = pd.read_excel(mapping["file"], sheet_name=mapping["sheet"])

        join_column = mapping["join_column"]
        df[join_column] = df[join_column].astype(str).str.strip()
        map_df[join_column] = map_df[join_column].astype(str).str.strip()

        columns_needed = [join_column] + list(mapping["source_to_target"].keys())
        df = df.merge(map_df[columns_needed], on=join_column, how="left", suffixes=('', '_map'))

        for src, target in mapping["source_to_target"].items():
            df[target] = df[src].fillna(f"{target} not available")

        df.drop(columns=list(mapping["source_to_target"].keys()), inplace=True)

        # Log unmatched values
        unmatched = df[df[list(mapping["source_to_target"].values())[0]] == f"{list(mapping['source_to_target'].values())[0]} not available"][join_column].dropna().unique()
        if len(unmatched) > 0:
            file_name = f"unmatched_{join_column.replace(' ', '_').lower()}.txt"
            with open(file_name, "w") as f:
                for value in unmatched:
                    f.write(f"{value}\n")
            print(f"📝 Unmatched values written to {file_name}")
        print(f"✅ Mapping complete.")

    except Exception as e:
        print(f"❌ Error in mapping '{mapping['name']}': {e}")

# Step 7: Save to Excel
df.to_excel(output_excel, index=False)
print(f"\n✅ Final file saved to: {output_excel}")
print(f"⏱️ Total time: {time.time() - start_time:.2f} seconds.")
