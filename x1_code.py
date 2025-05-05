import pandas as pd
import json
import time

# === Load configuration ===
with open("config.json") as f:
    config = json.load(f)

input_csv = config["input_csv"]
output_excel = config["output_excel"]
columns_to_remove = config.get("columns_to_remove", [])
columns_to_add = config.get("columns_to_add", [])
parse_account = config.get("parse_account_column", True)
account_col = config.get("account_column_name", "Account")
resource_col = config.get("resource_column_name", "Resource ID")
mappings = config.get("mappings", [])

start_time = time.time()
print("\n🚀 Starting processing using config file...")

# Step 1: Load CSV
df = pd.read_csv(input_csv)
print(f"✅ Loaded input file: {input_csv}")
print("🔎 Initial columns:", df.columns.tolist())

# Step 2: Extract Subscription ID and Name from 'Account'
if parse_account and account_col in df.columns:
    print(f"🔧 Extracting 'Subscription ID' and 'Subscription Name' from '{account_col}'")
    df["Subscription ID"] = df[account_col].str.extract(r"^(\S+)\s*\(")[0].str.replace(r"\s+", "", regex=True)
    df["Subscription Name"] = df[account_col].str.extract(r"\((.*?)\)")[0].str.replace(r"\s+", "", regex=True)
else:
    print(f"⚠️ Skipping account column parsing. '{account_col}' not found or disabled in config.")

# Step 3: Clean 'Resource ID'
if resource_col in df.columns:
    print(f"🔧 Extracting filename from '{resource_col}'")
    df[resource_col] = df[resource_col].apply(lambda x: str(x).split('/')[-1])
else:
    print(f"⚠️ Resource column '{resource_col}' not found — skipping filename extraction.")

# Step 4: Drop unwanted columns
existing_to_drop = [col for col in columns_to_remove if col in df.columns]
df.drop(columns=existing_to_drop, inplace=True)
print(f"🧹 Dropped columns: {existing_to_drop if existing_to_drop else 'None'}")

# Step 5: Add new blank columns
print(f"➕ Adding blank columns: {columns_to_add}")
for col in columns_to_add:
    df[col] = ''

# Step 6: Perform mappings from external sheets
for mapping in mappings:
    try:
        print(f"🔄 Performing: {mapping.get('name', 'Unnamed Mapping')}")
        mapping_file = mapping["file"]
        mapping_sheet = mapping["sheet"]
        join_column = mapping["join_column"]
        source_to_target = mapping["source_to_target"]

        map_df = pd.read_excel(mapping_file, sheet_name=mapping_sheet)

        # Strip and normalize join columns
        df[join_column] = df[join_column].astype(str).str.strip()
        map_df[join_column] = map_df[join_column].astype(str).str.strip()

        columns_needed = [join_column] + list(source_to_target.keys())
        df = df.merge(map_df[columns_needed], on=join_column, how="left", suffixes=('', '_map'))

        # Fill target fields with fallback text and drop source columns
        for src_col, target_col in source_to_target.items():
            df[target_col] = df[src_col].fillna(f"{target_col} not available")
        df.drop(columns=list(source_to_target.keys()), inplace=True)

        # Log unmatched join_column values
        unmatched = df[df[list(source_to_target.values())[0]] == f"{list(source_to_target.values())[0]} not available"][join_column].dropna().unique()
        if len(unmatched) > 0:
            filename = f"unmatched_{join_column.replace(' ', '_').lower()}.txt"
            with open(filename, "w") as f:
                for value in unmatched:
                    f.write(f"{value}\n")
            print(f"📝 Unmatched {join_column} values written to {filename}")

        print(f"✅ {mapping['name']} complete.\n")

    except Exception as e:
        print(f"❌ Error in mapping '{mapping.get('name', 'Unnamed Mapping')}': {e}")

# Step 7: Save final result
df.to_excel(output_excel, index=False)
print(f"✅ Final Excel saved as '{output_excel}' in {time.time() - start_time:.2f} seconds.")
