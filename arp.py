import pandas as pd
import os
import time

# === CONFIGURATION ===
csv_file_path = 'input_data.csv'               # Input CSV path
excel_file_path = 'output_data.xlsx'           # Output Excel path

columns_to_rename = {
    'EmpName': 'Employee Name',
    'Dept': 'Department'
}

columns_to_add = {
    'Reviewed': 'No',
    'Reviewer': ''
}

columns_to_split = {
    'Location': {
        'delimiter': ',',
        'new_columns': ['City', 'State']
    },
    'FullName': {
        'delimiter': ' ',
        'new_columns': ['FirstName', 'LastName']
    }
}

desired_column_order = [
    'Employee Name', 'Department', 'City', 'State', 'FirstName', 'LastName', 'Reviewed', 'Reviewer'
]

columns_to_remove = ['UnwantedCol1', 'UnwantedCol2']

# Rows to remove based on filter: {column_name: [value1, value2]}
rows_to_filter_out = {
    'Department': ['HR', 'Finance'],
    'Reviewed': ['No']
}

# === START TIMER ===
start_time = time.time()

if not os.path.isfile(csv_file_path):
    print(f"❌ File not found: {csv_file_path}")
else:
    try:
        # === READ CSV ===
        df = pd.read_csv(csv_file_path)
        print(f"\n✅ Loaded CSV: {csv_file_path}")

        # === LIST ORIGINAL COLUMNS ===
        print("\n📋 Columns in file:")
        for col in df.columns:
            print(f" - {col}")
        print(f"🧮 Total columns: {len(df.columns)}")

        # === RENAME COLUMNS ===
        df.rename(columns=columns_to_rename, inplace=True)
        print("\n✏️ Renamed Columns:")
        for old, new in columns_to_rename.items():
            print(f" - {old} → {new}")

        # === ADD NEW COLUMNS ===
        for col, default_val in columns_to_add.items():
            df[col] = default_val
        print("\n➕ Added Columns:")
        for col in columns_to_add:
            print(f" - {col} (Default: {columns_to_add[col]})")

        # === SPLIT COLUMNS ===
        print("\n🔀 Splitting Columns:")
        for col, cfg in columns_to_split.items():
            if col in df.columns:
                split_df = df[col].astype(str).str.split(cfg['delimiter'], n=1, expand=True)
                if len(split_df.columns) < 2:
                    split_df[1] = ''
                split_df.columns = cfg['new_columns']
                df = pd.concat([df, split_df], axis=1)
                print(f" - {col} → {cfg['new_columns'][0]}, {cfg['new_columns'][1]}")
            else:
                print(f" ⚠️ Column '{col}' not found for splitting.")

        # === REMOVE SPECIFIC COLUMNS ===
        to_remove = [col for col in columns_to_remove if col in df.columns]
        df.drop(columns=to_remove, inplace=True)
        if to_remove:
            print("\n🗑️ Removed Columns:")
            for col in to_remove:
                print(f" - {col}")

        # === FILTER ROWS OUT BASED ON COLUMN VALUES ===
        print("\n🚫 Filtering Rows:")
        for col, values in rows_to_filter_out.items():
            if col in df.columns:
                before = len(df)
                df = df[~df[col].isin(values)]
                after = len(df)
                print(f" - Removed {before - after} rows where '{col}' in {values}")
            else:
                print(f" ⚠️ Column '{col}' not found for filtering.")

        # === REORDER COLUMNS ===
        reordered = [col for col in desired_column_order if col in df.columns]
        remaining = [col for col in df.columns if col not in reordered]
        df = df[reordered + remaining]
        print("\n📐 Final Column Order:")
        for col in df.columns:
            print(f" - {col}")

        # === SAVE TO EXCEL ===
        df.to_excel(excel_file_path, index=False)
        print(f"\n💾 Excel saved: {excel_file_path}")

    except Exception as e:
        print(f"❌ Error: {e}")

# === END TIMER ===
end_time = time.time()
duration = end_time - start_time

print("\n⏱️ Execution Time:")
if duration < 60:
    print(f" - {duration:.2f} seconds")
elif duration < 3600:
    print(f" - {int(duration // 60)} minutes {duration % 60:.2f} seconds")
else:
    h = int(duration // 3600)
    m = int((duration % 3600) // 60)
    s = duration % 60
    print(f" - {h} hours {m} minutes {s:.2f} seconds")
