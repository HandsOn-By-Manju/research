import pandas as pd
import os
import time

# === CONFIGURATION ===
csv_file_path = 'input_data.csv'               # Input CSV file path
excel_file_path = 'output_data.xlsx'           # Output Excel file path

# Rename columns: {old_column_name: new_column_name}
columns_to_rename = {
    'EmpName': 'Employee Name',
    'Dept': 'Department'
}

# Add new columns with default values: {new_column_name: default_value}
columns_to_add = {
    'Reviewed': 'No',
    'Reviewer': ''
}

# Columns to split: {column_to_split: {'delimiter': str, 'new_columns': [new_col1, new_col2]}}
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

# Desired column order (only included columns will be reordered)
desired_column_order = [
    'Employee Name', 'Department', 'City', 'State', 'FirstName', 'LastName',
    'Reviewed', 'Reviewer'
]

# Columns to remove (AFTER all above steps)
columns_to_remove = ['UnwantedCol1', 'UnwantedCol2']  # Adjust as needed

# === START TIMER ===
start_time = time.time()

# === FILE CHECK ===
if not os.path.isfile(csv_file_path):
    print(f"❌ File not found: {csv_file_path}")
else:
    try:
        # === LOAD CSV ===
        df = pd.read_csv(csv_file_path)
        print(f"\n✅ Loaded CSV file: {csv_file_path}")

        # === LIST COLUMNS ===
        print("\n📋 Original Columns:")
        for col in df.columns:
            print(f" - {col}")
        print(f"🧮 Total columns: {len(df.columns)}")

        # === RENAME COLUMNS ===
        df.rename(columns=columns_to_rename, inplace=True)
        print("\n✏️ Renamed Columns:")
        for old, new in columns_to_rename.items():
            print(f" - {old} → {new}")

        # === ADD NEW COLUMNS ===
        for col, default in columns_to_add.items():
            df[col] = default
        print("\n➕ Added Columns:")
        for col in columns_to_add:
            print(f" - {col} (Default: {columns_to_add[col]})")

        # === SPLIT COLUMNS ===
        print("\n🔀 Splitting Columns:")
        for col, config in columns_to_split.items():
            if col in df.columns:
                split_cols = df[col].astype(str).str.split(config['delimiter'], n=1, expand=True)
                if len(split_cols.columns) < 2:
                    split_cols[1] = ''
                split_cols.columns = config['new_columns']
                df = pd.concat([df, split_cols], axis=1)
                print(f" - {col} → {config['new_columns'][0]}, {config['new_columns'][1]}")
            else:
                print(f" ⚠️ Column '{col}' not found to split")

        # === REORDER COLUMNS ===
        reordered = [col for col in desired_column_order if col in df.columns]
        remaining = [col for col in df.columns if col not in reordered]
        df = df[reordered + remaining]
        print("\n📐 Final Column Order:")
        for col in df.columns:
            print(f" - {col}")

        # === REMOVE COLUMNS ===
        removable = [col for col in columns_to_remove if col in df.columns]
        df.drop(columns=removable, inplace=True)
        if removable:
            print("\n🗑️ Removed Columns:")
            for col in removable:
                print(f" - {col}")

        # === EXPORT TO EXCEL ===
        df.to_excel(excel_file_path, index=False)
        print(f"\n💾 File saved to: {excel_file_path}")

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
