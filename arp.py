import pandas as pd
import os
import time

# === CONFIGURATION ===
csv_file_path = 'input_data.csv'               # Input CSV
excel_file_path = 'output_data.xlsx'           # Output Excel

columns_to_rename = {
    'EmpName': 'Employee Name',
    'Dept': 'Department'
}

columns_to_remove = ['Age', 'Salary']  # Columns to remove

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

desired_column_order = [  # Final column order in Excel
    'Employee Name', 'Department', 'City', 'State', 'FirstName', 'LastName',
    'Reviewed', 'Reviewer'
]

# === START TIMER ===
start_time = time.time()

if not os.path.isfile(csv_file_path):
    print(f"❌ File not found: {csv_file_path}")
else:
    try:
        # === LOAD CSV FILE ===
        df = pd.read_csv(csv_file_path)
        print(f"\n✅ Loaded CSV file: {csv_file_path}")

        # === LIST ORIGINAL COLUMNS ===
        print("\n📌 Original Columns:")
        for col in df.columns:
            print(f" - {col}")
        print(f"🧮 Total columns: {len(df.columns)}")

        # === RENAME COLUMNS ===
        df.rename(columns=columns_to_rename, inplace=True)
        print("\n✏️ Renamed Columns:")
        for old, new in columns_to_rename.items():
            print(f" - {old} → {new}")

        # === REMOVE COLUMNS ===
        removable = [col for col in columns_to_remove if col in df.columns]
        df.drop(columns=removable, inplace=True)
        print("\n🗑️ Removed Columns:")
        for col in removable:
            print(f" - {col}")

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
                print(f" ⚠️ Column '{col}' not found for splitting.")

        # === REARRANGE COLUMNS ===
        existing_order = [col for col in desired_column_order if col in df.columns]
        remaining_cols = [col for col in df.columns if col not in existing_order]
        final_order = existing_order + remaining_cols
        df = df[final_order]

        print("\n📐 Final Column Order:")
        for col in final_order:
            print(f" - {col}")

        # === EXPORT TO EXCEL ===
        df.to_excel(excel_file_path, index=False)
        print(f"\n📁 Saved to Excel: {excel_file_path}")

    except Exception as e:
        print(f"❌ Error occurred: {e}")

# === END TIMER & EXECUTION TIME ===
end_time = time.time()
duration = end_time - start_time

print("\n⏱️ Execution Time:")
if duration < 60:
    print(f" - {duration:.2f} seconds")
elif duration < 3600:
    print(f" - {int(duration // 60)} minutes {duration % 60:.2f} seconds")
else:
    hours = int(duration // 3600)
    minutes = int((duration % 3600) // 60)
    seconds = duration % 60
    print(f" - {hours} hours {minutes} minutes {seconds:.2f} seconds")
