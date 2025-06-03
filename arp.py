import pandas as pd
import os
import time

# === CONFIGURATION ===
csv_file_path = 'input_data.csv'        # <-- Replace with your CSV file
excel_file_path = 'output_data.xlsx'    # <-- Replace with your desired Excel output

# === START TIMER ===
start_time = time.time()

# === CHECK IF FILE EXISTS ===
if not os.path.isfile(csv_file_path):
    print(f"❌ File not found: {csv_file_path}")
else:
    try:
        # === READ CSV FILE ===
        df = pd.read_csv(csv_file_path)

        # === LIST COLUMNS ===
        print(f"\n✅ Loaded CSV file: {csv_file_path}")
        print("📌 Columns in the file:")
        for col in df.columns:
            print(f" - {col}")
        print(f"\n🧮 Total number of columns: {len(df.columns)}")

        # === WRITE TO EXCEL FILE ===
        df.to_excel(excel_file_path, index=False)
        print(f"\n✅ Successfully converted CSV to Excel: {excel_file_path}")

    except Exception as e:
        print(f"❌ Error: {e}")

# === END TIMER AND DISPLAY EXECUTION TIME ===
end_time = time.time()
duration = end_time - start_time

print("\n⏱️ Execution Time:")
if duration < 60:
    print(f" - {duration:.2f} seconds")
elif duration < 3600:
    minutes = duration // 60
    seconds = duration % 60
    print(f" - {int(minutes)} minutes {seconds:.2f} seconds")
else:
    hours = duration // 3600
    minutes = (duration % 3600) // 60
    seconds = duration % 60
    print(f" - {int(hours)} hours {int(minutes)} minutes {seconds:.2f} seconds")
