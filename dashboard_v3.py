import pandas as pd
import time

# ========== CONFIGURATION ==========
input_file = 'input.xlsx'          # Input Excel file
output_file = 'output.xlsx'        # Output Excel file

columns_to_add = ['NewColumn1', 'NewColumn2']
columns_to_rename = {
    'OldName1': 'RenamedColumn1',
    'OldName2': 'RenamedColumn2'
}
columns_to_delete = ['Unwanted1', 'Unwanted2']
final_column_order = ['RenamedColumn1', 'NewColumn1', 'Severity', 'SomeExistingColumn']
# ====================================

# Start timer
start_time = time.time()
print("▶ Starting the Excel transformation process...\n")

# Step 1: Load Excel file
try:
    df = pd.read_excel(input_file)
    print(f"✅ Successfully loaded '{input_file}'\n")
except Exception as e:
    print(f"❌ Failed to load Excel file: {e}")
    exit()

# Step 2: Show original columns
print("📋 Original columns:")
print(df.columns.tolist())

# Step 3: Add new columns
print("\n➕ Adding new columns...")
for col in columns_to_add:
    if col not in df.columns:
        df[col] = ''
        print(f"   - Added: {col}")
    else:
        print(f"   - Skipped (already exists): {col}")

# Step 4: Rename columns
print("\n✏️ Renaming columns...")
for old_name, new_name in columns_to_rename.items():
    if old_name in df.columns:
        df.rename(columns={old_name: new_name}, inplace=True)
        print(f"   - Renamed: '{old_name}' ➝ '{new_name}'")
    else:
        print(f"   - Skipped (not found): {old_name}")

# Step 5: Delete columns
print("\n🗑️ Deleting columns...")
for col in columns_to_delete:
    if col in df.columns:
        df.drop(columns=col, inplace=True)
        print(f"   - Deleted: {col}")
    else:
        print(f"   - Skipped (not found): {col}")

# Step 6: Rearranging columns
print("\n🔀 Rearranging columns...")
existing_final_order = [col for col in final_column_order if col in df.columns]
remaining_cols = [col for col in df.columns if col not in existing_final_order]
df = df[existing_final_order + remaining_cols]
print(f"   - Final column order set to: {existing_final_order + remaining_cols}")

# Step 7: Severity count summary
print("\n📊 Count of rows per Severity category:")
if 'Severity' in df.columns:
    severity_counts = df['Severity'].value_counts(dropna=False)
    print(severity_counts)

    # Step 8: Filter and delete 'Informational' severity rows
    info_rows = df[df['Severity'] == 'Informational']
    info_count = len(info_rows)
    df = df[df['Severity'] != 'Informational']
    print(f"\n❌ Deleted rows with 'Informational' severity: {info_count}")
else:
    print("   - 'Severity' column not found. Skipping severity filtering.")

# Step 9: Save final Excel
df.to_excel(output_file, index=False)
print(f"\n💾 Excel file saved as '{output_file}'")

# Step 10: Execution time
end_time = time.time()
elapsed = end_time - start_time
print("\n⏱️ Execution Time:", end=' ')
if elapsed < 60:
    print(f"{elapsed:.2f} seconds")
elif elapsed < 3600:
    print(f"{elapsed / 60:.2f} minutes")
else:
    print(f"{elapsed / 3600:.2f} hours")

print("\n✅ Process completed successfully.")
