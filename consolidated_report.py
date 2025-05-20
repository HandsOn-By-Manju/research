import os
import pandas as pd
import time
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

# --- CONFIGURATION ---
FILTER_SEVERITY_LIST = ["Informational", "Low"]
FILTER_POLICY_ID_LIST = ["XYZ-00123", "ABC-99999"]

def list_excel_files():
    excel_files = [f for f in os.listdir() if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
    print(f"\n🔍 Found {len(excel_files)} Excel file(s):\n")
    for f in excel_files:
        print(f"  - {f}")
    return excel_files

def read_columns(file):
    try:
        df = pd.read_excel(file)
        return list(df.columns), df
    except Exception as e:
        print(f"❌ Error reading file {file}: {e}")
        return [], None

def check_column_consistency(files):
    reference_columns, _ = read_columns(files[0])
    mismatch_files = []
    for file in files[1:]:
        current_columns, _ = read_columns(file)
        if current_columns != reference_columns:
            mismatch_files.append(file)
    return len(mismatch_files) == 0, reference_columns, mismatch_files

def merge_files(files, columns):
    merged_df = pd.DataFrame(columns=columns)
    for file in files:
        df = pd.read_excel(file)
        merged_df = pd.concat([merged_df, df], ignore_index=True)
        print(f"✅ Merged data from: {file} ({len(df)} rows)")
    return merged_df

def apply_excel_formatting(filename):
    print(f"🎨 Applying formatting to: {filename}")
    wb = load_workbook(filename)
    ws = wb.active

    ws.freeze_panes = "A2"
    header_fill = PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_length = max(max_length, len(str(cell.value)))
                cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2

    wb.save(filename)
    print(f"✅ Formatting completed: {filename}")

def format_time(seconds):
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    elif seconds < 3600:
        return f"{seconds/60:.2f} minutes"
    else:
        return f"{seconds/3600:.2f} hours"

def filter_rows(df):
    original_count = len(df)
    filtered_df = df.copy()

    if "Severity" in df.columns:
        severity_count = filtered_df[filtered_df["Severity"].isin(FILTER_SEVERITY_LIST)].shape[0]
        filtered_df = filtered_df[~filtered_df["Severity"].isin(FILTER_SEVERITY_LIST)]
        print(f"🗑️ Removed {severity_count} row(s) with Severity in {FILTER_SEVERITY_LIST}")
    else:
        print("⚠️ Column 'Severity' not found. Skipping severity filter.")

    if "Policy ID" in df.columns:
        policy_count = filtered_df[filtered_df["Policy ID"].isin(FILTER_POLICY_ID_LIST)].shape[0]
        filtered_df = filtered_df[~filtered_df["Policy ID"].isin(FILTER_POLICY_ID_LIST)]
        print(f"🗑️ Removed {policy_count} row(s) with Policy ID in {FILTER_POLICY_ID_LIST}")
    else:
        print("⚠️ Column 'Policy ID' not found. Skipping policy ID filter.")

    print(f"📉 Rows before filter: {original_count}, after filter: {len(filtered_df)}")
    return filtered_df

def add_summary_sheets(df, file_path, sheet_prefix=""):
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            # --- Overall Summary ---
            overall = df['Severity'].value_counts().reset_index()
            overall.columns = ['Severity', 'Count']
            overall.loc[len(overall.index)] = ['Total Findings', len(df)]
            overall.to_excel(writer, sheet_name=f"{sheet_prefix}Summary_Overall", index=False)

            # --- BU-wise Summary ---
            if 'BU' in df.columns:
                bu_summary = (
                    df.groupby(['BU', 'Severity'])
                    .size()
                    .unstack(fill_value=0)
                    .reset_index()
                )
                bu_summary["Total"] = bu_summary.iloc[:, 1:].sum(axis=1)
                bu_summary.to_excel(writer, sheet_name=f"{sheet_prefix}Summary_By_BU", index=False)
            else:
                print("⚠️ Column 'BU' not found. Skipping BU-wise summary.")
    except Exception as e:
        print(f"❌ Failed to write summary sheets to {file_path}: {e}")

def main():
    start_time = time.time()
    print("🚀 Starting Excel Merge Script with Multi-Filter + Summary...\n")

    files = list_excel_files()
    if not files:
        print("⚠️ No Excel files found. Exiting.")
        return

    print("\n🔎 Validating column consistency...\n")
    all_match, reference_columns, mismatch_files = check_column_consistency(files)

    if not all_match:
        print("❌ Column mismatch found in the following files:")
        for file in mismatch_files:
            print(f"  - {file}")
        print("🛑 Fix column issues and try again.")
        return

    print("\n📦 Merging files...\n")
    merged_df = merge_files(files, reference_columns)

    # Save merged
    merged_file = "merged_output.xlsx"
    merged_df.to_excel(merged_file, index=False)
    apply_excel_formatting(merged_file)
    add_summary_sheets(merged_df, merged_file, sheet_prefix="Merged_")
    print(f"\n📁 Merged file saved as: {merged_file}")

    # Filtered file
    print("\n🧹 Filtering based on Severity and Policy ID...\n")
    filtered_df = filter_rows(merged_df)

    filtered_file = "merged_filtered_output.xlsx"
    filtered_df.to_excel(filtered_file, index=False)
    apply_excel_formatting(filtered_file)
    add_summary_sheets(filtered_df, filtered_file, sheet_prefix="Filtered_")
    print(f"\n📁 Filtered file saved as: {filtered_file}")

    # Completion
    end_time = time.time()
    print(f"\n⏱️ Completed in {format_time(end_time - start_time)}.")

if __name__ == "__main__":
    main()
