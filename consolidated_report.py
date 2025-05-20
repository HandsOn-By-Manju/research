import os
import pandas as pd
import time
from openpyxl import load_workbook
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter

def list_excel_files():
    excel_files = [file for file in os.listdir() if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$')]
    print(f"\n🔍 Found {len(excel_files)} Excel file(s) in the current directory:\n")
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

    # Freeze top row
    ws.freeze_panes = "A2"

    # Style header row
    header_fill = PatternFill(start_color="87CEEB", end_color="87CEEB", fill_type="solid")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

    # Autofit column widths and apply alignment
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_length = max(max_length, len(str(cell.value)))
                cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(filename)
    print(f"✅ Formatting completed and saved: {filename}")

def format_time(seconds):
    if seconds < 60:
        return f"{seconds:.2f} seconds"
    elif seconds < 3600:
        return f"{seconds/60:.2f} minutes"
    else:
        return f"{seconds/3600:.2f} hours"

def main():
    start_time = time.time()
    print("🚀 Starting Excel Merge Script...\n")

    files = list_excel_files()
    if not files:
        print("⚠️ No Excel files found. Exiting.")
        return

    print("\n🔎 Checking if all Excel files have the same column names and order...\n")
    all_match, reference_columns, mismatch_files = check_column_consistency(files)

    if all_match:
        print("✅ All files have matching columns. Proceeding to merge...\n")
        merged_df = merge_files(files, reference_columns)
        output_file = "merged_output.xlsx"
        merged_df.to_excel(output_file, index=False)
        apply_excel_formatting(output_file)
        print(f"\n📁 Merged output saved as: {output_file}")
    else:
        print("❌ Column mismatch found in the following files:\n")
        for file in mismatch_files:
            print(f"  - {file}")
        print("\n🛑 Please fix column mismatch before merging. Exiting.")
        return

    end_time = time.time()
    elapsed = end_time - start_time
    print(f"\n⏱️ Script completed in {format_time(elapsed)}.")

if __name__ == "__main__":
    main()
