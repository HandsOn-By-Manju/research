import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Configurable input Excel file name
input_file = 'your_input_file.xlsx'  # <-- Replace with your actual file name
output_file = 'Subscription_Details_Master_Data.xlsx'

try:
    # Check if file exists
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file '{input_file}' not found.")

    # Read the Excel file
    df = pd.read_excel(input_file)

    # List all columns
    print("Columns in the Excel file:")
    print(df.columns.tolist())

    # Ensure required columns exist
    required_cols = ['id', 'name']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        raise KeyError(f"Missing required column(s): {', '.join(missing_cols)}")

    # Remove '/subscription/' prefix from 'id' column
    df['id'] = df['id'].astype(str).str.replace('/subscription/', '', regex=False)

    # Rename columns
    df.rename(columns={'name': 'Subscription Name', 'id': 'Subscription ID'}, inplace=True)

    # Add new columns
    df['Environment'] = ''
    df['BU'] = ''

    # Set BU = 'Commerce' if 'Subscription Name' contains 'commerce' (case-insensitive)
    df['BU'] = df['Subscription Name'].astype(str).apply(
        lambda x: 'Commerce' if 'commerce' in x.lower() else '')

    # Override BU = 'Alpha' if 'PrimaryContact' is 'test@test.com' (case-insensitive)
    if 'PrimaryContact' in df.columns:
        df.loc[df['PrimaryContact'].astype(str).str.lower() == 'test@test.com', 'BU'] = 'Alpha'

    # Save to Excel
    df.to_excel(output_file, index=False)

    # Apply top and left alignment using openpyxl
    wb = load_workbook(output_file)
    ws = wb.active
    align_style = Alignment(vertical='top', horizontal='left')

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = align_style

    wb.save(output_file)
    print(f"\n✅ Processed file saved with alignment as: {output_file}")

except FileNotFoundError as e:
    print(f"❌ Error: {e}")
except KeyError as e:
    print(f"❌ Error: {e}")
except Exception as e:
    print(f"❌ An unexpected error occurred: {e}")
