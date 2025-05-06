import pandas as pd
import os

# Configurable input Excel file name
input_file = 'your_input_file.xlsx'  # <-- Replace with your actual file name
output_file = 'Subscription_Details_Master_Data.xlsx'

try:
    # Step 1: Check if file exists
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file '{input_file}' not found.")

    # Step 2: Read the Excel file
    df = pd.read_excel(input_file)

    # Step 3: List all columns
    print("Columns in the Excel file:")
    print(df.columns.tolist())

    # Step 4: Ensure required columns exist
    if 'id' not in df.columns or 'name' not in df.columns:
        missing_cols = [col for col in ['id', 'name'] if col not in df.columns]
        raise KeyError(f"Missing required column(s): {', '.join(missing_cols)}")

    # Step 5: Remove '/subscription/' prefix from 'id'
    df['id'] = df['id'].astype(str).str.replace('/subscription/', '', regex=False)

    # Step 6: Rename columns
    df.rename(columns={'name': 'Subscription Name', 'id': 'Subscription ID'}, inplace=True)

    # Step 7: Add 'Environment' and 'BU' columns
    df['Environment'] = ''
    df['BU'] = df['Subscription Name'].astype(str).apply(
        lambda x: 'Commerce' if 'commerce' in x.lower() else '')

    # Step 8: Save to new Excel file
    df.to_excel(output_file, index=False)
    print(f"\n✅ Processed file saved as: {output_file}")

except FileNotFoundError as e:
    print(f"❌ Error: {e}")
except KeyError as e:
    print(f"❌ Error: {e}")
except Exception as e:
    print(f"❌ An unexpected error occurred: {e}")
