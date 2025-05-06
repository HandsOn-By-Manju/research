import pandas as pd
import os

# Configurable input Excel file name
input_file = 'your_input_file.xlsx'  # <-- Replace with your actual file name

# Output file name
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
    if 'id' not in df.columns or 'name' not in df.columns:
        missing_cols = [col for col in ['id', 'name'] if col not in df.columns]
        raise KeyError(f"Missing required column(s): {', '.join(missing_cols)}")

    # Remove '/subscription/' prefix from 'id' column
    df['id'] = df['id'].astype(str).str.replace('/subscription/', '', regex=False)

    # Rename columns
    df.rename(columns={'name': 'Subscription Name', 'id': 'Subscription ID'}, inplace=True)

    # Save to new Excel file
    df.to_excel(output_file, index=False)
    print(f"\n✅ Processed file saved as: {output_file}")

except FileNotFoundError as e:
    print(f"❌ Error: {e}")
except KeyError as e:
    print(f"❌ Error: {e}")
except Exception as e:
    print(f"❌ An unexpected error occurred: {e}")
