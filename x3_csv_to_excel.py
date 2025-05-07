import pandas as pd

# File paths
csv_file_path = "input_file.csv"
excel_file_path = "output_file.xlsx"

# Read the CSV file
df = pd.read_csv(csv_file_path)

# Check and process 'Account' column
if 'Account' in df.columns:
    # Extract new columns from 'Account'
    extracted = df['Account'].str.extract(r'([^()]+)\s*\(\s*([^()]+)\s*\)')
    extracted.columns = ['Subscription ID', 'Subscription Name']

    # Remove all spaces
    extracted['Subscription ID'] = extracted['Subscription ID'].str.replace(' ', '', regex=False)
    extracted['Subscription Name'] = extracted['Subscription Name'].str.replace(' ', '', regex=False)

    # Find index of 'Account' column
    account_index = df.columns.get_loc('Account')

    # Insert new columns right after 'Account'
    for i, col in enumerate(extracted.columns):
        df.insert(account_index + 1 + i, col, extracted[col])

    # Reorder columns: Account + extracted columns first, then others
    desired_order = ['Account', 'Subscription ID', 'Subscription Name'] + [
        col for col in df.columns if col not in ['Account', 'Subscription ID', 'Subscription Name']
    ]
    df = df[desired_order]
else:
    print("Column 'Account' not found in the CSV.")

# Save to Excel
df.to_excel(excel_file_path, index=False)

print(f"✅ Excel file saved as: {excel_file_path}")
print(f"📄 Columns in final file: {list(df.columns)}")
