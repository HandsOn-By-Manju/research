import pandas as pd

# File paths
csv_file_path = "input_file.csv"
excel_file_path = "output_file.xlsx"

# Read the CSV file
df = pd.read_csv(csv_file_path)

# Step 1: Process 'Account' column to extract 'Subscription ID' and 'Subscription Name'
if 'Account' in df.columns:
    extracted = df['Account'].str.extract(r'([^()]+)\s*\(\s*([^()]+)\s*\)')
    extracted.columns = ['Subscription ID', 'Subscription Name']
    extracted['Subscription ID'] = extracted['Subscription ID'].str.replace(' ', '', regex=False)
    extracted['Subscription Name'] = extracted['Subscription Name'].str.replace(' ', '', regex=False)
    
    account_index = df.columns.get_loc('Account')
    for i, col in enumerate(extracted.columns):
        df.insert(account_index + 1 + i, col, extracted[col])
    
    # Reorder: Account, Subscription ID, Subscription Name, then the rest
    desired_order = ['Account', 'Subscription ID', 'Subscription Name'] + [
        col for col in df.columns if col not in ['Account', 'Subscription ID', 'Subscription Name']
    ]
    df = df[desired_order]
else:
    print("Column 'Account' not found in the CSV.")

# Step 2: Process 'Resource ID' to keep only the last segment (after the last '/')
if 'Resource ID' in df.columns:
    df['Resource ID'] = df['Resource ID'].astype(str).apply(
        lambda x: x.rstrip('/').rsplit('/', 1)[-1] if '/' in x.rstrip('/') else x
    )
else:
    print("Column 'Resource ID' not found in the CSV.")

# Save to Excel
df.to_excel(excel_file_path, index=False)

print(f"✅ Excel file saved as: {excel_file_path}")
print(f"📄 Final columns: {list(df.columns)}")
