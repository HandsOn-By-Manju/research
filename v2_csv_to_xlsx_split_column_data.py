import pandas as pd

# Step 1: Read the CSV file
csv_file_path = 'input_file.csv'  # Replace with your actual file path
df = pd.read_csv(csv_file_path)

# Step 2: List all columns
print("Original Columns:")
print(df.columns.tolist())

# Step 3: Process the 'Details' column if it exists
if 'Details' in df.columns:
    # Extract ID and Name, then strip all surrounding and internal whitespace
    df['ID'] = df['Details'].str.extract(r'^\s*(\S+)\s*\(')[0].str.replace(r'\s+', '', regex=True)
    df['Name'] = df['Details'].str.extract(r'\((.*?)\)')[0].str.replace(r'\s+', '', regex=True)

# Step 4: Save to Excel format
excel_file_path = 'output_file.xlsx'
df.to_excel(excel_file_path, index=False)

print("\nFinal Columns in Excel:")
print(df.columns.tolist())
print(f"\nExcel file saved successfully at: {excel_file_path}")
