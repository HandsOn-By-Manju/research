import pandas as pd

# Step 1: Read the CSV file
csv_file_path = 'input_file.csv'  # Replace with your actual file name
df = pd.read_csv(csv_file_path)

# Step 2: List all the columns
print("Columns in the CSV file:")
for col in df.columns:
    print(col)

# Step 3: Save as Excel file
excel_file_path = 'output_file.xlsx'  # Output file path
df.to_excel(excel_file_path, index=False)  # index=False to avoid saving DataFrame index as a column

print(f"\nExcel file saved successfully as: {excel_file_path}")
