import pandas as pd

# File paths
csv_file = 'file1.csv'
excel_file = 'file2.xlsx'

# Read files
df_csv = pd.read_csv(csv_file)
df_excel = pd.read_excel(excel_file)

# Normalize column names to lowercase
df_csv.columns = df_csv.columns.str.lower()
df_excel.columns = df_excel.columns.str.lower()

# Define the lowercase version of the column to look for
subscription_column = 'subscription id'

# Check if column exists in both
if subscription_column in df_csv.columns and subscription_column in df_excel.columns:
    # Normalize Subscription ID values
    df_csv[subscription_column] = df_csv[subscription_column].astype(str).str.strip()
    df_excel[subscription_column] = df_excel[subscription_column].astype(str).str.strip()

    # Extract unique IDs
    ids_csv = set(df_csv[subscription_column])
    ids_excel = set(df_excel[subscription_column])

    # Find matches and differences
    common_ids = ids_csv & ids_excel
    only_in_csv = ids_csv - ids_excel
    only_in_excel = ids_excel - ids_csv

    # Filter rows
    df_common = df_csv[df_csv[subscription_column].isin(common_ids)]
    df_only_in_csv = df_csv[df_csv[subscription_column].isin(only_in_csv)]
    df_only_in_excel = df_excel[df_excel[subscription_column].isin(only_in_excel)]

    # Save results
    df_common.to_excel('matched_subscriptions.xlsx', index=False)
    df_only_in_csv.to_excel('only_in_csv.xlsx', index=False)
    df_only_in_excel.to_excel('only_in_excel.xlsx', index=False)

    # Console output
    print(f"✅ Total Common IDs: {len(common_ids)}")
    print(f"✅ Only in CSV: {len(only_in_csv)}")
    print(f"✅ Only in Excel: {len(only_in_excel)}")
else:
    print("❌ 'Subscription ID' column not found in one or both files (case-insensitive match failed).")
