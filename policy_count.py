import pandas as pd

# Define Excel file paths for each Business Unit
files = {
    'BU1': 'BU1.xlsx',
    'BU2': 'BU2.xlsx',
    'BU3': 'BU3.xlsx',
    'BU4': 'BU4.xlsx',
}

# Dictionary to store DataFrames of policy counts for each BU
bu_counts = {}

# Step 1: Read, clean, and count data from each BU file
for bu, file in files.items():
    df = pd.read_excel(file)

    # Clean: Trim whitespace and remove rows with blank Policy IDs
    df['Policy ID'] = df['Policy ID'].astype(str).str.strip()
    df['Policy Name'] = df['Policy Name'].astype(str).str.strip()
    df.dropna(subset=['Policy ID'], inplace=True)

    # Group by Policy ID and Name to count occurrences (don't drop duplicates)
    counts = df.groupby(['Policy ID', 'Policy Name']).size().reset_index(name=bu)
    bu_counts[bu] = counts

# Step 2: Merge BU-wise dataframes into one master dataframe
merged_df = bu_counts['BU1']
for bu in ['BU2', 'BU3', 'BU4']:
    merged_df = pd.merge(merged_df, bu_counts[bu], on=['Policy ID', 'Policy Name'], how='outer')

# Step 3: Detect and resolve Policy Name conflicts for the same Policy ID
name_conflict_df = merged_df.groupby('Policy ID')['Policy Name'].nunique().reset_index()
conflicted_ids = name_conflict_df[name_conflict_df['Policy Name'] > 1]['Policy ID'].tolist()

if conflicted_ids:
    print(f"⚠️ Warning: Conflicting Policy Names found for Policy IDs: {conflicted_ids}")

    # Sort by Policy Name and keep the first (alphabetically)
    merged_df.sort_values(by='Policy Name', inplace=True)

    # Group by Policy ID and keep the first name seen
    merged_df = merged_df.groupby('Policy ID', as_index=False).first()

# Step 4: Ensure all BU columns are present
for bu in ['BU1', 'BU2', 'BU3', 'BU4']:
    if bu not in merged_df:
        merged_df[bu] = 0

# Step 5: Replace NaNs with 0 and convert counts to integers
merged_df.fillna(0, inplace=True)
for bu in ['BU1', 'BU2', 'BU3', 'BU4']:
    merged_df[bu] = merged_df[bu].astype(int)

# Step 6: Add Total Count column
merged_df['Total Count'] = merged_df[['BU1', 'BU2', 'BU3', 'BU4']].sum(axis=1)

# Step 7: Save the output to Excel
output_file = "Combined_Policy_Counts.xlsx"
merged_df.to_excel(output_file, index=False)

print(f"✅ Output saved successfully to '{output_file}'")
