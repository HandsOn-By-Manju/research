import pandas as pd
import re
import time

# ===== CONFIGURATION =====
INPUT_FILE = 'input.xlsx'              # Name of your input Excel file
OUTPUT_FILE = 'output_cleaned.xlsx'    # Name of the output file
SHEET_NAME = 'Sheet1'                  # Change if your sheet name is different
COLUMN_NAME = 'Instructions'           # Column with HTML content to be cleaned
# ==========================

start_time = time.time()

# Step 1: Load the Excel file
print(f"📄 Reading file: {INPUT_FILE}")
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

# Step 2: Define a function to clean HTML and convert to steps
def convert_html_to_steps(text):
    if pd.isna(text):
        return text  # Skip empty cells

    # Remove <ol> and </ol> tags
    text = re.sub(r'</?ol>', '', text, flags=re.IGNORECASE)

    # Extract all <li>content</li> blocks
    li_items = re.findall(r'<li>(.*?)</li>', text, flags=re.IGNORECASE | re.DOTALL)

    # Add "Step #n: " prefix and a '.' suffix to each list item
    steps = [f"Step #{i+1}: {item.strip()}." for i, item in enumerate(li_items)]

    # Combine into multiline string
    return "\n".join(steps)

# Step 3: Apply the transformation to the target column
if COLUMN_NAME not in df.columns:
    raise ValueError(f"❌ Column '{COLUMN_NAME}' not found in the sheet '{SHEET_NAME}'.")

print(f"🔧 Processing column: {COLUMN_NAME}")
df[COLUMN_NAME] = df[COLUMN_NAME].apply(convert_html_to_steps)

# Step 4: Save the updated DataFrame to a new Excel file
df.to_excel(OUTPUT_FILE, index=False)
end_time = time.time()

# Step 5: Print completion message
print(f"✅ Finished processing.")
print(f"💾 Output saved to '{OUTPUT_FILE}'")
print(f"⏱️ Time taken: {round(end_time - start_time, 2)} seconds")
