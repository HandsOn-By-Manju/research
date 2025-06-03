import pandas as pd
import re
import time

# ===== CONFIGURATION =====
INPUT_FILE = 'input.xlsx'              # Path to your input Excel file
OUTPUT_FILE = 'output_cleaned.xlsx'    # File to save cleaned results
SHEET_NAME = 'Sheet1'                  # Sheet name or use index 0
COLUMN_NAME = 'Instructions'           # Column containing HTML content
# ==========================

start_time = time.time()
print(f"📄 Reading file: {INPUT_FILE}")
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

# Step 1: Define cleaning function
def convert_html_to_steps(text):
    if pd.isna(text):  # Skip empty cells
        return text

    # Step 2: Remove <ol>, </ol>, <ul>, </ul> tags
    text = re.sub(r'</?(ol|ul)>', '', text, flags=re.IGNORECASE)

    # Step 3: Find all <li>...</li> items (steps)
    li_items = re.findall(r'<li>(.*?)</li>', text, flags=re.IGNORECASE | re.DOTALL)

    steps = []
    for i, item in enumerate(li_items):
        # Step 4: Remove <a>, <pre>, <code> tags but keep the content
        item_clean = re.sub(r'</?(a|pre|code)[^>]*>', '', item, flags=re.IGNORECASE)

        # Step 5: Clean up extra whitespaces
        item_clean = re.sub(r'\s+', ' ', item_clean).strip()

        # Step 6: Add "Step #n: ..." with period
        steps.append(f"Step #{i+1}: {item_clean}.")

    # Step 7: Join all steps into one string with line breaks
    return "\n".join(steps)

# Step 8: Apply to target column
if COLUMN_NAME not in df.columns:
    raise ValueError(f"❌ Column '{COLUMN_NAME}' not found in sheet '{SHEET_NAME}'.")

print(f"🔧 Processing column: {COLUMN_NAME}")
df[COLUMN_NAME] = df[COLUMN_NAME].apply(convert_html_to_steps)

# Step 9: Save to Excel
df.to_excel(OUTPUT_FILE, index=False)
end_time = time.time()

# Step 10: Final logs
print(f"✅ Completed! Output saved to '{OUTPUT_FILE}'")
print(f"⏱️ Time taken: {round(end_time - start_time, 2)} seconds")
