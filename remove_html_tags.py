import pandas as pd
import re
import time

# ===== CONFIGURATION =====
INPUT_FILE = 'input.xlsx'
OUTPUT_FILE = 'output_cleaned.xlsx'
SHEET_NAME = 'Sheet1'
COLUMN_NAME = 'Instructions'
# ==========================

start_time = time.time()
print(f"📄 Reading file: {INPUT_FILE}")
df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

# Function to clean and convert HTML to steps
def convert_html_to_steps(text):
    if pd.isna(text):
        return text  # Keep blank cells untouched

    # Check for presence of <li> tags
    li_items = re.findall(r'<li>(.*?)</li>', text, flags=re.IGNORECASE | re.DOTALL)

    if li_items:
        # Convert <li> list items to step format
        steps = []
        for i, item in enumerate(li_items):
            # Remove all other HTML tags from item
            item_clean = re.sub(r'<[^>]+>', '', item).strip()
            item_clean = re.sub(r'\s+', ' ', item_clean)
            steps.append(f"Step #{i+1}: {item_clean}.")
        return "\n".join(steps)
    else:
        # No <li> present — just remove any HTML tags and preserve content
        return re.sub(r'<[^>]+>', '', text).strip()

# Apply to target column
if COLUMN_NAME not in df.columns:
    raise ValueError(f"❌ Column '{COLUMN_NAME}' not found in sheet '{SHEET_NAME}'.")

print(f"🔧 Cleaning column: {COLUMN_NAME}")
df[COLUMN_NAME] = df[COLUMN_NAME].apply(convert_html_to_steps)

# Save results
df.to_excel(OUTPUT_FILE, index=False)
end_time = time.time()

print(f"✅ Output saved to '{OUTPUT_FILE}'")
print(f"⏱️ Total time: {round(end_time - start_time, 2)} seconds")
