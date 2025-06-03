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

def convert_html_to_steps(text):
    if pd.isna(text):
        return text  # Leave blank cells unchanged

    # Find all <li>...</li> items
    li_items = re.findall(r'<li>(.*?)</li>', text, flags=re.IGNORECASE | re.DOTALL)

    if li_items:
        steps = []
        for i, item in enumerate(li_items):
            # Remove all inner HTML tags, keep content
            item_clean = re.sub(r'<[^>]+>', '', item).strip()
            item_clean = re.sub(r'\s+', ' ', item_clean)

            # Split at | and break into lines
            parts = [part.strip() for part in item_clean.split('|')]
            step_text = f"Step #{i+1}: {parts[0]}"
            if len(parts) > 1:
                extra_lines = "\n" + "\n".join(parts[1:])
                step_text += extra_lines
            steps.append(step_text + ".")
        return "\n".join(steps)
    else:
        # No <li> — clean all tags and split on |
        plain_text = re.sub(r'<[^>]+>', '', text).strip()
        parts = [part.strip() for part in plain_text.split('|')]
        return "\n".join(parts) + "."

# Apply the cleaning function
if COLUMN_NAME not in df.columns:
    raise ValueError(f"❌ Column '{COLUMN_NAME}' not found in sheet '{SHEET_NAME}'.")

print(f"🔧 Cleaning column: {COLUMN_NAME}")
df[COLUMN_NAME] = df[COLUMN_NAME].apply(convert_html_to_steps)

# Save to new Excel file
df.to_excel(OUTPUT_FILE, index=False)
end_time = time.time()

print(f"✅ Output saved to '{OUTPUT_FILE}'")
print(f"⏱️ Completed in {round(end_time - start_time, 2)} seconds")
