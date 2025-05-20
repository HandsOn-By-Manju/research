import pandas as pd
import os
import webbrowser

# === Configuration ===
excel_file_path = "your_excel_file.xlsx"               # <-- Input Excel file path
sheet_name = 0                                          # <-- Sheet index or name
filter_column_name = "Severity"                        # <-- Column to filter/group by
target_column_name = "Subscription ID"                 # <-- Column to extract values from

# Per-item prefix/suffix
prefix = "prefix_data_for_each_subscriptions"
suffix = "suffix_data_for_each_subscriptions"

# Final base URL and suffix
final_base_url = "https://example.com/trigger/"
final_suffix = "_FINAL_SUFFIX"

# Output file names
text_file = "Generated_URL_Report.txt"
markdown_file = "Generated_URL_Report.md"
html_file = "Generated_URL_Report.html"

# === Step 1: Load Excel ===
print("[INFO] Loading Excel file...")
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# === Step 2: Validate Columns ===
missing = [col for col in [filter_column_name, target_column_name] if col not in df.columns]
if missing:
    print(f"[ERROR] Missing columns in input file: {', '.join(missing)}")
    exit()

# === Step 3: Process Each Unique Filter Value ===
unique_values = df[filter_column_name].dropna().astype(str).unique().tolist()
print(f"[INFO] Found {len(unique_values)} unique values in '{filter_column_name}' column.\n")

txt_output = []
md_output = ["# Generated URL Report\n"]
html_output = [
    "<!DOCTYPE html><html><head><meta charset='UTF-8'>",
    "<title>Generated URL Report</title></head><body>",
    "<h1>Generated URL Report</h1>"
]

for filter_val in unique_values:
    print(f"[SECTION] Processing: {filter_column_name} = '{filter_val}'")

    # Filter rows
    filtered_df = df[df[filter_column_name].astype(str).str.strip().str.lower() ==
                     filter_val.strip().lower()]
    values = filtered_df[target_column_name].dropna().astype(str).unique().tolist()
    count = len(values)

    if count == 0:
        print(f"  [WARN] No '{target_column_name}' values found for '{filter_val}'")
        continue

    # Build URL string
    formatted_items = []
    for i, val in enumerate(values):
        item = f"{prefix}{val}"
        if i < len(values) - 1:
            item += suffix
        formatted_items.append(item)

    combined_string = "".join(formatted_items)
    full_url = f"{final_base_url}{combined_string}{final_suffix}"

    # Console output
    print(f"  [INFO] Unique '{target_column_name}' values: {count}")
    print(f"  [RESULT] URL: {full_url}\n")

    # Text output
    txt_output.extend([
        f"[Filter Value: {filter_val}]",
        f"Subscription Count: {count}",
        f"URL: {full_url}",
        ""
    ])

    # Markdown output
    md_output.extend([
        f"## {filter_column_name}: {filter_val}",
        f"- Subscription Count: **{count}**",
        f"- [Open URL]({full_url})",
        ""
    ])

    # HTML output
    html_output.extend([
        f"<h2>{filter_column_name}: {filter_val}</h2>",
        f"<p><b>Subscription Count:</b> {count}</p>",
        f"<p><a href='{full_url}' target='_blank'>Open URL</a></p><hr>"
    ])

# === Step 4: Write Output Files ===
with open(text_file, "w", encoding="utf-8") as f:
    f.write("\n".join(txt_output))

with open(markdown_file, "w", encoding="utf-8") as f:
    f.write("\n".join(md_output))

with open(html_file, "w", encoding="utf-8") as f:
    html_output.append("</body></html>")
    f.write("\n".join(html_output))

print(f"[SUCCESS] Reports written to:\n - {text_file}\n - {markdown_file}\n - {html_file}")

# === Step 5: Open HTML in Default Browser ===
print("[INFO] Opening HTML report in your default web browser...")
webbrowser.open('file://' + os.path.realpath(html_file))
