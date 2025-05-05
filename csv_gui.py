import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import time

def process_files():
    try:
        input_file = filedialog.askopenfilename(title="Select Input CSV", filetypes=[("CSV Files", "*.csv")])
        if not input_file:
            return

        start_time = time.time()
        df = pd.read_csv(input_file)

        if 'Account' in df.columns:
            df['Subscription ID'] = df['Account'].str.extract(r'^(\S+)\s*\(')[0].str.replace(r'\s+', '', regex=True)
            df['Subscription Name'] = df['Account'].str.extract(r'\((.*?)\)')[0].str.replace(r'\s+', '', regex=True)

        if 'Resource ID' in df.columns:
            df['Resource ID'] = df['Resource ID'].apply(lambda x: str(x).split('/')[-1])

        columns_to_remove = [
            "DummyColumn1", "DummyColumn2", "DummyColumn3",
            "DummyColumn4", "DummyColumn5", "DummyColumn6",
            "DummyColumn7", "DummyColumn8", "DummyColumn9"
        ]
        df.drop(columns=[col for col in columns_to_remove if col in df.columns], inplace=True)

        columns_to_add = ["Col1", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8"]
        for col in columns_to_add:
            df[col] = ''

        df.to_excel("output_step5.xlsx", index=False)

        elapsed = time.time() - start_time
        messagebox.showinfo("Success", f"✅ Process completed in {elapsed:.2f} seconds.\nSaved as output_step5.xlsx")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# --- GUI setup ---
root = tk.Tk()
root.title("CSV Processor - Step 1 to 5")

label = tk.Label(root, text="Click below to select CSV and run processing:", padx=20, pady=10)
label.pack()

run_button = tk.Button(root, text="Run CSV Processor", command=process_files, padx=20, pady=10, bg="lightblue")
run_button.pack()

root.mainloop()
