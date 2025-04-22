import pandas as pd
import os
from tkinter import Tk, filedialog

def find_int_strings_all_sheets(file_path):
    xls = pd.read_excel(file_path, sheet_name=None)

    print(f"\nScanning all sheets in: {file_path}")
    print("=" * 60)

    for sheet_name, df in xls.items():
        print(f"\nSheet: '{sheet_name}'")
        found = False
        for col in df.columns:
            for idx, val in df[col].items():
                if pd.notnull(val):
                    if isinstance(val, int):
                        print(f"  [INT]        Row {idx+1}, Column '{col}': {val}")
                        found = True
                    elif isinstance(val, str) and val.isdigit():
                        print(f"  [DIGIT-STR]  Row {idx+1}, Column '{col}': '{val}'")
                        found = True
        if not found:
            print("  No integer values found in this sheet.")

    print("\n" + "=" * 60)
    print("Scan complete.")

def select_excel_file():
    # Hide the main Tkinter window
    root = Tk()
    root.withdraw()

    # Ask the user to select an Excel file
    file_path = filedialog.askopenfilename(
        title="Select your Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    return file_path

def main():
    print("Please choose your Excel file from the file dialog...")
    file_path = select_excel_file()

    if not file_path:
        print("❌ No file selected. Exiting.")
        return

    if not os.path.isfile(file_path):
        print(f"❌ File not found: {file_path}")
        return

    try:
        find_int_strings_all_sheets(file_path)
    except Exception as e:
        print(f"⚠️ An error occurred: {e}")

if __name__ == "__main__":
    main()
