"""
This module allows the user to select an Excel file and scans all its sheets
for integer values and strings that represent digits.
"""

import os  # Standard library imports
from tkinter import Tk, filedialog  # GUI imports

import pandas as pd  # Third-party imports


def find_int_strings_all_sheets(file_path):
    """
    Scans all sheets in the given Excel file for integers and digit-strings.

    Parameters:
    file_path (str): The path to the Excel file to scan.
    """
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
    """
    Opens a file dialog for the user to select an Excel file.

    Returns:
    str: The path to the selected Excel file, or an empty string if cancelled.
    """
    root = Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select your Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    return file_path


def main():
    """
    Main entry point: prompts the user for an Excel file and scans it.
    """
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
    except pd.errors.ExcelFileError as e:
        print(f"⚠️ Failed to read Excel file: {e}")
    except Exception as e:  # Still broad, but fallback
        print(f"⚠️ An unexpected error occurred: {e}")


if __name__ == "__main__":
    main()
