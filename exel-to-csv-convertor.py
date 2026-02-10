import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

def get_excel_file():
    """
    Open a file dialog and return the selected Excel path (or None if cancelled).
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main Tk window

    path = filedialog.askopenfilename(
        title="Select an Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    return path or None  # Empty string if cancelled

def excel_to_csv(excel_path: str) -> None:
    """
    Convert one Excel file to CSV (same folder, same name, .csv extension).
    """
    try:
        df = pd.read_excel(excel_path)
        if df.empty:
            print("Warning: The Excel file has no rows to export.")
            return

        csv_path = os.path.splitext(excel_path)[0] + ".csv"

        # Save with semicolon delimiter and UTF-8 BOM (often helps Excel read Unicode correctly)
        df.to_csv(csv_path, sep=";", index=False, encoding="utf-8-sig")

        print(f"Success: The CSV file has saved in {{{csv_path}}}")

    except PermissionError:
        print("Error: Permission denied.")
    except ValueError as error:
        print(f"Error: Could not read the Excel file: {error}")
    except Exception as error:
        print(f"Unexpected error: {error}")

try :
    excel_to_csv(get_excel_file())
except :
    print("Cancelled.")