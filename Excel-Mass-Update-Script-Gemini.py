
# GEMINI PROMPTED RESPONSE: This script performs a mass update on an Excel sheet based on a predefined set of updates. It uses the pandas library to read and manipulate the Excel file. The script identifies rows in the "current shipments" sheet where the "Client PO" matches the keys in the updates dictionary and updates the corresponding "horse reg" and "trailer reg" columns accordingly. Finally, it saves the updated data to a new Excel file.
# Make sure to install pandas if you haven't already:
# pip install pandas openpyxl
# 1. Define your update data
# This dictionary uses the "Client PO" as the key

import pandas as pd
import sys

# 1. Define your update data
updates = {
    "BA3075": {"horse reg": "ZN 12345", "trailer reg": "TR 98765"},
    "BA3105": {"horse reg": "ZN 54321", "trailer reg": "TR 11223"},
    "BA3112": {"horse reg": "ZN 67890", "trailer reg": "TR 44556"},
    "BA3069": {"horse reg": "ZN 99887", "trailer reg": "TR 77889"},
    "BA3070": {"horse reg": "ZN 55443", "trailer reg": "TR 33221"},
}

file_path = r'C:\Users\Jason\Projects\TRACKING-FILE-UPDATE\BARTRAC - CONGO TRACKING 10-04-2026.xlsx'
target_sheet_name = 'current shipments'

try:
    # Use ExcelFile to inspect sheet names before loading
    xl = pd.ExcelFile(file_path)
    available_sheets = xl.sheet_names
    
    # Logic to find the sheet even if casing or spaces are off
    found_sheet = next((s for s in available_sheets if s.strip().lower() == target_sheet_name.lower()), None)

    if not found_sheet:
        print(f"Error: Could not find a sheet matching '{target_sheet_name}'")
        print(f"Available sheets in file: {available_sheets}")
        sys.exit(1)

    print(f"Loading sheet: '{found_sheet}'...")
    df = pd.read_excel(xl, sheet_name=found_sheet)

    # Standardize column names (strip spaces and lowercase for matching)
    df.columns = [str(col).strip().lower() for col in df.columns]

    # 3. Perform the mass update
    for po, values in updates.items():
        mask = df['client po'] == po
        
        if mask.any():
            df.loc[mask, 'horse reg'] = values['horse reg']
            df.loc[mask, 'trailer reg'] = values['trailer reg']
            print(f"Updated PO: {po}")
        else:
            print(f"Warning: PO {po} not found in the sheet.")

    # 4. Save the updated file
    output_file = 'shipments_updated.xlsx'
    df.to_excel(output_file, index=False)
    print(f"\nSuccess! Saved to '{output_file}'.")

except Exception as e:
    print(f"An unexpected error occurred: {e}")