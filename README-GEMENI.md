To handle a mass update of specific rows in an Excel workbook based on a unique identifier like a **Client PO**, Python with the `pandas` library is the most efficient and robust approach. This script will allow you to map your list of POs to their new registration details and update the workbook in one go.

### Python Mass Update Script

This script assumes your workbook is named `shipments.xlsx` and the sheet is named `current shipments`.

```python
import pandas as pd

# 1. Define your update data
# This dictionary uses the "Client PO" as the key
updates = {
    "BA3075": {"horse reg": "ZN 12345", "trailer reg": "TR 98765"},
    "BA3105": {"horse reg": "ZN 54321", "trailer reg": "TR 11223"},
    "BA3112": {"horse reg": "ZN 67890", "trailer reg": "TR 44556"},
    "BA3069": {"horse reg": "ZN 99887", "trailer reg": "TR 77889"},
    "BA3070": {"horse reg": "ZN 55443", "trailer reg": "TR 33221"},
}

file_path = 'shipments.xlsx'
sheet_name = 'current shipments'

# 2. Load the workbook
df = pd.read_excel(file_path, sheet_name=sheet_name)

# 3. Perform the mass update
# We iterate through the updates and use .loc to find the matching Client PO
for po, values in updates.items():
    # Find the mask (boolean index) where the Client PO matches
    mask = df['client po'] == po
    
    if mask.any():
        df.loc[mask, 'horse reg'] = values['horse reg']
        df.loc[mask, 'trailer reg'] = values['trailer reg']
        print(f"Updated PO: {po}")
    else:
        print(f"Warning: PO {po} not found in the sheet.")

# 4. Save the updated file
# You can overwrite the original or save as a new file to be safe
df.to_excel('shipments_updated.xlsx', index=False)
print("Mass update complete. Saved to 'shipments_updated.xlsx'.")
```

---

### How it Works

* **The Update Mapping:** Instead of searching manually, we use a dictionary. This makes it easy to scale—you could even load this "update list" from a separate CSV or clipboard later.
* **Vectorized Searching:** The `df['client po'] == po` line creates a "mask." If a PO appears multiple times in your sheet (e.g., across multiple lines for different items), this script will update **all** occurrences for that PO automatically.
* **Data Integrity:** By using `.loc`, we ensure we are writing specifically to the columns you named (`horse reg` and `trailer reg`) without disturbing other data in those rows.

### Key Prerequisites
To run this, you will need the `pandas` and `openpyxl` libraries installed in your environment:

```bash
pip install pandas openpyxl
```

### A Skeptical Note on Data Cleaning
Before running this mass update, ensure your **Client PO** column doesn't have leading or trailing spaces (e.g., `"BA3075 "` vs `"BA3075"`). If the script reports that POs are "not found" even though they exist, you might want to add `df['client po'] = df['client po'].str.strip()` after loading the Excel file to sanitize the identifiers.


The error is clear: the script is looking for a column exactly named `horse reg`, but your Excel sheet actually uses `horse reg no` and `trailer reg no`.

Additionally, the column names in your file have some messy formatting (extra spaces and non-breaking characters). I have updated the script to map to the correct "no" columns and added a "fuzzy" header matcher to handle those invisible spaces in your column names.

### Updated Formatting-Safe Script

```python
import openpyxl

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

wb = openpyxl.load_workbook(file_path)
found_sheet = next((s for s in wb.sheetnames if s.strip().lower() == target_sheet_name.lower()), None)

if not found_sheet:
    print(f"Error: Sheet '{target_sheet_name}' not found.")
    exit()

ws = wb[found_sheet]

# 2. Advanced Header Mapping (Handles 'horse reg no' and extra spaces)
header_map = {}
for cell in ws[1]:
    if cell.value:
        # Clean the header name: lowercase, strip spaces, remove internal double spaces
        clean_header = " ".join(str(cell.value).lower().split())
        header_map[clean_header] = cell.column

# Debugging: show what we mapped
print(f"Mapped 'client po' to column: {header_map.get('client po')}")
print(f"Mapped 'horse reg no' to column: {header_map.get('horse reg no')}")

# 3. Perform the update
update_count = 0

# Check for the correct column keys based on your output
col_po = header_map.get('client po')
col_horse = header_map.get('horse reg no')  # Updated to match your actual sheet
col_trailer = header_map.get('trailer reg no') # Updated to match your actual sheet

if not all([col_po, col_horse, col_trailer]):
    print("Error: Could not find required columns ('client po', 'horse reg no', or 'trailer reg no')")
    exit()

for row in range(2, ws.max_row + 1):
    raw_val = ws.cell(row=row, column=col_po).value
    if raw_val is None:
        continue
        
    po_val = str(raw_val).strip().upper()

    if po_val in updates:
        ws.cell(row=row, column=col_horse).value = updates[po_val]['horse reg']
        ws.cell(row=row, column=col_trailer).value = updates[po_val]['trailer reg']
        print(f"MATCH FOUND: Updated PO {po_val} at row {row}")
        update_count += 1

# 4. Save the file
output_path = r'C:\Users\Jason\Projects\TRACKING-FILE-UPDATE\BARTRAC_FINAL_UPDATED.xlsx'
wb.save(output_path)

print(f"\nUpdate Cycle Finished. Total Rows Updated: {update_count}")
```

### Key Fixes Applied:
* **Column Names:** Changed the lookup from `horse reg` to `horse reg no` and `trailer reg no` based on the column list you provided.
* **Whitespace Sanitization:** Your output showed `final documents   submitted` with multiple spaces. I added `" ".join(str(cell.value).lower().split())` which collapses any amount of whitespace into a single space, making the header matching much more reliable.
* **Safety:** It now checks if all three required columns are found before starting the loop to prevent the `KeyError` you just saw.

