Architected reusable Python script for bulk updating shipment records by client PO
Architected reusable Python script for bulk updating shipment records by client PO
Good — I have the exact layout. Header is row 1, CLIENT PO = col B, HORSE REG NO = col K, TRAILER REG NO = col L on the CURRENT SHIPMENTS  sheet.


Created a file, ran a command
Created a file, ran a command
All 5 POs found and updated cleanly. Let me verify one cell actually wrote correctly:


Ran a command, read a file
Ran a command, read a file
Script confirmed working. Here's what it does and how to use it:

Findings from your sheet:

Sheet: CURRENT SHIPMENTS  (trailing space — already handled)
CLIENT PO → col B, HORSE REG NO → col K, TRAILER REG NO → col L
Header is row 1, data starts row 2
How to use it:

Copy the script into the same folder as your workbook
Edit the UPDATES dict at the top — replace the dummy reg numbers with your real ones
Set OUTPUT_FILE to the same name as INPUT_FILE if you want to overwrite in-place
Run: python bartrac_update_vehicle_regs.py
Features:

Skips a field if you set it to None (e.g. only update horse reg, leave trailer alone)
Reports exactly which rows were updated
Lists any POs it couldn't find, so nothing silently fails
Works for any number of POs — just add more entries to UPDATES

Bartrac update vehicle regs
PY 
