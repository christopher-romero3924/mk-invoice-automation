# mk-invoice-automation
 

mk-invoice-automation
Automated invoice processing system for Mt. Kare. Watches a folder for completed invoices, parses key fields, writes data to a master Excel workbook, and archives the processed file — all without manual data entry.

Overview
When a booking invoice is finalized (marked complete in cell E36), the watcher script detects it, extracts the reservation details, appends a row to the Raw_Invoice_Data table in Priority_use_2026_CLEAN_3.xlsx, updates the correct year-specific sheet, and moves the invoice to the Complete Invoices archive folder.

How It Works
Working Invoices/
  └── MK_Invoice_XXXX.xlsx   ← Drop invoice here
          ↓
  mk_invoice_watcher.py detects E36 = "YES"
          ↓
  Parses fields from invoice
          ↓
  Writes row → Raw_Invoice_Data table + year sheet (24/25/26)
          ↓
Complete Invoices/
  └── MK_Invoice_XXXX.xlsx   ← File archived here
File Structure
MK Automation Test/
├── mk_invoice_watcher.py        # Main watcher script
├── Priority_use_2026_CLEAN_3.xlsx  # Master workbook (must be closed during processing)
├── Working Invoices/            # Drop completed invoices here
└── Complete Invoices/           # Processed invoices are moved here
Invoice Field Mapping
The script reads the following cells from each invoice file:

Cell	Field
B6	Group Name
E6	Invoice Number
B9	Contact
B12	Date Start
C12	Date End
B16	Campers
B17	Priority
B18	Days
B19	Explanation
B27	Phone
E36	Completion trigger (YES = process)
Master Workbook Structure
Priority_use_2026_CLEAN_3.xlsx

Sheet	Purpose
Raw_Invoice_Data	Master data table (InvoiceData, columns A:L) — all processed invoices append here
26	2026 year sheet — 65 data rows (rows 6–70), column G uses SUMIFS by group/priority/year
25	2025 year sheet
24	2024 year sheet
⚠️ The workbook must be closed before the watcher processes a file. Excel file locks will cause write failures.

Setup & Installation
Requirements
Python 3.x
openpyxl library
bash
pip install openpyxl
Configuration
Edit the paths at the top of mk_invoice_watcher.py if your folder locations differ:

python
BASE_DIR = r"C:\Users\Christopher\Desktop\MK Automation Test"
WORKING_DIR = os.path.join(BASE_DIR, "Working Invoices")
COMPLETE_DIR = os.path.join(BASE_DIR, "Complete Invoices")
MASTER_FILE = os.path.join(BASE_DIR, "Priority_use_2026_CLEAN_3.xlsx")
Running the Script
Manual run:

bash
python mk_invoice_watcher.py
Automatic startup via Task Scheduler:

The script is configured to launch automatically on login via Windows Task Scheduler. It runs silently in the background, polling the Working Invoices folder for new files.

To reconfigure Task Scheduler manually:

Open Task Scheduler → Create Basic Task
Trigger: At log on
Action: Start a program → python.exe, argument: full path to mk_invoice_watcher.py
Usage
Complete the invoice in Excel and confirm cell E36 displays YES
Save and close the invoice file
Move (or save) the file into the Working Invoices folder
The watcher detects the file, processes it automatically, and moves it to Complete Invoices
Open Priority_use_2026_CLEAN_3.xlsx to verify the new row in Raw_Invoice_Data
Troubleshooting
Problem	Likely Cause	Fix
File not processed	E36 ≠ YES	Confirm the red completion bar is marked in the invoice
Write error	Master workbook is open	Close Priority_use_2026_CLEAN_3.xlsx before dropping files
File stays in Working Invoices	Script not running	Check Task Scheduler or run the script manually
Wrong year sheet	Invoice date format issue	Verify Date Start cell (B12) is formatted as a date
Invoice Template
Use MK_Invoice_Template.xlsx for all new bookings. The template matches the cell layout expected by the parser. Do not move or rename key cells.

Author
Christopher Romero
github.com/christopher-romero3924


