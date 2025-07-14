import tkinter as tk
from tkinter import scrolledtext

def show_guide():
    guide_window = tk.Toplevel()
    guide_window.title("Excel File Preparation Guide")
    guide_window.geometry("700x500")
    
    guide_text = """=== Excel File Preparation Guide ===

1. REQUIRED COLUMNS:
   - Both files must contain these columns (names can vary):
     * JobNo / Job Number
     * PONumber / PO Number
     * StyleRefNo / Style Reference
     * Color
     * ExFactoryQty (Source file)
     * ShipQty (Target file)
   - For buyer-specific matching (NEXT/Vogue), target file must have 'Buyer' column

2. BUYER-SPECIFIC FEATURES (NEXT/Vogue):
   - Enable "Buyer-Specific Matching" checkbox
   - Select which file to create combined PO (StyleRefNo + PO)
   - Matching sequence:
     1. PO Number + Job Number (last 4 digits)
     2. Combined PO (StyleRefNo + PO) 
     3. Job No wise aggregation (if previous matches fail)

3. DATA CLEANING:
   - Remove empty rows above the header
   - Ensure consistent formatting (text vs numbers)
   - Check for merged cells - unmerge them
   - Remove special characters from column names

4. MATCHING PRIORITY:
   - For standard buyers:
     * First: Match by PO Number only
     * Second: Match by Job No (last 4 digits) + PO Number
     * Third: Match by Style Reference + Color
   - For NEXT/Vogue buyers:
     * First: Match by PO Number + Job Number
     * Second: Match by Combined PO (StyleRefNo + PO Number)
     * Third: Match by Job No wise aggregation

5. TROUBLESHOOTING:
   - If matching fails, check for:
     * Leading/trailing spaces in key fields
     * Different data types (text vs numbers)
     * Hidden characters in fields
     * Case sensitivity in text fields
"""

    text_widget = scrolledtext.ScrolledText(guide_window, wrap=tk.WORD, width=80, height=30)
    text_widget.insert(tk.INSERT, guide_text)
    text_widget.configure(state='disabled')
    text_widget.pack(padx=10, pady=10)

def show_developer_info():
    info_window = tk.Toplevel()
    info_window.title("Developer Information")
    info_window.geometry("400x200")
    
    info_text = """=== Developer Information ===

Application: Export Scan
Version: Alpha 0.3.1
Last Updated: 2025-07-14

Developer: Rakib Hasan Bulbul
Organization: SGL

Contact:
- Email: hrakib182@gmail.com
- Phone: +8801783924660

Technologies Used:
- Python 3.9
- Pandas
- Tkinter
- OpenPyXL
"""
    text_widget = scrolledtext.ScrolledText(info_window, wrap=tk.WORD, width=50, height=12)
    text_widget.insert(tk.INSERT, info_text)
    text_widget.configure(state='disabled')
    text_widget.pack(padx=10, pady=10)