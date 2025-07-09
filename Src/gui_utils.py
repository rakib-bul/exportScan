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

2. DATA CLEANING:
   - Remove empty rows above the header
   - Ensure consistent formatting (text vs numbers)
   - Check for merged cells - unmerge them
   - Remove special characters from column names

3. FORMATTING TIPS:
   - Save files as .xlsx format
   - Use one row per record
   - Keep consistent data types in columns
   - Remove unnecessary sheets

4. MATCHING PRIORITY:
   - First: Match by PO Number only
   - Second: Match by Job No (last 4 digits) + PO Number
   - Third: Match by Style Reference + Color

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
Version: Alpha 0.2.0
Last Updated: 2025-07-09

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