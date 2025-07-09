import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import os
from datetime import datetime
import logging
from typing import Tuple, Optional, Dict, Any

# Constants
REQUIRED_COLS_DF1 = ['jobno', 'ponumber', 'exfactoryqty', 'stylerefno', 'color']
REQUIRED_COLS_DF2 = ['jobno', 'ponumber', 'shipqty', 'stylerefno', 'color']
STATUS_COLUMN = 'Status'
RECENT_FILES_MAX = 5

# Configure logging
logging.basicConfig(
    filename='export_checker.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def clean_column_name(col: str) -> str:
    """Clean and standardize column names.
    
    Args:
        col: Column name to clean
        
    Returns:
        Cleaned column name with spaces, underscores, and hyphens removed in lowercase
    """
    return str(col).strip().lower().replace(' ', '').replace('_', '').replace('-', '')

def get_excel_engine(file_path: str) -> Optional[str]:
    """Determine appropriate Excel engine based on file extension.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Engine name ('openpyxl' for .xlsx, 'xlrd' for .xls) or None for auto-detection
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.xlsx':
        return 'openpyxl'
    elif ext == '.xls':
        return 'xlrd'
    return None

def validate_file(file_path: str) -> Tuple[bool, str]:
    """Validate that a file exists and is readable.
    
    Args:
        file_path: Path to the file to validate
        
    Returns:
        Tuple of (is_valid, error_message)
    """
    if not os.path.exists(file_path):
        return False, f"File does not exist: {file_path}"
    if not os.access(file_path, os.R_OK):
        return False, f"File is not readable: {file_path}"
    return True, ""

def compare_excel_files(file1_path: str, file2_path: str, status_var: tk.StringVar, 
                       result_text: scrolledtext.ScrolledText) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """Compare two Excel files for export quantity matching.
    
    Args:
        file1_path: Path to source Excel file (from Logic)
        file2_path: Path to target Excel file (from Store)
        status_var: Tkinter StringVar for status updates
        result_text: ScrolledText widget for progress output
        
    Returns:
        Tuple of (result_dataframe, error_message)
    """
    try:
        # Initialize logging
        logging.info(f"Starting comparison between {file1_path} and {file2_path}")
        
        # Validate files
        for path in [file1_path, file2_path]:
            is_valid, error = validate_file(path)
            if not is_valid:
                return None, error

        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, "Starting comparison...\n")
        result_text.update()
        
        # Get appropriate engines
        engine1 = get_excel_engine(file1_path)
        engine2 = get_excel_engine(file2_path)
        
        # Read Excel files
        result_text.insert(tk.END, "Reading files...\n")
        result_text.update()
        df1 = pd.read_excel(file1_path, engine=engine1)
        df2 = pd.read_excel(file2_path, engine=engine2)
        
        # Clean column names
        result_text.insert(tk.END, "Standardizing column names...\n")
        result_text.update()
        df1.columns = [clean_column_name(col) for col in df1.columns]
        df2.columns = [clean_column_name(col) for col in df2.columns]
        
        # Check required columns
        required_cols = {
            'df1': REQUIRED_COLS_DF1,
            'df2': REQUIRED_COLS_DF2
        }
        
        for df_name, cols in required_cols.items():
            df = df1 if df_name == 'df1' else df2
            missing = [col for col in cols if col not in df.columns]
            if missing:
                error_msg = f"Missing columns in {df_name}: {', '.join(missing)}"
                logging.error(error_msg)
                return None, error_msg
        
        # Initialize status column and counters
        df2[STATUS_COLUMN] = "Not Checked"
        match_counts = {
            'po_match': 0,
            'job_po_match': 0,
            'style_color_match': 0,
            'no_match': 0,
            'qty_mismatch': 0,
            'less_shipment': 0,
            'over_shipment': 0,
            'no_shipment': 0
        }
        
        # Create last4 columns for jobno
        df1['jobno_last4'] = df1['jobno'].astype(str).str.strip().str[-4:]
        df2['jobno_last4'] = df2['jobno'].astype(str).str.strip().str[-4:]
        
        result_text.insert(tk.END, "\n=== Matching Process Started ===\n")
        result_text.update()

        # FIRST PRIORITY MATCHING: PO Number only
        result_text.insert(tk.END, "\n1. Matching by PO Number...\n")
        result_text.update()
        po_mask = (
            df2['ponumber'].notna() &
            df2['ponumber'].isin(df1['ponumber'])
        )
        
        for idx in df2[po_mask].index:
            if df2.at[idx, STATUS_COLUMN] != "Not Checked":
                continue
                
            po_val = df2.at[idx, 'ponumber']
            match = df1[df1['ponumber'] == po_val]
            
            if not match.empty:
                exfactory_qty = match.iloc[0]['exfactoryqty']
                ship_qty = df2.at[idx, 'shipqty']
                
                # Check if export quantity is empty
                if pd.isna(exfactory_qty) or exfactory_qty == '':
                    df2.at[idx, STATUS_COLUMN] = 'No Shipment (PO Match)'
                    match_counts['no_shipment'] += 1
                elif exfactory_qty == ship_qty:
                    df2.at[idx, STATUS_COLUMN] = 'Ok (PO Match)'
                    match_counts['po_match'] += 1
                elif exfactory_qty < ship_qty:
                    df2.at[idx, STATUS_COLUMN] = f'Over Shipment (PO Match: {exfactory_qty} vs {ship_qty})'
                    match_counts['over_shipment'] += 1
                else:
                    df2.at[idx, STATUS_COLUMN] = f'Less Shipment (PO Match: {exfactory_qty} vs {ship_qty})'
                    match_counts['less_shipment'] += 1
        
        # SECOND PRIORITY MATCHING: Job No (last4) + PO Number
        result_text.insert(tk.END, "2. Matching by Job No (last4) + PO Number...\n")
        result_text.update()
        job_po_mask = (
            (df2[STATUS_COLUMN] == "Not Checked") &
            df2['jobno_last4'].notna() & 
            df2['ponumber'].notna() &
            df2['jobno_last4'].isin(df1['jobno_last4']) &
            df2['ponumber'].isin(df1['ponumber'])
        )
        
        for idx in df2[job_po_mask].index:
            job_val = df2.at[idx, 'jobno_last4']
            po_val = df2.at[idx, 'ponumber']
            match = df1[(df1['jobno_last4'] == job_val) & (df1['ponumber'] == po_val)]
            
            if not match.empty:
                exfactory_qty = match.iloc[0]['exfactoryqty']
                ship_qty = df2.at[idx, 'shipqty']
                
                # Check if export quantity is empty
                if pd.isna(exfactory_qty) or exfactory_qty == '':
                    df2.at[idx, STATUS_COLUMN] = 'No Shipment (Job+PO Match)'
                    match_counts['no_shipment'] += 1
                elif exfactory_qty == ship_qty:
                    df2.at[idx, STATUS_COLUMN] = 'Ok (Job+PO Match)'
                    match_counts['job_po_match'] += 1
                elif exfactory_qty < ship_qty:
                    df2.at[idx, STATUS_COLUMN] = f'Over Shipment (Job+PO: {exfactory_qty} vs {ship_qty})'
                    match_counts['over_shipment'] += 1
                else:
                    df2.at[idx, STATUS_COLUMN] = f'Less Shipment (Job+PO: {exfactory_qty} vs {ship_qty})'
                    match_counts['less_shipment'] += 1
        
        # THIRD PRIORITY MATCHING: Style Ref + Color
        result_text.insert(tk.END, "3. Matching by Style Ref + Color...\n")
        result_text.update()
        style_color_mask = (df2[STATUS_COLUMN] == "Not Checked")
        secondary_df = df2[style_color_mask].copy()
        
        for idx in secondary_df.index:
            style_val = secondary_df.at[idx, 'stylerefno']
            color_val = secondary_df.at[idx, 'color']
            
            match = df1[
                (df1['stylerefno'] == style_val) &
                (df1['color'] == color_val)
            ]
            
            if not match.empty:
                exfactory_qty = match.iloc[0]['exfactoryqty']
                ship_qty = secondary_df.at[idx, 'shipqty']
                
                # Check if export quantity is empty
                if pd.isna(exfactory_qty) or exfactory_qty == '':
                    df2.at[idx, STATUS_COLUMN] = 'No Shipment (Style+Color Match)'
                    match_counts['no_shipment'] += 1
                elif exfactory_qty == ship_qty:
                    df2.at[idx, STATUS_COLUMN] = 'Ok (Style+Color Match)'
                    match_counts['style_color_match'] += 1
                elif exfactory_qty < ship_qty:
                    df2.at[idx, STATUS_COLUMN] = f'Over Shipment (Style+Color: {exfactory_qty} vs {ship_qty})'
                    match_counts['over_shipment'] += 1
                else:
                    df2.at[idx, STATUS_COLUMN] = f'Less Shipment (Style+Color: {exfactory_qty} vs {ship_qty})'
                    match_counts['less_shipment'] += 1
            else:
                df2.at[idx, STATUS_COLUMN] = 'No Match Found'
                match_counts['no_match'] += 1
        
        # Drop temporary columns before returning
        df2 = df2.drop(columns=['jobno_last4'])
        
        # Generate summary report
        summary = (
            f"\n=== Matching Summary ===\n"
            f"Perfect Matches: {match_counts['po_match'] + match_counts['job_po_match'] + match_counts['style_color_match']}\n"
            f"Less Shipment Cases: {match_counts['less_shipment']}\n"
            f"Over Shipment Cases: {match_counts['over_shipment']}\n"
            f"No Shipment Cases: {match_counts['no_shipment']}\n"
            f"No Matches Found: {match_counts['no_match']}\n"
            f"Total Records Processed: {len(df2)}\n"
        )
        result_text.insert(tk.END, summary)
        result_text.update()
        
        logging.info("Comparison completed successfully")
        return df2, None
    
    except ImportError as e:
        error_msg = "Reading .xls files requires xlrd. Install with: pip install xlrd" if 'xlrd' in str(e).lower() else f"Import Error: {str(e)}"
        logging.error(error_msg)
        return None, error_msg
    except Exception as e:
        error_msg = f"Error during processing: {str(e)}"
        logging.error(error_msg, exc_info=True)
        return None, error_msg

# [Rest of the code remains exactly the same as in your original file]
# All functions from browse_file() to main() should be kept exactly as they were
# Only the compare_excel_files() function has been modified as shown above

def browse_file(entry_widget: ttk.Entry, recent_files: list) -> None:
    """Open file dialog and update entry widget.
    
    Args:
        entry_widget: Entry widget to update with selected file path
        recent_files: List to track recently used files
    """
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)
        update_recent_files(file_path, recent_files)

def update_recent_files(file_path: str, recent_files: list) -> None:
    """Update the list of recently used files.
    
    Args:
        file_path: New file path to add
        recent_files: List of recent files to update
    """
    if file_path in recent_files:
        recent_files.remove(file_path)
    recent_files.insert(0, file_path)
    if len(recent_files) > RECENT_FILES_MAX:
        recent_files.pop()

def show_summary_stats(df: pd.DataFrame) -> Dict[str, Any]:
    """Generate summary statistics from comparison results.
    
    Args:
        df: DataFrame with comparison results
        
    Returns:
        Dictionary with various statistics
    """
    return {
        'Total Records': len(df),
        'Perfect Matches': len(df[df[STATUS_COLUMN].str.startswith('Ok')]),
        'Quantity Mismatches': len(df[df[STATUS_COLUMN].str.contains('Qty Mismatch')]),
        'No Matches': len(df[df[STATUS_COLUMN] == 'No Match Found'])
    }

def show_guide() -> None:
    """Display the Excel file preparation guide in a new window."""
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

def show_developer_info() -> None:
    """Display developer information in a new window."""
    info_window = tk.Toplevel()
    info_window.title("Developer Information")
    info_window.geometry("400x200")
    
    info_text = """=== Developer Information ===

Application: ExportScan
Version: 1.0
Last Updated: 2025-07-10

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

Note: This is a personal project and is not affiliated with Southern Garments Ltd. or Reneissance Group LTD.
"""
    text_widget = scrolledtext.ScrolledText(info_window, wrap=tk.WORD, width=50, height=12)
    text_widget.insert(tk.INSERT, info_text)
    text_widget.configure(state='disabled')
    text_widget.pack(padx=10, pady=10)

def execute_comparison(recent_files: list) -> None:
    """Handle comparison process with progress indication.
    
    Args:
        recent_files: List to track recently used files
    """
    file1 = entry_file1.get()
    file2 = entry_file2.get()
    
    if not file1 or not file2:
        messagebox.showerror("Error", "Please select both Excel files")
        return
    
    # Validate files before processing
    for path, name in [(file1, "Source File"), (file2, "Target File")]:
        is_valid, error = validate_file(path)
        if not is_valid:
            messagebox.showerror("Error", f"{name}: {error}")
            return
    
    status_var.set("Processing...")
    progress_bar.start(10)
    result_text.delete(1.0, tk.END)
    root.update_idletasks()
    
    try:
        result, error = compare_excel_files(file1, file2, status_var, result_text)
        
        if error:
            messagebox.showerror("Error", error)
            result_text.insert(tk.END, f"\nERROR: {error}\n")
        else:
            # Add to recent files
            update_recent_files(file1, recent_files)
            update_recent_files(file2, recent_files)
            
            # Show summary statistics
            stats = show_summary_stats(result)
            result_text.insert(tk.END, "\n=== Detailed Statistics ===\n")
            for stat, value in stats.items():
                result_text.insert(tk.END, f"{stat}: {value}\n")
            
            # Save results
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
            )
            if save_path:
                if save_path.endswith('.csv'):
                    result.to_csv(save_path, index=False)
                else:
                    result.to_excel(save_path, index=False)
                
                message = f"\nFile saved successfully at:\n{save_path}\n"
                messagebox.showinfo("Success", message)
                result_text.insert(tk.END, message)
    except Exception as e:
        messagebox.showerror("Error", str(e))
        result_text.insert(tk.END, f"\nERROR: {str(e)}\n")
    finally:
        progress_bar.stop()
        status_var.set("Ready")

def create_recent_files_menu(menubar: tk.Menu, recent_files: list) -> None:
    """Create a recent files submenu.
    
    Args:
        menubar: Main menu bar
        recent_files: List of recent files
    """
    recent_menu = tk.Menu(menubar, tearoff=0)
    
    if not recent_files:
        recent_menu.add_command(label="No recent files", state='disabled')
    else:
        for file_path in recent_files:
            recent_menu.add_command(
                label=os.path.basename(file_path),
                command=lambda f=file_path: load_recent_file(f)
            )
    
    menubar.add_cascade(label="Recent Files", menu=recent_menu)

def load_recent_file(file_path: str) -> None:
    """Load a file from the recent files list.
    
    Args:
        file_path: Path to the file to load
    """
    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"File no longer exists:\n{file_path}")
        return
    
    # Determine which entry widget to update based on current focus
    focused_widget = root.focus_get()
    if focused_widget in [entry_file1, entry_file2]:
        focused_widget.delete(0, tk.END)
        focused_widget.insert(0, file_path)
    else:
        # If no specific widget is focused, ask which one to update
        choice = messagebox.askquestion(
            "Load File",
            f"Load '{os.path.basename(file_path)}' into which field?",
            detail="Yes for Source, No for Target",
            icon='question'
        )
        target = entry_file1 if choice == 'yes' else entry_file2
        target.delete(0, tk.END)
        target.insert(0, file_path)

# Main application
def main():
    global root, entry_file1, entry_file2, status_var, progress_bar, result_text
    
    root = tk.Tk()
    root.title("ExportScan - v1.0")
    root.geometry("900x700")
    
    # Initialize recent files list
    recent_files = []
    
    # Menu Bar
    menubar = tk.Menu(root)
    root.config(menu=menubar)
    
    # File menu
    file_menu = tk.Menu(menubar, tearoff=0)
    file_menu.add_command(label="Exit", command=root.quit)
    menubar.add_cascade(label="File", menu=file_menu)
    
    # Recent files menu
    create_recent_files_menu(menubar, recent_files)
    
    # Help menu
    help_menu = tk.Menu(menubar, tearoff=0)
    help_menu.add_command(label="Excel Preparation Guide", command=show_guide)
    help_menu.add_command(label="Developer Information", command=show_developer_info)
    menubar.add_cascade(label="Help", menu=help_menu)
    
    # File Selection Frame
    frame_files = ttk.LabelFrame(root, text="File Selection")
    frame_files.pack(padx=10, pady=10, fill='x')
    
    # File 1
    ttk.Label(frame_files, text="Source File (Excel from Logic):").grid(row=0, column=0, padx=5, pady=5, sticky='w')
    entry_file1 = ttk.Entry(frame_files, width=60)
    entry_file1.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
    ttk.Button(frame_files, text="Browse", command=lambda: browse_file(entry_file1, recent_files)).grid(row=0, column=2, padx=5, pady=5)
    
    # File 2
    ttk.Label(frame_files, text="Target File (Export Report from Store):").grid(row=1, column=0, padx=5, pady=5, sticky='w')
    entry_file2 = ttk.Entry(frame_files, width=60)
    entry_file2.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
    ttk.Button(frame_files, text="Browse", command=lambda: browse_file(entry_file2, recent_files)).grid(row=1, column=2, padx=5, pady=5)
    
    # Configure grid weights
    frame_files.columnconfigure(1, weight=1)
    
    # Progress/Status
    status_frame = ttk.Frame(root)
    status_frame.pack(fill='x', padx=10, pady=5)
    status_var = tk.StringVar(value="Ready")
    ttk.Label(status_frame, textvariable=status_var).pack(side='left')
    progress_bar = ttk.Progressbar(status_frame, mode='indeterminate', length=300)
    progress_bar.pack(side='right', padx=10)
    
    # Action Buttons
    button_frame = ttk.Frame(root)
    button_frame.pack(pady=10)
    ttk.Button(button_frame, text="Compare and Save", command=lambda: execute_comparison(recent_files)).pack(side='left', padx=5)
    ttk.Button(button_frame, text="Clear Fields", command=lambda: [entry_file1.delete(0, tk.END), entry_file2.delete(0, tk.END)]).pack(side='left', padx=5)
    
    # Results Display
    result_frame = ttk.LabelFrame(root, text="Comparison Results")
    result_frame.pack(padx=10, pady=10, fill='both', expand=True)
    
    result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, width=120, height=25)
    result_text.pack(padx=5, pady=5, fill='both', expand=True)
    
    root.mainloop()

if __name__ == "__main__":
    main()