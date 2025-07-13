import os
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox, Menu
import logging
from comparison import compare_excel_files, show_summary_stats
from file_handling import browse_file, update_recent_files, load_recent_file
from gui_utils import show_guide, show_developer_info
from constants import BUYER_SPECIFIC_BUYERS

# Configure logging
logging.basicConfig(
    filename='export_checker.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


class ExportCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Export Scan - Alpha Build 0.3.1")
        self.root.geometry("900x720")
        self.recent_files = []
        self.setup_ui()
        
    def setup_ui(self):
        # Menu Bar
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = Menu(menubar, tearoff=0)
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Recent files menu
        self.recent_menu = Menu(menubar, tearoff=0)
        self.update_recent_files_menu()
        menubar.add_cascade(label="Recent Files", menu=self.recent_menu)
        
        # Help menu
        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label="Excel Preparation Guide", command=show_guide)
        help_menu.add_command(label="Developer Information", command=show_developer_info)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        # File Selection Frame
        frame_files = ttk.LabelFrame(self.root, text="File Selection")
        frame_files.pack(padx=10, pady=10, fill='x')
        
        # File 1
        ttk.Label(frame_files, text="Source File (Excel from Logic):").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.entry_file1 = ttk.Entry(frame_files, width=60)
        self.entry_file1.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(frame_files, text="Browse", command=lambda: browse_file(self.entry_file1, self.recent_files)).grid(row=0, column=2, padx=5, pady=5)
        
        # File 2
        ttk.Label(frame_files, text="Target File (Export Report from Store):").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.entry_file2 = ttk.Entry(frame_files, width=60)
        self.entry_file2.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(frame_files, text="Browse", command=lambda: browse_file(self.entry_file2, self.recent_files)).grid(row=1, column=2, padx=5, pady=5)
        
        # Configure grid weights
        frame_files.columnconfigure(1, weight=1)
        
        # Buyer-Specific Options Frame
        options_frame = ttk.LabelFrame(self.root, text="Comparison Options")
        options_frame.pack(padx=10, pady=5, fill='x')
        
        # Buyer-specific matching checkbox
        self.buyer_specific_var = tk.BooleanVar(value=False)
        buyer_specific_cb = ttk.Checkbutton(
            options_frame,
            text=f"Enable Buyer-Specific Matching ({', '.join(BUYER_SPECIFIC_BUYERS)})",
            variable=self.buyer_specific_var
        )
        buyer_specific_cb.pack(anchor='w', padx=5, pady=2)
        
        # Frame for combined PO selection
        combine_frame = ttk.Frame(options_frame)
        combine_frame.pack(fill='x', padx=5, pady=2)
        ttk.Label(combine_frame, text="Combine PO with StyleRefNo for:").pack(side='left')
        
        # Radio buttons for file selection
        self.combine_po_var = tk.StringVar(value="df1")
        ttk.Radiobutton(combine_frame, text="Source File", 
                        variable=self.combine_po_var, value="df1").pack(side='left', padx=5)
        ttk.Radiobutton(combine_frame, text="Target File", 
                        variable=self.combine_po_var, value="df2").pack(side='left', padx=5)
        
        # Progress/Status
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill='x', padx=10, pady=5)
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side='left')
        self.progress_bar = ttk.Progressbar(status_frame, mode='indeterminate', length=300)
        self.progress_bar.pack(side='right', padx=10)
        
        # Action Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Compare and Save", command=self.execute_comparison).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Clear Fields", command=self.clear_fields).pack(side='left', padx=5)
        
        # Results Display
        result_frame = ttk.LabelFrame(self.root, text="Comparison Results")
        result_frame.pack(padx=10, pady=10, fill='both', expand=True)
        
        self.result_text = scrolledtext.ScrolledText(result_frame, wrap=tk.WORD, width=120, height=25)
        self.result_text.pack(padx=5, pady=5, fill='both', expand=True)
    
    def update_recent_files_menu(self):
        self.recent_menu.delete(0, 'end')
        if not self.recent_files:
            self.recent_menu.add_command(label="No recent files", state='disabled')
        else:
            for file_path in self.recent_files:
                self.recent_menu.add_command(
                    label=os.path.basename(file_path),
                    command=lambda f=file_path: load_recent_file(f, self.entry_file1, self.entry_file2, self.root)
                )
    
    def clear_fields(self):
        self.entry_file1.delete(0, tk.END)
        self.entry_file2.delete(0, tk.END)
    
    def execute_comparison(self):
        file1 = self.entry_file1.get()
        file2 = self.entry_file2.get()
        
        if not file1 or not file2:
            messagebox.showerror("Error", "Please select both Excel files")
            return
        
        # Get buyer-specific options
        buyer_specific = self.buyer_specific_var.get()
        combine_po_in = self.combine_po_var.get()
        
        self.status_var.set("Processing...")
        self.progress_bar.start(10)
        self.result_text.delete(1.0, tk.END)
        self.root.update_idletasks()
        
        try:
            result, error = compare_excel_files(
                file1, file2, 
                self.status_var, 
                self.result_text,
                buyer_specific,
                combine_po_in
            )
            
            if error:
                messagebox.showerror("Error", error)
                self.result_text.insert(tk.END, f"\nERROR: {error}\n")
            else:
                # Add to recent files
                update_recent_files(file1, self.recent_files)
                update_recent_files(file2, self.recent_files)
                self.update_recent_files_menu()
                
                # Show summary statistics
                stats = show_summary_stats(result)
                self.result_text.insert(tk.END, "\n=== Detailed Statistics ===\n")
                for stat, value in stats.items():
                    self.result_text.insert(tk.END, f"{stat}: {value}\n")
                
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
                    self.result_text.insert(tk.END, message)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.result_text.insert(tk.END, f"\nERROR: {str(e)}\n")
        finally:
            self.progress_bar.stop()
            self.status_var.set("Ready")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExportCheckerApp(root)
    root.mainloop()