import os
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import List
from constants import RECENT_FILES_MAX

def browse_file(entry_widget: tk.Entry, recent_files: list) -> None:
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)
        update_recent_files(file_path, recent_files)

def update_recent_files(file_path: str, recent_files: list) -> None:
    if file_path in recent_files:
        recent_files.remove(file_path)
    recent_files.insert(0, file_path)
    if len(recent_files) > RECENT_FILES_MAX:
        recent_files.pop()

def load_recent_file(file_path: str, entry_file1: tk.Entry, entry_file2: tk.Entry, root: tk.Tk) -> None:
    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"File no longer exists:\n{file_path}", parent=root)
        return
    
    focused_widget = root.focus_get()
    if focused_widget in [entry_file1, entry_file2]:
        focused_widget.delete(0, tk.END)
        focused_widget.insert(0, file_path)
    else:
        choice = messagebox.askquestion(
            "Load File",
            f"Load '{os.path.basename(file_path)}' into which field?",
            detail="Yes for Source, No for Target",
            icon='question',
            parent=root
        )
        target = entry_file1 if choice == 'yes' else entry_file2
        target.delete(0, tk.END)
        target.insert(0, file_path)