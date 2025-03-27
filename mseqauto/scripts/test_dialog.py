# test_dialog.py
import os
import tkinter as tk
from tkinter import filedialog

def get_folder_dialog():
    print("Opening dialog...")
    root = tk.Tk()
    root.withdraw()
    
    # Force update to ensure window is fully initialized
    root.update()
    
    folder_path = filedialog.askdirectory(
        title="Test dialog",
        mustexist=True
    )
    
    root.destroy()
    print(f"Selected: {folder_path}")
    return folder_path

# Try with direct imports first
result = get_folder_dialog()
print(f"Direct import result: {result}")

# Now try with package imports
print("Now importing from package...")
from mseqauto.core import FileSystemDAO, FolderProcessor
print("Package imported")

# Try again after imports
result = get_folder_dialog()
print(f"After package import result: {result}")