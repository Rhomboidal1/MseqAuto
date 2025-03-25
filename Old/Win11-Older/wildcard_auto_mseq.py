# wildcard_auto_mseq.py
import os
import tkinter as tk
import tkinter.filedialog as filedialog
from config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import re

# Modified wildcard_auto_mseq.py with improved dialog handling
import os
import tkinter as tk
from tkinter import filedialog
from config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import re
import threading

def main():
    print("Starting wildcard auto mSeq...")
    
    # Initialize components first
    config = MseqConfig()
    print("Config loaded")
    file_dao = FileSystemDAO(config)
    print("FileSystemDAO initialized")
    ui_automation = MseqAutomation(config)
    print("UI Automation initialized")
    processor = FolderProcessor(file_dao, ui_automation, config)
    print("Folder processor initialized")
    
    # Select folder using a separate function to isolate any issues
    data_folder = get_folder_from_user()
    
    if not data_folder:
        print("No folder selected or dialog failed, using fallback path")
        # Fallback path
        data_folder = r"C:\Users\rhomb\Documents"
    
    print(f"Selected folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)
    
    # Get all folders without filtering
    print("Getting all folders...")
    all_folders = file_dao.get_folders(data_folder)
    print(f"Found {len(all_folders)} folders")
    
    if len(all_folders) == 0:
        print("No folders found to process, exiting")
        return
    
    # List folders that will be processed
    print("Folders to process:")
    for i, folder in enumerate(all_folders):
        print(f"{i+1}. {folder}")
    
    # Ask user to confirm before proceeding
    confirmation = input("Proceed with processing these folders? (y/n): ")
    if confirmation.lower() != 'y':
        print("Cancelled by user")
        return
    
    # Process each folder
    for i, folder in enumerate(all_folders):
        print(f"Processing folder {i+1}/{len(all_folders)}: {folder}")
        processor.process_wildcard_folder(folder)
    
    # Close mSeq application
    print("Closing mSeq application...")
    ui_automation.close()
    print("All done!")

def get_folder_from_user():
    """Get folder selection from user"""
    print("Preparing folder selection dialog...")
    
    try:
        root = tk.Tk()
        root.title("Select Folder")
        
        # Create a variable to store the result
        result = {"folder": None}
        
        # Create a minimal UI
        label = tk.Label(root, text="Select a folder to process:")
        label.pack(pady=10)
        
        result_label = tk.Label(root, text="No folder selected yet")
        result_label.pack(pady=10)
        
        def on_select_folder():
        # Open the folder selection dialog
            folder = filedialog.askdirectory(title="Select a folder for testing")
        
            if folder:
                # Update the result label with the selected path
                result_label.config(text=f"Selected: {folder}")
                print(f"Selected folder: {folder}")
            else:
                # User cancelled
                result_label.config(text="Selection cancelled")
                print("Selection cancelled")
        
        # Create a button to open the dialog
        select_button = tk.Button(root, text="Select Folder", command=on_select_folder)
        select_button.pack(pady=10)
        
        # Create a cancel button
        cancel_button = tk.Button(root, text="Cancel", command=root.destroy)
        cancel_button.pack(pady=10)
        
        # Run the Tkinter event loop
        root.mainloop()
        
        return result["folder"]
    
    except Exception as e:
        print(f"Error creating folder dialog: {e}")
        return None

if __name__ == "__main__":
    main()