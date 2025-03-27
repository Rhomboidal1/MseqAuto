# plate_auto_mseq.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
from mseqauto.core import FileSystemDAO, FolderProcessor
from mseqauto.core import MseqAutomation
from mseqauto.config import MseqConfig
import re
import time

# Check for 32-bit Python requirement - gracefully fallback if not available
if sys.maxsize > 2**32:
    # Path to 32-bit Python
    py32_path = MseqConfig.PYTHON32_PATH
    
    # Only relaunch if we have a different 32-bit Python
    if os.path.exists(py32_path) and py32_path != sys.executable:
        # Get the full path of the current script
        script_path = os.path.abspath(__file__)
        
        # Re-run this script with 32-bit Python and exit current process
        subprocess.run([py32_path, script_path])
        sys.exit(0)
    else:
        print("32-bit Python not specified or same as current interpreter")
        print("Continuing with current Python interpreter")

def get_folder_from_user():
    """Get folder selection from user using a simple approach"""
    print("Opening folder selection dialog...")
    
    # Create and immediately withdraw (hide) the root window
    root = tk.Tk()
    root.withdraw()
    
    # Show a directory selection dialog
    folder_path = filedialog.askdirectory(
        title="Select a folder to process plates",
        mustexist=True
    )
    
    # Destroy the root window
    root.destroy()
    
    if folder_path:
        print(f"Selected folder: {folder_path}")
        return folder_path
    else:
        print("No folder selected")
        return None

def main():
    print("Starting plate auto mSeq...")
    
    # Initialize components
    config = MseqConfig()
    print("Config loaded")
    file_dao = FileSystemDAO(config)
    print("FileSystemDAO initialized")
    ui_automation = MseqAutomation(config)
    print("UI Automation initialized")
    processor = FolderProcessor(file_dao, ui_automation, config)
    print("Folder processor initialized")
    
    # Select folder 
    data_folder = get_folder_from_user()
    
    if not data_folder:
        print("No folder selected, exiting")
        return
    
    print(f"Using folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)
    
    # Get all plate folders (starting with 'p')
    plate_folders = file_dao.get_folders(data_folder, r'p\d+.+')
    print(f"Found {len(plate_folders)} plate folders")
    
    if len(plate_folders) == 0:
        print("No plate folders found to process, exiting")
        return
    
    # List folders that will be processed
    print("Plate folders to process:")
    for i, folder in enumerate(plate_folders):
        print(f"{i+1}. {os.path.basename(folder)}")
    
    # Ask user to confirm before proceeding
    confirmation = input("Proceed with processing these plate folders? (y/n): ")
    if confirmation.lower() != 'y':
        print("Cancelled by user")
        return
    
    # Process each plate folder
    for i, folder in enumerate(plate_folders):
        print(f"Processing plate folder {i+1}/{len(plate_folders)}: {os.path.basename(folder)}")
        processor.process_plate_folder(folder)
        
        # Add a short delay between folder processing
        if i < len(plate_folders) - 1:
            time.sleep(1)
    
    # Close mSeq application
    print("Closing mSeq application...")
    ui_automation.close()
    print("All done!")

if __name__ == "__main__":
    main()