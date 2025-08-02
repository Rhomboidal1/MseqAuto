# wildcard_auto_mseq.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
import re
import time

from mseqauto.core import FileSystemDAO, FolderProcessor
from mseqauto.config import MseqConfig
from mseqauto.core import MseqAutomation

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
        title="Select a folder to process",
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
    print("Starting wildcard auto mSeq...")

    # Initialize components
    config = MseqConfig()
    print("Config loaded")
    file_dao = FileSystemDAO(config)
    print("FileSystemDAO initialized")
    ui_automation = MseqAutomation(config)
    print("UI Automation initialized")
    processor = FolderProcessor(file_dao, ui_automation, config)
    print("Folder processor initialized")

    # Select folder using a simple dialog approach
    data_folder = get_folder_from_user()

    if not data_folder:
        print("No folder selected, using fallback path")
        # Fallback path - use Documents folder
        from pathlib import Path
        data_folder = str(Path.home() / "Documents")

    print(f"Using folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)

    # Get all folders without filtering
    print("Getting all folders...")
    all_folders = file_dao.get_folders(data_folder)
    print(f"Found {len(all_folders)} folders")

    if len(all_folders) == 0:
        print("No folders found to process, exiting")
        return

    # List folders that will be processed (show only first 10 if there are many)
    print("Folders to process:")
    max_display = min(10, len(all_folders))
    for i in range(max_display):
        folder = all_folders[i]
        print(f"{i+1}. {os.path.basename(folder)}")

    if len(all_folders) > max_display:
        print(f"... and {len(all_folders) - max_display} more folders")

    # Ask user to confirm before proceeding
    confirmation = input("Proceed with processing these folders? (y/n): ")
    if confirmation.lower() != 'y':
        print("Cancelled by user")
        return

    # Process each folder
    for i, folder in enumerate(all_folders):
        print(f"Processing folder {i+1}/{len(all_folders)}: {os.path.basename(folder)}")
        processor.process_wildcard_folder(folder)

        # Add a short delay between folder processing
        if i < len(all_folders) - 1:
            time.sleep(1)

    # Close mSeq application
    print("Closing mSeq application...")
    ui_automation.close()
    print("All done!")

if __name__ == "__main__":
    main()