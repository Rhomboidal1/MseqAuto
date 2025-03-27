# plate_sort_files.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
import re

from mseqauto.core import FileSystemDAO, FolderProcessor
from mseqauto.utils import setup_logger
from mseqauto.config import MseqConfig


# Check for 32-bit Python requirement
if sys.maxsize > 2**32:
    py32_path = MseqConfig.PYTHON32_PATH
    if os.path.exists(py32_path) and py32_path != sys.executable:
        script_path = os.path.abspath(__file__)
        subprocess.run([py32_path, script_path])
        sys.exit(0)
    else:
        print("32-bit Python not specified or same as current interpreter")
        print("Continuing with current Python interpreter")

def get_folder_from_user():
    """Get folder selection from user"""
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(
        title="Select folder containing plate folders to sort",
        mustexist=True
    )
    root.destroy()
    
    if folder_path:
        print(f"Selected folder: {folder_path}")
        return folder_path
    else:
        print("No folder selected")
        return None

def main():
    # Setup logger
    logger = setup_logger("plate_sort_files")
    logger.info("Starting plate sort files...")
    
    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")
    processor = FolderProcessor(file_dao, None, config, logger=logger.info)
    logger.info("Folder processor initialized")
    
    # Select folder
    data_folder = get_folder_from_user()
    
    if not data_folder:
        logger.error("No folder selected, exiting")
        print("No folder selected, exiting")
        return
    
    logger.info(f"Using folder: {data_folder}")
    
    # Get plate folders
    plate_folders = file_dao.get_folders(data_folder, pattern=r'p\d+.+')
    logger.info(f"Found {len(plate_folders)} plate folders")
    
    if not plate_folders:
        logger.warning("No plate folders found, exiting")
        print("No plate folders found, exiting")
        return
    
    # Process each plate folder
    for i, folder in enumerate(plate_folders):
        logger.info(f"Processing plate folder {i+1}/{len(plate_folders)}: {os.path.basename(folder)}")
        
        # Sort controls
        logger.info(f"Sorting controls in {os.path.basename(folder)}")
        processor.sort_controls(folder)
        
        # Sort blanks
        logger.info(f"Sorting blanks in {os.path.basename(folder)}")
        processor.sort_blanks(folder)
        
        # Remove braces from filenames
        logger.info(f"Removing braces from filenames in {os.path.basename(folder)}")
        processor.remove_braces_from_filenames(folder)
    
    logger.info("All plate folders processed")
    print("All done!")

if __name__ == "__main__":
    main()