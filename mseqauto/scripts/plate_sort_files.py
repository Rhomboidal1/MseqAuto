# plate_sort_files.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
import re

# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

def get_folder_from_user():
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Add this line - force an update
    
    folder_path = filedialog.askdirectory(
        title="Select today's data folder to sort",
        mustexist=True
    )
    
    root.destroy()
    return folder_path

def main():
    # Get folder path FIRST before any package imports
    data_folder = get_folder_from_user()

    if not data_folder:
        print("No folder selected, exiting")
        return

    # NOW import package modules
    from mseqauto.utils import setup_logger
    from mseqauto.config import MseqConfig
    from mseqauto.core import OSCompatibilityManager, FileSystemDAO, MseqAutomation, FolderProcessor

    # Setup logger
    logger = setup_logger("ind_auto_mseq")
    logger.info("Starting IND auto mSeq...")

    # Log that we already selected folder
    logger.info(f"Using folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)

    # Check for 32-bit Python requirement
    OSCompatibilityManager.py32_check(
        script_path=__file__,
        logger=logger
    )

    # Log OS environment information
    OSCompatibilityManager.log_environment_info(logger)

    # Initialize components
    logger.info("Initializing components...")
    config = MseqConfig()
    logger.info("Config loaded")

    # Use OS compatibility manager for timeouts
    logger.info("Using OS-specific timeouts")

    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")

    ui_automation = MseqAutomation(config, use_fast_navigation=True)
    logger.info("UI Automation initialized with fast navigation")

    processor = FolderProcessor(file_dao, ui_automation, config, logger=logger.info)
    logger.info("Folder processor initialized")

    
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