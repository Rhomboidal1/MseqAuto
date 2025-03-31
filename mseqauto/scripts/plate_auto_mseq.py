# plate_auto_mseq.py
import tkinter as tk
from tkinter import filedialog
import subprocess
import re
import time
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
                
def get_folder_from_user():
    """Simple folder selection dialog that works reliably"""
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force an update to ensure dialog shows

    folder_path = filedialog.askdirectory(
        title="Select today's data folder to mseq orders",
        mustexist=True
    )

    root.destroy()
    return folder_path

def main():
    print("Starting plate auto mSeq...")
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
    logger = setup_logger("plate_auto_mseq")
    logger.info("Starting Plate auto mSeq...")

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

    # Initialize components
    config = MseqConfig()
    print("Config loaded")
    file_dao = FileSystemDAO(config)
    print("FileSystemDAO initialized")
    ui_automation = MseqAutomation(config)
    print("UI Automation initialized")
    processor = FolderProcessor(file_dao, ui_automation, config)
    print("Folder processor initialized")
    
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