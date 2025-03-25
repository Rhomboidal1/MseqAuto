# ind_auto_mseq.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
from config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import re
import time
from logger import setup_logger

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
        title="Select today's data folder to mseq orders",
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
    # Setup logger
    logger = setup_logger("ind_auto_mseq")
    logger.info("Starting IND auto mSeq...")
    
    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")
    ui_automation = MseqAutomation(config)
    logger.info("UI Automation initialized")
    processor = FolderProcessor(file_dao, ui_automation, config, logger=logger.info)
    logger.info("Folder processor initialized")
    
    # Run batch file to generate order key
    try:
        logger.info(f"Running batch file: {config.BATCH_FILE_PATH}")
        subprocess.run(config.BATCH_FILE_PATH, shell=True, check=True)
        logger.info("Batch file completed successfully")
    except subprocess.CalledProcessError:
        logger.error(f"Batch file {config.BATCH_FILE_PATH} failed to run")
        print(f"Error: Batch file {config.BATCH_FILE_PATH} failed to run")
        return
    
    # Select folder
    data_folder = get_folder_from_user()
    
    if not data_folder:
        logger.error("No folder selected, exiting")
        print("No folder selected, exiting")
        return
    
    logger.info(f"Using folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)
    
    # Process BioI folders
    bio_folders = file_dao.get_folders(data_folder, r'bioi-\d+')
    logger.info(f"Found {len(bio_folders)} BioI folders")
    
    immediate_orders = file_dao.get_folders(data_folder, r'bioi-\d+_.+_\d+')
    logger.info(f"Found {len(immediate_orders)} immediate order folders")
    
    pcr_folders = file_dao.get_folders(data_folder, r'fb-pcr\d+_\d+')
    logger.info(f"Found {len(pcr_folders)} PCR folders")
    
    # Process BioI folders
    for i, folder in enumerate(bio_folders):
        logger.info(f"Processing BioI folder {i+1}/{len(bio_folders)}: {os.path.basename(folder)}")
        processor.process_bio_folder(folder)
    
    # Determine if we're processing the IND Not Ready folder
    is_ind_not_ready = os.path.basename(data_folder) == config.IND_NOT_READY_FOLDER
    logger.info(f"Is IND Not Ready folder: {is_ind_not_ready}")
    
    # Process immediate orders
    for i, folder in enumerate(immediate_orders):
        logger.info(f"Processing order folder {i+1}/{len(immediate_orders)}: {os.path.basename(folder)}")
        processor.process_order_folder(folder, data_folder)
    
    # Process PCR folders
    for i, folder in enumerate(pcr_folders):
        logger.info(f"Processing PCR folder {i+1}/{len(pcr_folders)}: {os.path.basename(folder)}")
        processor.process_pcr_folder(folder)
    
    logger.info("All processing completed")
    print("")
    print("ALL DONE")
    
    # Close mSeq application
    logger.info("Closing mSeq application")
    ui_automation.close()

if __name__ == "__main__":
    main()