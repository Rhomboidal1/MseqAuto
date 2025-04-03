# ind_sort_files.py
import os
import subprocess
import tkinter as tk
import sys
from datetime import datetime
from tkinter import filedialog

# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

# print(sys.path)

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

def get_recent_inumbers(file_dao, logger):
    """Get list of recent I numbers to process based on order key"""
    paths = ['G:\\Lab\\Spreadsheets\\Individual Uploaded to ABI', 'G:\\Lab\\Spreadsheets']
    
    # Try to find the most recent I number from the Data/Individuals folder
    logger.info("Finding most recent I number from Data/Individuals")
    lower_inum = file_dao.get_most_recent_inumber('P:\\Data\\Individuals')
    
    # Get recent files (past 5 days)
    logger.info("Getting files from last 5 days")
    recent_files = file_dao.get_recent_files(paths, 5)
    
    # Filter to get only I numbers greater than the most recent one
    logger.info(f"Filtering I numbers greater than {lower_inum}")
    final_inum_files = file_dao.get_inumbers_greater_than(recent_files, lower_inum)
    final_inums = [file_dao.get_inumber_from_name(f) for f in final_inum_files]
    
    # Also include I numbers from files modified in last 12 hours
    logger.info("Including files from last 12 hours")
    recent_12hr_files = file_dao.get_recent_files(paths, hours=12)
    additional_inums = [file_dao.get_inumber_from_name(f) for f in recent_12hr_files]
    
    # Combine all I numbers (removing duplicates)
    for inum in additional_inums:
        if inum and inum not in final_inums:
            final_inums.append(inum)
    
    return [i for i in final_inums if i]  # Filter out None values

def main():
    # Get folder path first before any package imports
    data_folder = get_folder_from_user()
    
    if not data_folder:
        print("No folder selected, exiting")
        return

    # ONLY NOW import package modules
    from mseqauto.config import MseqConfig
    from mseqauto.core import FileSystemDAO, FolderProcessor
    from mseqauto.utils import setup_logger
    
    # Setup logger
    logger = setup_logger("ind_sort_files")

    logger.info("Starting IND sort files...")
    
    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")
    processor = FolderProcessor(file_dao, None, config, logger=logger.info)
    logger.info("Folder processor initialized")
    logger.info(f"Using folder: {data_folder}")

    # Run batch file to generate order key
    try:
        logger.info(f"Running batch file: {config.BATCH_FILE_PATH}")
        subprocess.run(config.BATCH_FILE_PATH, shell=True, check=True)
        logger.info("Batch file completed successfully")
    except subprocess.CalledProcessError:
        logger.error(f"Batch file {config.BATCH_FILE_PATH} failed to run")
        print(f"Error: Batch file {config.BATCH_FILE_PATH} failed to run")
        return
    
    # Store the selected folder in the processor for later reference
    processor.current_data_folder = data_folder

    # Get today's I numbers and BioI folders
    i_numbers, bio_folders = processor.get_todays_inumbers_from_folder(data_folder)
    logger.info(f"Found {len(i_numbers)} I numbers and {len(bio_folders)} BioI folders")
    
    # Get recent I numbers for order key filtering
    recent_inumbers = get_recent_inumbers(file_dao, logger)
    logger.info(f"Using recent I numbers: {recent_inumbers}")
    
    # Load order key and adjust characters
    order_key = file_dao.load_order_key(config.KEY_FILE_PATH)
    logger.info("Order key loaded")
    
    # Get complete list of reinjects
    reinject_path = f"P:\\Data\\Reinjects\\Reinject List_{datetime.now().strftime('%m-%d-%Y')}.xlsx"
    try:
        reinject_list = processor.get_reinject_list(recent_inumbers, reinject_path)
        logger.info(f"Found {len(reinject_list)} reinjects")
    except Exception as e:
        logger.error(f"Error loading reinject list: {e}")
        reinject_list = []
    
    # Process each BioI folder
    for i, folder in enumerate(bio_folders):
        logger.info(f"Processing folder {i+1}/{len(bio_folders)}: {os.path.basename(folder)}")
        processor.sort_ind_folder(folder, reinject_list, order_key)
    
    logger.info("All folders processed")
    print("All done!")

if __name__ == "__main__":
    main()