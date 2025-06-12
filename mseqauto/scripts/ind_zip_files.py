# ind_zip_files.py
import subprocess
import re
import tkinter as tk
from tkinter import filedialog
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
import warnings
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

# Check if running under GUI
GUI_MODE = os.environ.get('MSEQAUTO_GUI_MODE', 'False') == 'True'

def get_folder_from_user():
    # Check for GUI-provided folder first
    if GUI_MODE and 'MSEQAUTO_DATA_FOLDER' in os.environ:
        return os.environ['MSEQAUTO_DATA_FOLDER']
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Add this line - force an update
    
    folder_path = filedialog.askdirectory(
        title="Select today's data folder to zip",
        mustexist=True
    )
    
    root.destroy()
    return folder_path

def main():
    # Get folder path first before any package imports
    data_folder = get_folder_from_user()
    
    if not data_folder:
        print("No folder selected, exiting")
        return

    # ONLY NOW import package modules
    from mseqauto.config import MseqConfig # type: ignore
    from mseqauto.core import FileSystemDAO, FolderProcessor # type: ignore
    from mseqauto.utils import setup_logger # type: ignore
    
    # Get the script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "logs")

    # Setup logger
    logger = setup_logger("ind_zip_files", log_dir=log_dir)

    logger.info("Starting IND zip files...")
    
    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    REGEX = config.REGEX_PATTERNS
    file_dao = FileSystemDAO(config, logger=logger)
    logger.info("FileSystemDAO initialized")
    processor = FolderProcessor(file_dao, None, config, logger=logger.info)
    logger.info("Folder processor initialized")
    logger.info(f"Using folder: {data_folder}")
    
    # Create zip dump folder
    zip_dump_folder = os.path.join(data_folder, config.ZIP_DUMP_FOLDER)
    if not os.path.exists(zip_dump_folder):
        os.makedirs(zip_dump_folder)
        logger.info(f"Created zip dump folder: {zip_dump_folder}")

    bio_folders = file_dao.get_folders(data_folder, pattern=REGEX['bioi_folder'].pattern)
    logger.info(f"Found {len(bio_folders)} BioI folders")

    # Get PCR folders - use the new pcr_folder pattern
    pcr_folders = file_dao.get_folders(data_folder, pattern=REGEX['pcr_folder'].pattern)
    logger.info(f"Found {len(pcr_folders)} PCR folders")

    # Copy recent existing zips to the zip dump folder (recovery logic)
    recovered_count = file_dao.copy_recent_zips_to_dump(
        bio_folders + pcr_folders, 
        zip_dump_folder, 
        max_age_minutes=15
    )

    if recovered_count > 0:
        logger.info(f"Recovered {recovered_count} recently created zip files")

    # Process BioI folders
    order_count = 0
    for bio_folder in bio_folders:
        # Get order folders

        contents = os.listdir(bio_folder)

        # Use the order_folder pattern directly
        order_folders = file_dao.get_folders(bio_folder, pattern=REGEX['order_folder'].pattern)

        # Filter out reinject folders
        original_count = len(order_folders)
        order_folders = [folder for folder in order_folders 
                        if not REGEX['reinject'].search(os.path.basename(folder).lower())]

        logger.info(f"Found {len(order_folders)} order folders in {os.path.basename(bio_folder)}")
        
        for order_folder in order_folders:
            # Check if order already has a zip file
            if file_dao.check_for_zip(order_folder):
                logger.info(f"Skipping {os.path.basename(order_folder)} - already has zip file")
                continue
            
            # Add code to zip the order folder here
            logger.info(f"Zipping {os.path.basename(order_folder)}")
            zip_path = processor.zip_order_folder(order_folder, include_txt=True)
            
            if zip_path:
                # Copy zip to dump folder
                logger.info(f"Copying zip to dump folder: {os.path.basename(zip_path)}")
                file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
                
                order_count += 1
                logger.info(f"Successfully processed {os.path.basename(order_folder)}")
            else:
                logger.warning(f"Failed to zip {os.path.basename(order_folder)}")
    
    # Process PCR folders
    for pcr_folder in pcr_folders:
        # Check if PCR folder already has a zip file
        if file_dao.check_for_zip(pcr_folder):
            logger.info(f"Skipping {os.path.basename(pcr_folder)} - already has zip file")
            continue
        
        # Zip the PCR folder
        logger.info(f"Zipping {os.path.basename(pcr_folder)}")
        zip_path = processor.zip_order_folder(pcr_folder, include_txt=True)
        
        if zip_path:
            # Copy zip to dump folder
            logger.info(f"Copying zip to dump folder: {os.path.basename(zip_path)}")
            file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
            
            order_count += 1  # Use the same count as order folders
            logger.info(f"Successfully processed {os.path.basename(pcr_folder)}")
        else:
            logger.warning(f"Failed to zip {os.path.basename(pcr_folder)}")
    
    # Calculate total processed files
    total_processed = order_count + recovered_count
    
    # Remove empty zip dump folder if nothing was processed
    if total_processed == 0 and os.path.exists(zip_dump_folder) and not os.listdir(zip_dump_folder):
        os.rmdir(zip_dump_folder)
        logger.info("Removed empty zip dump folder")
    
    # Log the total count
    logger.info(f"Total files processed: {total_processed} (New: {order_count}, Recovered: {recovered_count})")
    
    # Update the console message to include both newly zipped and recovered files
    if total_processed == 0:
        print("All done! No files were processed.")
    elif recovered_count > 0 and order_count > 0:
        print(f"All done! {order_count} orders zipped and {recovered_count} recent files recovered.")
    elif recovered_count > 0:
        print(f"All done! {recovered_count} recent files recovered.")
    else:
        print(f"All done! {order_count} orders zipped.")

if __name__ == "__main__":
    main()