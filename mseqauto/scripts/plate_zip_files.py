# plate_zip_files.py
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
        title="Select folder containing plate folders to zip",
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
    logger = setup_logger("plate_zip_files")
    logger.info("Starting plate zip files...")
    
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
    
    # Create zip dump folder
    zip_dump_folder = os.path.join(data_folder, config.ZIP_DUMP_FOLDER)
    if not os.path.exists(zip_dump_folder):
        os.makedirs(zip_dump_folder)
        logger.info(f"Created zip dump folder: {zip_dump_folder}")
    
    # Get plate folders
    plate_folders = file_dao.get_folders(data_folder, pattern=r'p\d+.+')
    logger.info(f"Found {len(plate_folders)} plate folders")
    
    if not plate_folders:
        logger.warning("No plate folders found, exiting")
        print("No plate folders found, exiting")
        return
    
    # Process each plate folder
    plate_count = 0
    for plate_folder in plate_folders:
        # Check if plate folder already has a zip file
        if file_dao.check_for_zip(plate_folder):
            logger.info(f"Skipping {os.path.basename(plate_folder)} - already has zip file")
            continue
        
        # Check if this is an FSA plate
        has_fsa = file_dao.contains_file_type(plate_folder, config.FSA_EXTENSION)
        
        if has_fsa:
            # Zip FSA files only
            logger.info(f"Zipping FSA files in {os.path.basename(plate_folder)}")
            zip_path = processor.zip_order_folder(plate_folder, fsa_only=True)
        else:
            # Zip AB1 and txt files
            logger.info(f"Zipping AB1 and txt files in {os.path.basename(plate_folder)}")
            zip_path = processor.zip_order_folder(plate_folder, fsa_only=False)
        
        if zip_path:
            # Copy zip to dump folder
            logger.info(f"Copying zip to dump folder: {os.path.basename(zip_path)}")
            file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
            plate_count += 1
            
            logger.info(f"Successfully processed {os.path.basename(plate_folder)}")
        else:
            logger.warning(f"Failed to zip {os.path.basename(plate_folder)}")
    
    # Remove empty zip dump folder if nothing was processed
    if plate_count == 0 and os.path.exists(zip_dump_folder) and not os.listdir(zip_dump_folder):
        os.rmdir(zip_dump_folder)
        logger.info("Removed empty zip dump folder")
    
    logger.info(f"Total plates zipped: {plate_count}")
    print(f"All done! {plate_count} plates zipped.")

if __name__ == "__main__":
    main()