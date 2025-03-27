# plate_sort_files.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
from datetime import datetime
import re


def get_folder_from_user():
    """Get folder selection from user"""
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force an update to ensure dialog appears properly
    
    folder_path = filedialog.askdirectory(
        title="Select today's data folder to sort plate files",
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
    # Get folder path first before any package imports
    data_folder = get_folder_from_user()
    
    if not data_folder:
        print("No folder selected, exiting")
        return

    # Now import package modules
    from mseqauto.config import MseqConfig
    from mseqauto.core import FileSystemDAO, FolderProcessor
    from mseqauto.utils import setup_logger

    # Setup logger
    logger = setup_logger("plate_sort_files")
    logger.info("Starting plate sort files...")
    
    # Check for 32-bit Python requirement
    if sys.maxsize > 2**32:
        py32_path = MseqConfig.PYTHON32_PATH
        if os.path.exists(py32_path) and py32_path != sys.executable:
            script_path = os.path.abspath(__file__)
            subprocess.run([py32_path, script_path])
            sys.exit(0)
        else:
            logger.warning("32-bit Python not specified or same as current interpreter")
            logger.warning("Continuing with current Python interpreter")
    
    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")
    processor = FolderProcessor(file_dao, None, config, logger=logger.info)
    logger.info("Folder processor initialized")
    
    logger.info(f"Using folder: {data_folder}")
    
    # Get plate folders - plates start with 'p' followed by digits
    plate_folders = file_dao.get_folders(data_folder, pattern=r'p\d+.+')
    logger.info(f"Found {len(plate_folders)} plate folders")
    
    if not plate_folders:
        logger.warning("No plate folders found, exiting")
        print("No plate folders found, exiting")
        return
    
    # Process each plate folder
    for i, folder in enumerate(plate_folders):
        folder_name = os.path.basename(folder)
        logger.info(f"Processing plate folder {i+1}/{len(plate_folders)}: {folder_name}")
        
        # Sort controls by creating Controls folder and moving control files
        controls_folder = os.path.join(folder, config.CONTROLS_FOLDER)
        if not os.path.exists(controls_folder):
            os.makedirs(controls_folder)
            logger.info(f"Created Controls folder in {folder_name}")
            
        # Sort blanks by creating Blank folder and moving blank files
        blank_folder = os.path.join(folder, config.BLANK_FOLDER)
        if not os.path.exists(blank_folder):
            os.makedirs(blank_folder)
            logger.info(f"Created Blank folder in {folder_name}")
        
        # Process all files in the folder
        files_processed = 0
        for file_name in os.listdir(folder):
            file_path = os.path.join(folder, file_name)
            
            # Skip directories and non-AB1 files
            if os.path.isdir(file_path) or not file_name.endswith(config.ABI_EXTENSION):
                continue
            
            # Check if file is a control
            if any(control in file_name.lower() for control in [c.lower() for c in config.PLATE_CONTROLS]):
                # Move to Controls folder
                target_path = os.path.join(controls_folder, file_name)
                file_dao.move_file(file_path, target_path)
                logger.info(f"Moved control file {file_name} to Controls folder")
                files_processed += 1
                continue
                
            # Check if file is a blank - for plates, blanks often have format like "01A__.ab1"
            if file_dao.is_blank_file(file_name):
                # Move to Blank folder
                target_path = os.path.join(blank_folder, file_name)
                file_dao.move_file(file_path, target_path)
                logger.info(f"Moved blank file {file_name} to Blank folder")
                files_processed += 1
                continue
                
            # Remove braces from filenames for regular files
            if '{' in file_name or '}' in file_name:
                clean_name = re.sub(r'{.*?}', '', file_name)
                if clean_name != file_name:  # Only rename if actually changed
                    new_path = os.path.join(folder, clean_name)
                    try:
                        os.rename(file_path, new_path)
                        logger.info(f"Renamed: {file_name} -> {clean_name}")
                        files_processed += 1
                    except Exception as e:
                        logger.error(f"Failed to rename {file_name}: {str(e)}")
        
        logger.info(f"Processed {files_processed} files in {folder_name}")
        
        # Check if Controls or Blank folders are empty and remove them if so
        for special_folder in [controls_folder, blank_folder]:
            if os.path.exists(special_folder) and not os.listdir(special_folder):
                os.rmdir(special_folder)
                folder_type = "Controls" if special_folder == controls_folder else "Blank"
                logger.info(f"Removed empty {folder_type} folder from {folder_name}")
    
    logger.info("All plate folders processed successfully")
    print("All done!")


if __name__ == "__main__":
    main()