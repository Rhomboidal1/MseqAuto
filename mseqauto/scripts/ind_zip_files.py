# ind_zip_files.py
import subprocess
import re
import tkinter as tk
from tkinter import filedialog
# Add parent directory to PYTHONPATH for imports
import os
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parents[2]))
import warnings
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

# Check if running under GUI
GUI_MODE = os.environ.get('MSEQAUTO_GUI_MODE', 'False') == 'True'

def get_folder_from_user():
    """
    Get folder path from user, either from environment variable (in GUI mode) 
    or via file dialog.
    
    Returns:
        Path: Selected folder path, or None if no folder selected
    """
    # Check for GUI-provided folder first
    if GUI_MODE:
        env_folder = os.getenv('MSEQAUTO_DATA_FOLDER')
        if env_folder:
            folder_path = Path(env_folder)
            if folder_path.exists() and folder_path.is_dir():
                return folder_path
            else:
                # Could add logging here: folder exists in env var but is invalid
                pass
    
    # Fall back to file dialog
    root = tk.Tk()
    root.withdraw()
    
    try:
        folder_path = filedialog.askdirectory(
            title="Select today's data folder to zip",
            mustexist=True
        )
        return Path(folder_path) if folder_path else None
    finally:
        root.destroy()

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
    script_dir = Path(__file__).resolve().parent
    log_dir = script_dir / "logs"

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
    zip_dump_folder = data_folder / config.ZIP_DUMP_FOLDER
    if not zip_dump_folder.exists():
        zip_dump_folder.mkdir()
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

        contents = list(bio_folder.iterdir())

        # Use the order_folder pattern directly
        order_folders = file_dao.get_folders(bio_folder, pattern=REGEX['order_folder'].pattern)

        # Filter out reinject folders
        original_count = len(order_folders)
        order_folders = [folder for folder in order_folders 
                        if not REGEX['reinject'].search(folder.name.lower())]

        logger.info(f"Found {len(order_folders)} order folders in {bio_folder.name}")
        
        for order_folder in order_folders:
            # Check if order already has a zip file
            if file_dao.check_for_zip(order_folder):
                logger.info(f"Skipping {order_folder.name} - already has zip file")
                continue
            
            # Add code to zip the order folder here
            logger.info(f"Zipping {order_folder.name}")
            zip_path = processor.zip_order_folder(order_folder, include_txt=True)
            
            if zip_path:
                # Copy zip to dump folder
                logger.info(f"Copying zip to dump folder: {Path(zip_path).name}")
                file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
                
                order_count += 1
                logger.info(f"Successfully processed {order_folder.name}")
            else:
                logger.warning(f"Failed to zip {order_folder.name}")
    
    # Process PCR folders
    for pcr_folder in pcr_folders:
        # Check if PCR folder already has a zip file
        if file_dao.check_for_zip(pcr_folder):
            logger.info(f"Skipping {pcr_folder.name} - already has zip file")
            continue
        
        # Zip the PCR folder
        logger.info(f"Zipping {pcr_folder.name}")
        zip_path = processor.zip_order_folder(pcr_folder, include_txt=True)
        
        if zip_path:
            # Copy zip to dump folder
            logger.info(f"Copying zip to dump folder: {Path(zip_path).name}")
            file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
            
            order_count += 1  # Use the same count as order folders
            logger.info(f"Successfully processed {pcr_folder.name}")
        else:
            logger.warning(f"Failed to zip {pcr_folder.name}")
    
    # Calculate total processed files
    total_processed = order_count + recovered_count
    
    # Remove empty zip dump folder if nothing was processed
    if total_processed == 0 and zip_dump_folder.exists() and not list(zip_dump_folder.iterdir()):
        zip_dump_folder.rmdir()
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