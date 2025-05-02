# plate_zip_files.py
import tkinter as tk
from tkinter import filedialog
import subprocess
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
import warnings
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

#print(sys.path)

def get_folder_from_user():
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
    logger = setup_logger("plate_zip_files", log_dir=log_dir)

    logger.info("Starting Plate zip files...")
    
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