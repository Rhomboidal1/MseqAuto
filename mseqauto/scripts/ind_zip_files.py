# ind_zip_files.py
import subprocess
import tkinter as tk
from tkinter import filedialog
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

print(sys.path)

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
    from mseqauto.config import MseqConfig
    from mseqauto.core import FileSystemDAO, FolderProcessor, OSCompatibilityManager
    from mseqauto.utils import setup_logger
    
    # Setup logger
    logger = setup_logger("ind_sort_files")

    # Immediately check for 32-bit Python requirements
    OSCompatibilityManager.py32_check(
        script_path=__file__,
        logger=logger
    )

    logger.info("Starting IND sort files...")
    
    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")
    processor = FolderProcessor(file_dao, None, config, logger=logger.info)
    logger.info("Folder processor initialized")
    logger.info(f"Using folder: {data_folder}")
    
    # Create zip dump folder
    zip_dump_folder = os.path.join(data_folder, config.ZIP_DUMP_FOLDER)
    if not os.path.exists(zip_dump_folder):
        os.makedirs(zip_dump_folder)
        logger.info(f"Created zip dump folder: {zip_dump_folder}")
    
    # Get BioI folders
    bio_folders = file_dao.get_folders(data_folder, pattern=r'bioi-\d+')
    logger.info(f"Found {len(bio_folders)} BioI folders")
    
    # Get PCR folders
    pcr_folders = file_dao.get_folders(data_folder, pattern=r'fb-pcr\d+_\d+')
    logger.info(f"Found {len(pcr_folders)} PCR folders")
    
    # Process BioI folders
    order_count = 0
    for bio_folder in bio_folders:
        # Get order folders
        order_folders = processor.get_order_folders(bio_folder)
        logger.info(f"Found {len(order_folders)} order folders in {os.path.basename(bio_folder)}")
        
        for order_folder in order_folders:
            # Check if order already has a zip file
            if file_dao.check_for_zip(order_folder):
                logger.info(f"Skipping {os.path.basename(order_folder)} - already has zip file")
                continue
            
            # Special handling for Andreev orders
            is_andreev = config.ANDREEV_NAME in os.path.basename(order_folder).lower()
            
            # Zip the order folder
            logger.info(f"Zipping {os.path.basename(order_folder)}")
            zip_path = processor.zip_order_folder(order_folder, include_txt=not is_andreev)
            
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
            
            logger.info(f"Successfully processed {os.path.basename(pcr_folder)}")
        else:
            logger.warning(f"Failed to zip {os.path.basename(pcr_folder)}")
    
    # Remove empty zip dump folder if nothing was processed
    if order_count == 0 and os.path.exists(zip_dump_folder) and not os.listdir(zip_dump_folder):
        os.rmdir(zip_dump_folder)
        logger.info("Removed empty zip dump folder")
    
    logger.info(f"Total orders zipped: {order_count}")
    print(f"All done! {order_count} orders zipped.")

if __name__ == "__main__":
    main()