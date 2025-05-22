# full_plasmid_zip_files.py
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog
import warnings

sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

# Configuration settings
USE_7ZIP = True          # Set to True to use 7-Zip compression (slower but better compression)
                          # Set to False to use Python's zipfile (faster but slightly larger files)
COMPRESSION_LEVEL = 6     # Compression level (0-9):
                          # 0 = No compression (fastest, largest file size)
                          # 6 = Default compression (balance of speed and size)
                          # 9 = Maximum compression (slowest, smallest file size)

def get_folder_from_user():
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force an update
    
    folder_path = filedialog.askdirectory(
        title="Select data folder containing full plasmid sequencing orders",
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
    from mseqauto.core import FileSystemDAO, FolderProcessor
    from mseqauto.utils import setup_logger
    
    # Get the script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "logs")

    # Setup logger
    logger = setup_logger("full_plasmid_zip_files", log_dir=log_dir)

    logger.info("Starting full plasmid zip files...")
    logger.info(f"Compression method: {'7-Zip' if USE_7ZIP else 'Python zipfile'}")
    logger.info(f"Compression level: {COMPRESSION_LEVEL} (0=None, 6=Default, 9=Maximum)")
    
    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
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
    
    # Load order key for getting sample names
    order_key = file_dao.load_order_key(config.KEY_FILE_PATH)
    logger.info("Order key loaded")
    
    # Find folders that match plasmid order pattern
    # This pattern should match your plasmid order folder naming convention
    # Assuming they follow the same pattern as regular orders: BioI-#####_CustomerName_OrderNumber
    plasmid_folder_pattern = config.REGEX_PATTERNS['order_folder'].pattern
    
    # Get all BioI folders
    bio_folders = file_dao.get_folders(data_folder, pattern=config.REGEX_PATTERNS['bioi_folder'].pattern)
    logger.info(f"Found {len(bio_folders)} BioI folders")
    
    # Process order folders within BioI folders
    order_count = 0
    recovered_count = 0
    
    for bio_folder in bio_folders:
        # Get order folders using the pattern
        order_folders = file_dao.get_folders(bio_folder, pattern=plasmid_folder_pattern)
        
        # Filter out reinject folders
        order_folders = [folder for folder in order_folders 
                        if not config.REGEX_PATTERNS['reinject'].search(os.path.basename(folder).lower())]
        
        logger.info(f"Found {len(order_folders)} potential plasmid order folders in {os.path.basename(bio_folder)}")
        
        for order_folder in order_folders:
            # Check if folder contains plasmid files
            folder_name = os.path.basename(order_folder)
            has_plasmid_files = False
            
            # Check for plasmid sequencing files (just check for one type as a quick filter)
            for item in os.listdir(order_folder):
                if item.endswith('.final.fasta') or item.endswith('wf-clone-validation-report.html'):
                    has_plasmid_files = True
                    break
            
            if not has_plasmid_files:
                logger.info(f"Skipping {folder_name} - not a plasmid sequencing order")
                continue
                
            # Check if order already has a zip file
            if file_dao.check_for_zip(order_folder):
                logger.info(f"Skipping {folder_name} - already has zip file")
                continue
            
            # Get order number for reference
            order_number = None
            match = re.search(r'_(\d+)$', folder_name)
            if match:
                order_number = match.group(1)
            
            # Zip the plasmid order folder - Pass the USE_7ZIP and COMPRESSION_LEVEL settings
            logger.info(f"Zipping plasmid order: {folder_name}")
            zip_path = processor.zip_full_plasmid_order_folder(
                order_folder, 
                order_key=order_key, 
                order_number=order_number,
                use_7zip=USE_7ZIP,
                compression_level=COMPRESSION_LEVEL
            )
            
            if zip_path:
                # Copy zip to dump folder
                logger.info(f"Copying zip to dump folder: {os.path.basename(zip_path)}")
                file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
                
                order_count += 1
                logger.info(f"Successfully processed {folder_name}")
            else:
                logger.warning(f"Failed to zip {folder_name}")
    
    # Look for direct plasmid order folders in the main directory (outside BioI folders)
    direct_order_folders = file_dao.get_folders(data_folder, pattern=plasmid_folder_pattern)
    direct_order_folders = [folder for folder in direct_order_folders 
                           if not config.REGEX_PATTERNS['reinject'].search(os.path.basename(folder).lower())]
    
    logger.info(f"Found {len(direct_order_folders)} potential direct plasmid order folders")
    
    for order_folder in direct_order_folders:
        # Check if folder contains plasmid files
        folder_name = os.path.basename(order_folder)
        has_plasmid_files = False
        
        # Check for plasmid sequencing files (just check for one type as a quick filter)
        for item in os.listdir(order_folder):
            if item.endswith('.final.fasta') or item.endswith('wf-clone-validation-report.html'):
                has_plasmid_files = True
                break
        
        if not has_plasmid_files:
            logger.info(f"Skipping {folder_name} - not a plasmid sequencing order")
            continue
            
        # Check if order already has a zip file
        if file_dao.check_for_zip(order_folder):
            logger.info(f"Skipping {folder_name} - already has zip file")
            continue
        
        # Get order number for reference
        order_number = None
        match = re.search(r'_(\d+)$', folder_name)
        if match:
            order_number = match.group(1)
        
        # Zip the plasmid order folder - Pass the USE_7ZIP and COMPRESSION_LEVEL settings
        logger.info(f"Zipping direct plasmid order: {folder_name}")
        zip_path = processor.zip_full_plasmid_order_folder(
            order_folder, 
            order_key=order_key, 
            order_number=order_number,
            use_7zip=USE_7ZIP,
            compression_level=COMPRESSION_LEVEL
        )
        
        if zip_path:
            # Copy zip to dump folder
            logger.info(f"Copying zip to dump folder: {os.path.basename(zip_path)}")
            file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
            
            order_count += 1
            logger.info(f"Successfully processed {folder_name}")
        else:
            logger.warning(f"Failed to zip {folder_name}")
    
    # Remove empty zip dump folder if nothing was processed
    if order_count == 0 and os.path.exists(zip_dump_folder) and not os.listdir(zip_dump_folder):
        os.rmdir(zip_dump_folder)
        logger.info("Removed empty zip dump folder")
    
    # Log the total count
    total_processed = order_count
    logger.info(f"Total plasmid orders processed: {total_processed}")
    
    # Update the console message
    if total_processed == 0:
        print("All done! No plasmid orders were processed.")
    else:
        print(f"All done! {total_processed} plasmid orders zipped.")
        print(f"Compression: {'7-Zip' if USE_7ZIP else 'Python zipfile'}, level {COMPRESSION_LEVEL}")

if __name__ == "__main__":
    main()