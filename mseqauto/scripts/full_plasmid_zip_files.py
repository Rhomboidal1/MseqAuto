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

def has_plasmid_files(folder_path):
    """Check if a folder contains plasmid sequencing files"""
    try:
        for item in os.listdir(folder_path):
            if item.endswith('.final.fasta') or item.endswith('wf-clone-validation-report.html'):
                return True
    except (OSError, PermissionError):
        pass
    return False

def has_standard_zip(folder_path, config):
    """Check if a folder contains a standard zip file (not a raw data zip)"""
    try:
        for item in os.listdir(folder_path):
            if item.endswith('.zip'):
                # Skip raw data zips (ending with rd#.zip)
                if config.REGEX_PATTERNS['raw_data_zip'].search(item):
                    continue
                # Found a non-raw-data zip file
                return True
    except (OSError, PermissionError):
        pass
    return False

def should_ignore_folder(folder_name):
    """Check if a folder should be ignored"""
    return folder_name.lower() == "posted"

def process_plasmid_folder(folder_path, file_dao, processor, order_key, zip_dump_folder, logger, config):
    """Process a single plasmid folder and return success status"""
    folder_name = os.path.basename(folder_path)

    # Check if folder contains plasmid files
    if not has_plasmid_files(folder_path):
        logger.info(f"Skipping {folder_name} - not a plasmid sequencing order")
        return False

    # Check if order already has a standard zip file (ignore raw data zips)
    if has_standard_zip(folder_path, config):
        logger.info(f"Skipping {folder_name} - already has standard zip file")
        return False

    # Get order number for reference (works for both BioI and P# patterns)
    order_number = None
    match = config.REGEX_PATTERNS['order_number_extract'].search(folder_name)
    if match:
        order_number = match.group(1)

    # Zip the plasmid folder
    logger.info(f"Zipping plasmid folder: {folder_name}")
    zip_path = processor.zip_full_plasmid_order_folder(
        folder_path,
        order_key=order_key,
        order_number=order_number,
        use_7zip=USE_7ZIP,
        compression_level=COMPRESSION_LEVEL
    )

    if zip_path:
        # Copy zip to dump folder
        logger.info(f"Copying zip to dump folder: {os.path.basename(zip_path)}")
        file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
        logger.info(f"Successfully processed {folder_name}")
        return True
    else:
        logger.warning(f"Failed to zip {folder_name}")
        return False

def main():
    # Get folder path first before any package imports
    data_folder = get_folder_from_user()

    if not data_folder:
        print("No folder selected, exiting")
        return

    # ONLY NOW import package modules
    from mseqauto.config import MseqConfig #type: ignore
    from mseqauto.core import FileSystemDAO, FolderProcessor #type: ignore
    from mseqauto.utils import setup_logger #type: ignore

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

    # Create zip dump folder at the selected folder level
    zip_dump_folder = os.path.join(data_folder, config.ZIP_DUMP_FOLDER)
    if not os.path.exists(zip_dump_folder):
        os.makedirs(zip_dump_folder)
        logger.info(f"Created zip dump folder: {zip_dump_folder}")

    # Load order key for getting sample names
    order_key = file_dao.load_order_key(config.KEY_FILE_PATH)
    logger.info("Order key loaded")

    # Define patterns
    bioi_folder_pattern = config.REGEX_PATTERNS['bioi_folder'].pattern
    order_folder_pattern = config.REGEX_PATTERNS['order_folder'].pattern
    plate_folder_pattern = config.REGEX_PATTERNS['plate_folder']

    total_processed = 0

    # Process folders in the selected directory
    logger.info("Scanning selected directory for folders to process...")

    try:
        for item in os.listdir(data_folder):
            item_path = os.path.join(data_folder, item)

            # Skip if not a directory
            if not os.path.isdir(item_path):
                continue

            # Skip ignored folders
            if should_ignore_folder(item):
                logger.info(f"Ignoring folder: {item}")
                continue

            # Check if this is a P# folder (plate folder)
            if plate_folder_pattern.match(item):
                logger.info(f"Found P# folder: {item}")
                if process_plasmid_folder(item_path, file_dao, processor, order_key, zip_dump_folder, logger, config):
                    total_processed += 1
                continue

            # Check if this is a BioI folder
            if re.search(bioi_folder_pattern, item, re.IGNORECASE):
                logger.info(f"Found BioI folder: {item}, scanning for order folders...")

                # Look for order folders within this BioI folder
                try:
                    order_folders = file_dao.get_folders(item_path, pattern=order_folder_pattern)

                    # Filter out reinject folders
                    order_folders = [folder for folder in order_folders
                                    if not config.REGEX_PATTERNS['reinject'].search(os.path.basename(folder).lower())]

                    logger.info(f"Found {len(order_folders)} potential order folders in {item}")

                    for order_folder in order_folders:
                        if process_plasmid_folder(order_folder, file_dao, processor, order_key, zip_dump_folder, logger, config):
                            total_processed += 1

                except Exception as e:
                    logger.error(f"Error processing BioI folder {item}: {e}")
                continue

            # Check if this is a direct order folder (BioI-#####_CustomerName_OrderNumber pattern)
            if re.search(order_folder_pattern, item, re.IGNORECASE):
                # Skip reinject folders
                if config.REGEX_PATTERNS['reinject'].search(item.lower()):
                    logger.info(f"Skipping reinject folder: {item}")
                    continue

                logger.info(f"Found direct order folder: {item}")
                if process_plasmid_folder(item_path, file_dao, processor, order_key, zip_dump_folder, logger, config):
                    total_processed += 1
                continue

            # Log unrecognized folders for debugging
            logger.debug(f"Unrecognized folder pattern: {item}")

    except Exception as e:
        logger.error(f"Error scanning directory {data_folder}: {e}")

    # Remove empty zip dump folder if nothing was processed
    if total_processed == 0 and os.path.exists(zip_dump_folder) and not os.listdir(zip_dump_folder):
        os.rmdir(zip_dump_folder)
        logger.info("Removed empty zip dump folder")

    # Log the total count
    logger.info(f"Total plasmid folders processed: {total_processed}")

    # Update the console message
    if total_processed == 0:
        print("All done! No plasmid folders were processed.")
    else:
        print(f"All done! {total_processed} plasmid folders zipped.")
        print(f"Compression: {'7-Zip' if USE_7ZIP else 'Python zipfile'}, level {COMPRESSION_LEVEL}")

if __name__ == "__main__":
    main()