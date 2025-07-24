# validate_zip_files.py
import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog
import warnings
from pathlib import Path

# Add parent directory to PYTHONPATH for imports
sys.path.append(str(Path(__file__).parents[2]))
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")


def get_folder_from_user():
    """Get folder selection from user"""
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()
    
    folder_path = filedialog.askdirectory(
        title="Select folder containing zip files to validate",
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

    # Import package modules
    from mseqauto.config import MseqConfig # type: ignore
    from mseqauto.core import FileSystemDAO, FolderProcessor # type: ignore
    from mseqauto.utils import setup_logger, ExcelDAO # type: ignore
    
    # Setup logging
    script_dir = Path(__file__).parent.resolve()
    log_dir = script_dir / "logs"
    logger = setup_logger("validate_zip_files", log_dir=log_dir)
    logger.info("Starting zip file validation...")
    
    # Initialize components
    config = MseqConfig()
    file_dao = FileSystemDAO(config, logger=logger)
    excel_dao = ExcelDAO(config)
    processor = FolderProcessor(file_dao, None, config, logger=logger.info)
    logger.info(f"Using folder: {data_folder}")

    # Run batch file to generate order key
    try:
        logger.info(f"Running batch file: {config.BATCH_FILE_PATH}")
        subprocess.run(config.BATCH_FILE_PATH, check=True)
        logger.info("Batch file completed successfully")
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        logger.error(f"Batch file error: {e}")
        print(f"Error: Batch file {config.BATCH_FILE_PATH} failed to run")
        return

    # Load order key
    order_key = file_dao.load_order_key(config.KEY_FILE_PATH)
    if order_key is None:
        logger.error("Failed to load order key")
        print("Error: Failed to load order key")
        return
    logger.info("Order key loaded")

    # Setup Excel files
    excel_filename = f"zip order summary - {datetime.now().strftime('%Y-%m-%d')}.xlsx"
    excel_path = Path(data_folder) / excel_filename
    summary_exists = excel_path.exists()
    logger.info(f"Summary exists: {summary_exists}")

    # Load existing workbook if it exists
    existing_workbook = None
    existing_worksheet = None
    if summary_exists:
        existing_workbook = excel_dao.load_workbook(excel_path)
        if not existing_workbook:
            print("Error: Could not load existing summary. Make sure the file is not open in Excel.")
            return
        existing_worksheet = existing_workbook.active
        logger.info("Existing workbook loaded successfully")
    
    # Create new workbook for this session's data
    new_workbook = excel_dao.create_workbook()
    new_worksheet = new_workbook.active

    # Get all order folders using FolderProcessor method
    order_folders = processor.get_order_folders_for_validation(data_folder)
    logger.info(f"Found {len(order_folders)} total order folders to check")

    # Find FB-PCR zip files
    fb_pcr_zips = file_dao.find_fb_pcr_zips(data_folder)
    logger.info(f"Found {len(fb_pcr_zips)} FB-PCR zip files to process")
    
    # Find plate folder zip files
    plate_zips = file_dao.find_plate_folder_zips(data_folder)
    logger.info(f"Found {len(plate_zips)} plate folder zip files to process")
    
    # Set headers based on what data types we have
    if len(order_folders) > 0 and (len(fb_pcr_zips) > 0 or len(plate_zips) > 0):
        # Mixed data - use validation headers as primary, other data will use available columns
        excel_dao.set_validation_headers(new_worksheet)
    elif len(fb_pcr_zips) > 0 and len(plate_zips) > 0:
        # Both FB-PCR and plate data - use validation headers to accommodate both
        excel_dao.set_validation_headers(new_worksheet)
    elif len(fb_pcr_zips) > 0:
        # Only FB-PCR data
        excel_dao.set_fb_pcr_headers(new_worksheet)
    elif len(plate_zips) > 0:
        # Only plate data
        excel_dao.set_plate_headers(new_worksheet)
    else:
        # Only validation data
        excel_dao.set_validation_headers(new_worksheet)

    # Process each order folder
    order_count = 0
    new_row_count = 2  # Start after headers

    for order_folder, i_number in order_folders:
        # Find zip file using existing FolderProcessor method
        zip_path = processor.find_zip_file(order_folder)  # Use existing method
        if not zip_path:
            logger.info(f"Skipping {Path(order_folder).name} - no zip file")
            continue

        # Get order number
        order_number = processor.get_order_number_from_folder_name(order_folder)
        if not order_number:
            logger.warning(f"Could not extract order number from {Path(order_folder).name}")
            continue

        logger.info(f"Processing order: I-{i_number}, Order: {order_number}")

        # Check if order already validated with current or newer zip
        if summary_exists:
            _, _, existing_mod_time = excel_dao.find_order_in_summary(existing_worksheet, order_number)
            current_mod_time = Path(zip_path).stat().st_mtime
            
            if existing_mod_time and float(existing_mod_time) >= current_mod_time:
                logger.info(f"Skipping {Path(order_folder).name} - already validated with current or newer zip")
                continue

        # Validate zip contents using FolderProcessor method
        logger.info(f"Validating zip contents for order {order_number}")
        validation_result = processor.validate_zip_contents(zip_path, i_number, order_number, order_key)

        if validation_result:
            order_count += 1
            
            # Check if this is an Andreev order
            is_andreev = config.ANDREEV_NAME.lower() in Path(order_folder).name.lower()
            
            # Add validation results to new worksheet
            new_row_count = excel_dao.add_validation_result(
                new_worksheet, new_row_count, validation_result, zip_path,
                i_number, order_number, is_andreev
            )
            
            # If updating existing order, mark old one as resolved
            if summary_exists and existing_mod_time:
                logger.info(f"Marking previous version of order {order_number} as resolved")
                excel_dao.resolve_order_status(existing_worksheet, order_number)

    # Process FB-PCR zip files
    fb_pcr_count = 0
    for zip_path, pcr_number, order_number, version in fb_pcr_zips:
        logger.info(f"Processing FB-PCR zip: PCR-{pcr_number}, Order: {order_number}, Version: {version}")
        
        # Process FB-PCR zip to get file count
        fb_pcr_result = processor.process_fb_pcr_zip(zip_path, pcr_number, order_number, version)
        
        if fb_pcr_result:
            fb_pcr_count += 1
            
            # Determine if we're using mixed headers
            mixed_headers = len(order_folders) > 0 or len(plate_zips) > 0
            
            # Add FB-PCR results to new worksheet
            new_row_count = excel_dao.add_fb_pcr_result(
                new_worksheet, new_row_count, fb_pcr_result, zip_path, mixed_headers
            )

    # Process plate folder zip files
    plate_count = 0
    for zip_path, plate_number, description in plate_zips:
        logger.info(f"Processing plate zip: P{plate_number}, Description: {description}")
        
        # Process plate zip to get file count
        plate_result = processor.process_plate_zip(zip_path, plate_number, description)
        
        if plate_result:
            plate_count += 1
            
            # Determine if we're using mixed headers
            mixed_headers = len(order_folders) > 0 or len(fb_pcr_zips) > 0
            
            # Add plate results to new worksheet
            new_row_count = excel_dao.add_plate_result(
                new_worksheet, new_row_count, plate_result, zip_path, mixed_headers
            )

    # Update order count to include FB-PCR and plate zips
    total_processed = order_count + fb_pcr_count + plate_count

    # Finalize and save results
    if total_processed > 0:
        excel_dao.finalize_workbook(new_worksheet, add_break_at_end=summary_exists)
        logger.info(f"Processed {order_count} orders, {fb_pcr_count} FB-PCR zips, and {plate_count} plate zips")

        if summary_exists:
            # Update existing summary
            logger.info("Updating existing summary with new data")
            success = excel_dao.update_existing_summary(existing_workbook, new_workbook, excel_path)
        else:
            # Create new summary
            logger.info("Creating new summary file")
            success = excel_dao.save_with_error_handling(new_workbook, excel_path)

        if not success:
            logger.error("Failed to save summary")
            print("Error: Could not save summary. Make sure the file is not open in Excel.")
            return

    logger.info(f"Validation complete. Total orders processed: {order_count}, FB-PCR zips processed: {fb_pcr_count}, Plate zips processed: {plate_count}")
    if total_processed > 0:
        # Build output message based on what was processed
        processed_parts = []
        if order_count > 0:
            processed_parts.append(f"{order_count} orders")
        if fb_pcr_count > 0:
            processed_parts.append(f"{fb_pcr_count} FB-PCR zips")
        if plate_count > 0:
            processed_parts.append(f"{plate_count} plate zips")
        
        processed_text = " and ".join(processed_parts)
        print(f"All done! {processed_text} validated. Summary saved to {excel_filename}")
    else:
        print("All done! No new orders, FB-PCR zips, or plate zips found to validate.")


if __name__ == "__main__":
    main()