# validate_zip_files.py
import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog
import warnings

# Add parent directory to PYTHONPATH for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
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
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "logs")
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
        subprocess.run(config.BATCH_FILE_PATH, shell=True, check=True)
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
    excel_path = os.path.join(data_folder, excel_filename)
    summary_exists = os.path.exists(excel_path)
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
    excel_dao.set_validation_headers(new_worksheet)

    # Get all order folders using FolderProcessor method
    order_folders = processor.get_order_folders_for_validation(data_folder)
    logger.info(f"Found {len(order_folders)} total order folders to check")

    # Process each order folder
    order_count = 0
    new_row_count = 2  # Start after headers

    for order_folder, i_number in order_folders:
        # Find zip file using existing FolderProcessor method
        zip_path = processor.find_zip_file(order_folder)  # Use existing method
        if not zip_path:
            logger.info(f"Skipping {os.path.basename(order_folder)} - no zip file")
            continue

        # Get order number
        order_number = processor.get_order_number_from_folder_name(order_folder)
        if not order_number:
            logger.warning(f"Could not extract order number from {os.path.basename(order_folder)}")
            continue

        logger.info(f"Processing order: I-{i_number}, Order: {order_number}")

        # Check if order already validated with current or newer zip
        if summary_exists:
            _, _, existing_mod_time = excel_dao.find_order_in_summary(existing_worksheet, order_number)
            current_mod_time = os.path.getmtime(zip_path)
            
            if existing_mod_time and float(existing_mod_time) >= current_mod_time:
                logger.info(f"Skipping {os.path.basename(order_folder)} - already validated with current or newer zip")
                continue

        # Validate zip contents using FolderProcessor method
        logger.info(f"Validating zip contents for order {order_number}")
        validation_result = processor.validate_zip_contents(zip_path, i_number, order_number, order_key)

        if validation_result:
            order_count += 1
            
            # Check if this is an Andreev order
            is_andreev = config.ANDREEV_NAME.lower() in os.path.basename(order_folder).lower()
            
            # Add validation results to new worksheet
            new_row_count = excel_dao.add_validation_result(
                new_worksheet, new_row_count, validation_result, zip_path,
                i_number, order_number, is_andreev
            )
            
            # If updating existing order, mark old one as resolved
            if summary_exists and existing_mod_time:
                logger.info(f"Marking previous version of order {order_number} as resolved")
                excel_dao.resolve_order_status(existing_worksheet, order_number)

    # Finalize and save results
    if order_count > 0:
        excel_dao.finalize_workbook(new_worksheet, add_break_at_end=summary_exists)
        logger.info(f"Processed {order_count} orders")

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

    logger.info(f"Validation complete. Total orders processed: {order_count}")
    if order_count > 0:
        print(f"All done! {order_count} orders validated. Summary saved to {excel_filename}")
    else:
        print("All done! No new orders found to validate.")


if __name__ == "__main__":
    main()