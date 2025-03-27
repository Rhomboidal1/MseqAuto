# validate_zip_files.py
import os
import subprocess
import sys
import tkinter as tk
from datetime import datetime
from tkinter import filedialog

from mseqauto.core import FileSystemDAO, FolderProcessor
from mseqauto.utils import ExcelDAO, setup_logger
from mseqauto.config import MseqConfig

# Check for 32-bit Python requirement
if sys.maxsize > 2 ** 32:
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
    # Setup logger
    logger = setup_logger("validate_zip_files")
    logger.info("Starting zip file validation...")

    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")
    excel_dao = ExcelDAO(config)
    logger.info("ExcelDAO initialized")
    processor = FolderProcessor(file_dao, None, config, logger=logger.info)
    logger.info("Folder processor initialized")

    # Run batch file to generate order key
    try:
        logger.info(f"Running batch file: {config.BATCH_FILE_PATH}")
        subprocess.run(config.BATCH_FILE_PATH, shell=True, check=True)
        logger.info("Batch file completed successfully")
    except subprocess.CalledProcessError:
        logger.error(f"Batch file {config.BATCH_FILE_PATH} failed to run")
        print(f"Error: Batch file {config.BATCH_FILE_PATH} failed to run")
        return

    # Select folder
    data_folder = get_folder_from_user()

    if not data_folder:
        logger.error("No folder selected, exiting")
        print("No folder selected, exiting")
        return

    logger.info(f"Using folder: {data_folder}")

    # Load order key
    order_key = file_dao.load_order_key(config.KEY_FILE_PATH)
    logger.info("Order key loaded")

    # Create Excel file name
    excel_filename = f"zip order summary - {datetime.now().strftime('%Y-%m-%d')}.xlsx"
    excel_path = os.path.join(data_folder, excel_filename)

    # Check if summary file already exists
    summary_exists = os.path.exists(excel_path)

    # Create or load workbook
    if summary_exists:
        logger.info(f"Loading existing summary: {excel_filename}")
        try:
            workbook = excel_dao.load_workbook(excel_path)
            if not workbook:
                raise Exception("Failed to load existing workbook")
            worksheet = workbook.active
        except Exception as e:
            logger.error(f"Error loading existing workbook: {e}")
            print(f"Error: Could not load existing summary. Make sure the file is not open in Excel.")
            return
    else:
        logger.info("Creating new summary workbook")
        workbook = excel_dao.create_workbook()
        worksheet = workbook.active

        # Add headers
        headers = ['I Number', 'Order Number', 'Status', 'Zip Filename',
                   'Order Items', 'File Names', 'Match Status', 'Zip Timestamp']
        excel_dao.set_header_row(worksheet, headers)

    # Get BioI folders
    bio_folders = file_dao.get_folders(data_folder, pattern=r'bioi-\d+')
    logger.info(f"Found {len(bio_folders)} BioI folders")

    # Track number of orders validated
    order_count = 0
    row_count = worksheet.max_row + 1 if summary_exists else 2

    # Process each BioI folder
    for bio_folder in bio_folders:
        # Get order folders
        order_folders = file_dao.get_folders(bio_folder, pattern=r'bioi-\d+_.+_\d+')
        logger.info(f"Found {len(order_folders)} order folders in {os.path.basename(bio_folder)}")

        for order_folder in order_folders:
            # Skip if no zip file
            zip_path = processor.find_zip_file(order_folder)
            if not zip_path:
                logger.info(f"Skipping {os.path.basename(order_folder)} - no zip file")
                continue

            # Get order information
            i_number = file_dao.get_inumber_from_name(bio_folder)
            order_number = processor.get_order_number_from_folder_name(order_folder)

            # Check if order already in summary with same zip timestamp
            if summary_exists:
                current_mod_time = os.path.getmtime(zip_path)
                existing_mod_time = processor.get_zip_mod_time(worksheet, order_number)

                if existing_mod_time and float(existing_mod_time) >= current_mod_time:
                    logger.info(f"Skipping {os.path.basename(order_folder)} - already validated")
                    continue

            # Validate zip contents against order key
            logger.info(f"Validating {os.path.basename(order_folder)}")
            validation_result = processor.validate_zip_contents(
                zip_path, i_number, order_number, order_key
            )

            # Add validation results to worksheet
            if validation_result:
                order_count += 1

                # Use the add_validation_result method instead of direct worksheet access
                is_andreev = getattr(config, 'ANDREEV_NAME', 'andreev') in os.path.basename(order_folder).lower()
                row_count = excel_dao.add_validation_result(
                    worksheet, row_count, validation_result, zip_path,
                    i_number, order_number, is_andreev
                )

    # If this is an update to existing workbook, add a break
    if summary_exists and order_count > 0:
        excel_dao.set_cell_value(worksheet, row_count, 1, "Break")
        excel_dao.apply_style(worksheet, f'A{row_count}', 'break')

    # Auto-adjust column widths
    if order_count > 0:
        excel_dao.adjust_column_widths(worksheet)

    # Save workbook
    try:
        excel_dao.save_workbook(workbook, excel_path)
        logger.info(f"Summary saved to {excel_filename}")
    except Exception as e:
        logger.error(f"Error saving workbook: {e}")
        print(f"Error: Could not save summary. Make sure the file is not open in Excel.")

    logger.info(f"Total orders validated: {order_count}")
    print(f"All done! {order_count} orders validated.")


if __name__ == "__main__":
    main()