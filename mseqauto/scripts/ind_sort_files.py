# ind_sort_files.py
import os
import subprocess
import tkinter as tk
import sys
from datetime import datetime
from tkinter import filedialog
import warnings
from pathlib import Path

sys.path.append(str(Path(__file__).parents[2]))
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

# Check if running under GUI
GUI_MODE = os.getenv('MSEQAUTO_GUI_MODE', 'False') == 'True'

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
            title="Select today's data folder to sort",
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
    script_dir = Path(__file__).parent.resolve()
    log_dir = script_dir / "logs"

    # Setup logger
    logger = setup_logger("ind_sort_files", log_dir=log_dir)

    logger.info("Starting IND sort files...")

    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")
    processor = FolderProcessor(file_dao, None, config, logger=logger.info)
    logger.info("Folder processor initialized")
    logger.info(f"Using folder: {data_folder}")

    # Run batch file to generate order key
    try:
        logger.info(f"Running batch file: {config.BATCH_FILE_PATH}")
        batch_file_path = str(config.BATCH_FILE_PATH)  # Convert Path to string
        subprocess.run(batch_file_path, shell=True, check=True)
        logger.info("Batch file completed successfully")
    except subprocess.CalledProcessError:
        logger.error(f"Batch file {config.BATCH_FILE_PATH} failed to run")
        print(f"Error: Batch file {config.BATCH_FILE_PATH} failed to run")
        return

    # Store the selected folder in the processor for later reference
    processor.current_data_folder = data_folder #type: ignore

    # Get today's I numbers and BioI folders
    i_numbers, bio_folders = file_dao.get_folders_with_inumbers(data_folder)
    logger.info(f"Found {len(i_numbers)} I numbers and {len(bio_folders)} BioI folders")

    # Get recent I numbers for order key filtering
    recent_inumbers = file_dao.collect_active_inumbers(
        paths=['G:\\Lab\\Spreadsheets\\Individual Uploaded to ABI', 'G:\\Lab\\Spreadsheets'],
        min_inum=file_dao.get_most_recent_inumber('P:\\Data\\Individuals')
    )

    # Load order key and adjust characters
    order_key = file_dao.load_order_key(str(config.KEY_FILE_PATH))
    logger.info("Order key loaded")

    # Get complete list of reinjects
    reinject_path = f"P:\\Data\\Reinjects\\Reinject List_{datetime.now().strftime('%m-%d-%Y')}.xlsx"
    try:
        reinject_list = processor.get_reinject_list(recent_inumbers, reinject_path)
        logger.info(f"Found {len(reinject_list)} reinjects")
    except Exception as e:
        logger.error(f"Error loading reinject list: {e}")
        reinject_list = []

    # Process each BioI folder
    for i, folder in enumerate(bio_folders):
        logger.info(f"Processing folder {i+1}/{len(bio_folders)}: {Path(folder).name}")
        processor.sort_ind_folder(folder, reinject_list, order_key)

    # Final cleanup pass for the entire data folder
    processor.final_cleanup(data_folder)

    logger.info("All folders processed")
    print("All done!")

if __name__ == "__main__":
    main()