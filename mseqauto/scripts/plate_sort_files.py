# plate_sort_files.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
import re
from pathlib import Path
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(str(Path(__file__).parents[2]))

def get_folder_from_user():
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Add this line - force an update

    folder_path = filedialog.askdirectory(
        title="Select today's data folder to sort",
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
    script_dir = Path(__file__).parent.resolve()
    log_dir = script_dir / "logs"

    # Setup logger
    logger = setup_logger("plate_sort_files", log_dir=log_dir)

    logger.info("Starting Plate sort files...")

    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")
    processor = FolderProcessor(file_dao, None, config, logger=logger.info)
    logger.info("Folder processor initialized")
    logger.info(f"Using folder: {data_folder}")



    # Get plate folders
    plate_folders = file_dao.get_folders(data_folder, pattern=r'p\d+.+')
    logger.info(f"Found {len(plate_folders)} plate folders")

    if not plate_folders:
        logger.warning("No plate folders found, exiting")
        print("No plate folders found, exiting")
        return

    # Process each plate folder
    for i, folder in enumerate(plate_folders):
        logger.info(f"Processing plate folder {i+1}/{len(plate_folders)}")
        processor.sort_plate_folder(folder)


    logger.info("All plate folders processed")
    print("All done!")

if __name__ == "__main__":
    main()