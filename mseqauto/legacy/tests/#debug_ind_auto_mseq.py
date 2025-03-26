# progressive_ind_auto_mseq.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
import re
from mseqauto.utils import logger, setup_logger

# Only import the necessary components - we'll add more later
from mseqauto.config import MseqConfig
from mseqauto.core import FileSystemDAO


# We'll add the other imports later

def get_folder_from_user():
    """Get folder selection from user using a simple approach"""
    print("Opening folder selection dialog...")

    root = tk.Tk()
    root.withdraw()

    folder_path = filedialog.askdirectory(
        title="Select today's data folder to mseq orders",
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
    logger = setup_logger("progressive_ind_auto_mseq")
    logger.info("Starting progressive IND auto mSeq...")

    # Get folder selection FIRST
    logger.info("About to select folder...")
    data_folder = get_folder_from_user()

    if not data_folder:
        logger.error("No folder selected, exiting")
        print("No folder selected, exiting")
        return

    logger.info(f"Using folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)

    # Initialize components one by one
    logger.info("Initializing components...")
    config = MseqConfig()
    logger.info("Config loaded")

    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")

    # Now we'll add the UI automation
    logger.info("About to import UI automation...")
    from mseqauto.core import MseqAutomation
    logger.info("UI automation imported")

    ui_automation = MseqAutomation(config)
    logger.info("UI Automation initialized")

    # Finally add the folder processor
    logger.info("About to import Folder Processor...")
    from mseqauto.core import FolderProcessor
    logger.info("Folder processor imported")

    processor = FolderProcessor(file_dao, ui_automation, config, logger=logger.info)
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

    logger.info("All initialization completed")
    print("All components initialized successfully")

    # We won't actually process any folders for this test
    logger.info("Test successful")


if __name__ == "__main__":
    main()