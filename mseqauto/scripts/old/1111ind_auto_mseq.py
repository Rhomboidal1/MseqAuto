# ind_auto_mseq.py
import tkinter as tk
from tkinter import filedialog
import subprocess
import re
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

print(sys.path)

def get_folder_from_user():
    """Simple folder selection dialog that works reliably"""
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force an update to ensure dialog shows

    folder_path = filedialog.askdirectory(
        title="Select today's data folder to mseq orders",
        mustexist=True
    )

    root.destroy()
    return folder_path

def main():
    # Get folder path FIRST before any package imports
    data_folder = get_folder_from_user()

    if not data_folder:
        print("No folder selected, exiting")
        return

    # NOW import package modules
    from mseqauto.utils import setup_logger
    from mseqauto.config import MseqConfig
    from mseqauto.core import OSCompatibilityManager, FileSystemDAO, MseqAutomation, FolderProcessor
    import logging

    # Diagnostic logging information
    print("ROOT LOGGER LEVEL:", logging.getLogger().level)
    print("ROOT LOGGER HANDLERS:", logging.getLogger().handlers)

    # After setup_logger
    logger = setup_logger("ind_auto_mseq")
    logger.info("Starting IND auto mSeq...")
    print("Data folder selected:", data_folder)
    print("Starting IND auto mSeq...")

    # Configure ui_automation logger
    ui_automation_logger = logging.getLogger("mseqauto.core.ui_automation")
    ui_automation_logger.setLevel(logging.DEBUG)
    ui_automation_logger.propagate = True  # Ensure logs propagate to parent
    print("ui_automation LOGGER LEVEL:", ui_automation_logger.level)
    print("ui_automation LOGGER HANDLERS:", ui_automation_logger.handlers)    

    # More diagnostics
    print("ind_auto_mseq LOGGER LEVEL:", logger.level)
    print("ind_auto_mseq LOGGER HANDLERS:", logger.handlers)


    # Log that we already selected folder
    logger.info(f"Using folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)

    # Check for 32-bit Python requirement
    OSCompatibilityManager.py32_check(
        script_path=__file__,
        logger=logger
    )

    # Log OS environment INFOrmation
    OSCompatibilityManager.log_environment_info(logger)

    # Initialize components
    logger.info("Initializing components...")
    config = MseqConfig()
    logger.info("Config loaded")

    # Use OS compatibility manager for timeouts
    logger.info("Using OS-specific timeouts")

    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")

    ui_automation = MseqAutomation(config, use_fast_navigation=True, log_level=logging.DEBUG)
    logger.info("UI Automation initialized with fast navigation")

    processor = FolderProcessor(file_dao, ui_automation, config, logger=logger.info)
    logger.info("Folder processor initialized")

    # # Run batch file to generate order key
    # try:
    #     logger.info(f"Running batch file: {config.BATCH_FILE_PATH}")
    #     subprocess.run(config.BATCH_FILE_PATH, shell=True, check=True)
    #     logger.info("Batch file completed successfully")
    # except subprocess.CalledProcessError:
    #     logger.error(f"Batch file {config.BATCH_FILE_PATH} failed to run")
    #     print(f"Error: Batch file {config.BATCH_FILE_PATH} failed to run")
    #     return

    try:
        # Process BioI folders
        bio_folders = file_dao.get_folders(data_folder, r'bioi-\d+')
        logger.info(f"Found {len(bio_folders)} BioI folders")
        print(f"Found {len(bio_folders)} BioI folders")

        immediate_orders = file_dao.get_folders(data_folder, r'bioi-\d+_.+_\d+')
        logger.info(f"Found {len(immediate_orders)} immediate order folders")
        print(f"Found {len(immediate_orders)} immediate order folders")

        pcr_folders = file_dao.get_folders(data_folder, r'fb-pcr\d+_\d+')
        logger.info(f"Found {len(pcr_folders)} PCR folders")
        print(f"Found {len(pcr_folders)} PCR folders")

        # Process BioI folders
        for i, folder in enumerate(bio_folders):
            logger.info(f"Processing BioI folder {i + 1}/{len(bio_folders)}: {os.path.basename(folder)}")
            print(f"Processing BioI folder {i + 1}/{len(bio_folders)}: {os.path.basename(folder)}")
            processor.process_bio_folder(folder)

        # Determine if we're processing the IND Not Ready folder
        is_ind_not_ready = os.path.basename(data_folder) == config.IND_NOT_READY_FOLDER
        logger.info(f"Is IND Not Ready folder: {is_ind_not_ready}")
        print(f"Is IND Not Ready folder: {is_ind_not_ready}")
        if is_ind_not_ready:
            logger.info("Processing IND Not Ready folder")
            print("#Processing IND Not Ready folder")
            # Process the IND Not Ready folder if applicable
            processor.process_ind_not_ready_folder(data_folder)
        else:
            logger.info("Not processing IND Not Ready folder, continuing with immediate orders")

        # Process immediate orders
        for i, folder in enumerate(immediate_orders):
            logger.info(f"Processing order folder {i + 1}/{len(immediate_orders)}: {os.path.basename(folder)}")
            print(f"#Processing order folder {i + 1}/{len(immediate_orders)}: {os.path.basename(folder)}")
            processor.process_order_folder(folder, data_folder)

        # Process PCR folders
        for i, folder in enumerate(pcr_folders):
            logger.info(f"Processing PCR folder {i + 1}/{len(pcr_folders)}: {os.path.basename(folder)}")
            print(f"#Processing PCR folder {i + 1}/{len(pcr_folders)}: {os.path.basename(folder)}")
            processor.process_pcr_folder(folder)

        logger.info("All processing completed")
        print("")
        print("ALL DONE")
    except Exception as e:
        import traceback
        logger.error(f"Unexpected error: {e}")
        logger.error(traceback.format_exc())
        print(f"Unexpected error: {e}")
    finally:
        # Close mSeq application
        try:
            if ui_automation is not None:
                logger.info("Closing mSeq application")
                ui_automation.close()
        except Exception as e:
            logger.error(f"Error closing mSeq: {e}")
            print(f"Error closing mSeq: {e}")

if __name__ == "__main__":
    main()