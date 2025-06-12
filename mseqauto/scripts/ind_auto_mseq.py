# ind_auto_mseq.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import re
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
import subprocess
import ctypes

import warnings
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

# Configuration settings
use_confirmation = False  # Set to False to skip the confirmation prompt

# Check if running under GUI
GUI_MODE = os.environ.get('MSEQAUTO_GUI_MODE', 'False') == 'True'

def get_folder_from_user():
    # Check for GUI-provided folder first
    if GUI_MODE and 'MSEQAUTO_DATA_FOLDER' in os.environ:
        return os.environ['MSEQAUTO_DATA_FOLDER']
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force update to ensure dialog shows
    
    folder_path = filedialog.askdirectory(
        title="Select today's data folder for IND mSeq",
        mustexist=True
    )
    
    root.destroy()
    return folder_path

def check_relaunch_status():
    """Check if we're after a relaunch"""
    return os.environ.get('MSEQ_RELAUNCHED', '') == 'True'

def main(provided_folder=None):
    """Main function with improved 32-bit check logic"""
    # Check if we're in a relaunched process
    is_relaunched = check_relaunch_status()
    
    # SKIP the 32-bit check entirely when running under GUI
    # The GUI is already launching us with 32-bit Python
    if not GUI_MODE:
        # Only do the 32-bit check in standalone mode
        if not is_relaunched and provided_folder is None and sys.maxsize > 2 ** 32:
            print("Detected 64-bit Python, need to relaunch in 32-bit for mSeq automation...")
            
            # Set environment variable before relaunching
            os.environ['MSEQ_RELAUNCHED'] = 'True'
            
            # Import and use the relaunch logic
            from mseqauto.core import OSCompatibilityManager  # type: ignore
            
            # This will relaunch in 32-bit Python and exit current process
            OSCompatibilityManager.py32_check(
                script_path=__file__,
                logger=None  # No logger yet since we haven't selected folder
            )
            return  # Should not reach here as py32_check exits
    
    # If we get here, we're either:
    # 1. Already in 32-bit Python from the start, OR 
    # 2. In the relaunched 32-bit process, OR
    # 3. Using a provided folder (from combined script)
    # 4. Running under GUI (which handles 32-bit Python for us)
    
    if is_relaunched:
        print("Running in relaunched 32-bit Python process")
    elif provided_folder is None:
        print("Already running in 32-bit Python")
    
    # Get folder path (only once, in the correct Python version)
    if provided_folder:
        data_folder = provided_folder
        print(f"Using provided folder: {data_folder}")
    else:
        data_folder = get_folder_from_user()
    
    if not data_folder:
        print("No folder selected, exiting")
        return
    
    # NOW import package modules (after folder selection and 32-bit check)
    from mseqauto.utils import setup_logger  # type: ignore
    from mseqauto.config import MseqConfig  # type: ignore
    from mseqauto.core import OSCompatibilityManager, FileSystemDAO, MseqAutomation, FolderProcessor  # type: ignore
    
    # Get the script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "logs")

    # Setup logger
    logger = setup_logger("ind_auto_mseq", log_dir=log_dir)
    logger.info("Starting IND auto mSeq...")
    logger.info(f"Using folder: {data_folder}")
    data_folder = os.path.normpath(data_folder)
    
    # Log OS information
    OSCompatibilityManager.log_environment_info(logger)
    
    # Initialize components
    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    ui_automation = MseqAutomation(config)
    processor = FolderProcessor(file_dao, ui_automation, config, logger=logger.info)
    
    try:
        # Get folders to process
        bio_folders = file_dao.get_folders(data_folder, r'bioi-\d+')
        logger.info(f"Found {len(bio_folders)} BioI folders")
        
        immediate_orders = file_dao.get_folders(data_folder, r'bioi-\d+_.+_\d+')
        logger.info(f"Found {len(immediate_orders)} immediate order folders")
        
        pcr_folders = file_dao.get_folders(data_folder, r'fb-pcr\d+_.*')
        logger.info(f"Found {len(pcr_folders)} PCR folders")
        
        # List what will be processed
        if bio_folders:  # Overarching plate folders containing individual orders
            print("\nBioI folders to process:")
            for i, folder in enumerate(bio_folders):
                print(f"{i+1}. {os.path.basename(folder)}")
                
        if immediate_orders:  # Individual customer order folders
            print("\nImmediate order folders to process:")
            for i, folder in enumerate(immediate_orders):
                print(f"{i+1}. {os.path.basename(folder)}")
                
        if pcr_folders:
            print("\nPCR folders to process:")
            for i, folder in enumerate(pcr_folders):
                print(f"{i+1}. {os.path.basename(folder)}")
        
        # Ask for confirmation if configured to do so
        proceed = True
        if use_confirmation and (bio_folders or immediate_orders or pcr_folders):
            confirmation = input("\nProceed with processing these folders? (y/n): ")
            proceed = confirmation.lower() == 'y'
            
        if not proceed:
            print("Cancelled by user")
            return
            
        # If no folders found at all
        if not bio_folders and not immediate_orders and not pcr_folders:
            print("No folders found to process, exiting")
            return
            
        # Process BioI folders
        for i, folder in enumerate(bio_folders):
            logger.info(f"Processing BioI folder {i+1}/{len(bio_folders)}: {os.path.basename(folder)}")
            processor.process_bio_folder(folder)
        
        # Check if processing IND Not Ready folder
        is_ind_not_ready = os.path.basename(data_folder) == config.IND_NOT_READY_FOLDER
        logger.info(f"Is IND Not Ready folder: {is_ind_not_ready}")
        
        # Process immediate orders
        for i, folder in enumerate(immediate_orders):
            logger.info(f"Processing order folder {i+1}/{len(immediate_orders)}: {os.path.basename(folder)}")
            processor.process_order_folder(folder, data_folder)
        
        # Process PCR folders
        for i, folder in enumerate(pcr_folders):
            logger.info(f"Processing PCR folder {i+1}/{len(pcr_folders)}: {os.path.basename(folder)}")
            processor.process_pcr_folder(folder)
        
        logger.info("All processing completed")
        print("\nALL DONE")
        
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

if __name__ == "__main__":
    main()