# plate_auto_mseq.py
import tkinter as tk
from tkinter import filedialog
import subprocess
import re
import time
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
        title="Select today's data folder for Plate mSeq",
        mustexist=True
    )
    
    root.destroy()
    return folder_path
    

def check_relaunch_status():
    """Check if we're after a relaunch"""
    return os.environ.get('MSEQ_RELAUNCHED', '') == 'True'

def cleanup_temp_files():
    """Clean up temporary files after successful run"""
    try:
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
        temp_file = os.path.join(temp_dir, "selected_folder.txt")
        if os.path.exists(temp_file):
            os.remove(temp_file)
            print("Cleaned up temporary folder selection file")
    except Exception as e:
        print(f"Error cleaning up temp files: {e}")

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
    logger = setup_logger("plate_auto_mseq", log_dir=log_dir)
    logger.info("Starting Plate auto mSeq...")
    logger.info(f"Using folder: {data_folder}")
    data_folder = os.path.normpath(data_folder)
    
    # Log OS information
    OSCompatibilityManager.log_environment_info(logger)
    
    # Initialize components
    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    ui_automation = MseqAutomation(config)
    processor = FolderProcessor(file_dao, ui_automation, config, logger=logger.info)

    
    # Get all plate folders (starting with 'p')
    plate_folders = file_dao.get_folders(data_folder, r'p\d+.+')
    print(f"Found {len(plate_folders)} plate folders")
    
    if len(plate_folders) == 0:
        print("No plate folders found to process, exiting")
        cleanup_temp_files()  # Clean up temp files before exiting
        return
    
    # List folders that will be processed
    print("Plate folders to process:")
    for i, folder in enumerate(plate_folders):
        print(f"{i+1}. {os.path.basename(folder)}")
    
    # Ask for confirmation only if use_confirmation is True
    proceed = True
    if use_confirmation:
        confirmation = input("Proceed with processing these plate folders? (y/n): ")
        proceed = confirmation.lower() == 'y'
    
    if not proceed:
        print("Cancelled by user")
        cleanup_temp_files()  # Clean up temp files before exiting
        return
    
    try:
        # Process each plate folder
        for i, folder in enumerate(plate_folders):
            print(f"Processing plate folder {i+1}/{len(plate_folders)}: {os.path.basename(folder)}")
            
            # Check if this is an FSA plate which should be skipped
            if file_dao.contains_file_type(folder, config.FSA_EXTENSION):
                print(f"Skipping FSA plate: {os.path.basename(folder)}")
                continue
                
            # Check if folder was already processed
            was_mseqed, has_braces, has_ab1_files = processor.check_order_status(folder)
            
            if was_mseqed:
                print(f"Folder already processed: {os.path.basename(folder)}")
                continue
                
            if not has_ab1_files:
                print(f"No AB1 files found in: {os.path.basename(folder)}")
                continue
                
            # Skip folders with files that have braces - uncomment if needed
            # Plate orders typically don't get reinjects, but braces might indicate special handling
            # Note: Some customers might include braces in their sample names, which would be filtered unintentionally
            # if has_braces:
            #     print(f"Folder contains files with braces that may need cleaning: {os.path.basename(folder)}")
            #     continue
                
            # Process the folder directly with ui_automation
            ui_automation.process_folder(folder)
            logger.info(f"Processed plate folder: {os.path.basename(folder)}")
            
            # Add a short delay between folder processing
            if i < len(plate_folders) - 1:
                time.sleep(1)
        
        print("All done!")
    except Exception as e:
        import traceback
        logger.error(f"Error during processing: {e}")
        logger.error(traceback.format_exc())
        print(f"Error occurred: {e}")
    finally:
        # Always close mSeq application and clean up temp files
        print("Closing mSeq application...")
        try:
            ui_automation.close()
        except Exception as e:
            print(f"Error closing mSeq: {e}")
        
        # Clean up temp files at the end of a successful run
        if not is_relaunched:  # Only clean up if we're in the main process
            cleanup_temp_files()

if __name__ == "__main__":
    main()