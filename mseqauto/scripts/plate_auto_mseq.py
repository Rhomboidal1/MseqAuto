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

def get_folder_from_user(is_relaunched=False):
    """
    Get folder from user with a simple dialog
    
    Args:
        is_relaunched: True if this is running after a Python relaunch
    """
    # Only check for stored path if we're in a relaunched session
    if is_relaunched:
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
        temp_file = os.path.join(temp_dir, "selected_folder.txt")
        
        if os.path.exists(temp_file):
            try:
                with open(temp_file, "r") as f:
                    stored_path = f.read().strip()
                    if stored_path and os.path.exists(stored_path):
                        print(f"Using stored folder path: {stored_path}")
                        return stored_path
            except Exception as e:
                print(f"Error reading stored folder path: {e}")
    
    # Always prompt for folder selection on new runs
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force update to ensure dialog shows
    
    folder_path = filedialog.askdirectory(
        title="Select today's data folder for plate mSeq",
        mustexist=True
    )
    
    root.destroy()
    
    # Store the path in a temporary file for potential relaunch
    if folder_path:
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
        os.makedirs(temp_dir, exist_ok=True)
        temp_file = os.path.join(temp_dir, "selected_folder.txt")
        with open(temp_file, "w") as f:
            f.write(folder_path)
            
    return folder_path

def check_relaunch_status():
    """Check if we're after a relaunch and clean up temp files if needed"""
    # If this is the 32-bit Python process after relaunch, we'll know by environment variable
    is_relaunched = os.environ.get('MSEQ_RELAUNCHED', '') == 'True'
    
    if is_relaunched:
        print("Running in relaunched 32-bit Python process")
    
    return is_relaunched

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

def main():
    # Check if we're in a relaunched process
    is_relaunched = check_relaunch_status()

    # Get folder path FIRST before any imports, passing relaunch status
    data_folder = get_folder_from_user(is_relaunched)
    
    if not data_folder:
        print("No folder selected, exiting")
        return

    # NOW import package modules
    from mseqauto.utils import setup_logger #type: ignore
    from mseqauto.config import MseqConfig #type: ignore
    from mseqauto.core import OSCompatibilityManager, FileSystemDAO, MseqAutomation, FolderProcessor #type: ignore

    # Get the script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "logs")

    # Setup logger
    logger = setup_logger("plate_auto_mseq", log_dir=log_dir)
    logger.info("Starting Plate auto mSeq...")
    logger.info(f"Using folder: {data_folder}")
    data_folder = os.path.normpath(data_folder)
    
    # Only check for 32-bit Python if we haven't already relaunched
    if not is_relaunched:
        # Set environment variable before relaunching
        os.environ['MSEQ_RELAUNCHED'] = 'True'
        
        # Check for 32-bit Python requirement and potentially relaunch
        # This will exit the current process if it relaunches
        OSCompatibilityManager.py32_check(
            script_path=__file__,
            logger=logger
        )
    
    
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