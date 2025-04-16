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

# Store the selected folder path globally to avoid asking twice
selected_folder = None

def get_folder_from_user():
    """Get folder from user with a simple dialog"""
    global selected_folder
    
    # First check if we're relaunching and have folder path in temp file
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
    temp_file = os.path.join(temp_dir, "selected_folder.txt")
    
    if os.path.exists(temp_file):
        try:
            with open(temp_file, "r") as f:
                stored_path = f.read().strip()
                if stored_path:
                    selected_folder = stored_path
                    print(f"Using stored folder path: {selected_folder}")
                    return selected_folder
        except Exception as e:
            print(f"Error reading stored folder path: {e}")
    
    # If folder was already selected (in case of relaunch), use it
    if selected_folder:
        return selected_folder
        
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force update to ensure dialog shows
    
    folder_path = filedialog.askdirectory(
        title="Select today's data folder to mseq orders",
        mustexist=True
    )
    
    root.destroy()
    
    # Store the selection globally
    selected_folder = folder_path
    
    # Store the path in a temporary file for potential relaunch
    if folder_path:
        os.makedirs(temp_dir, exist_ok=True)
        with open(temp_file, "w") as f:
            f.write(folder_path)
            
    return folder_path

def check_and_cleanup_relaunch():
    """Check if we're after a relaunch and clean up temp files if needed"""
    # If this is the 32-bit Python process after relaunch, we'll know by environment variable
    is_relaunched = os.environ.get('MSEQ_RELAUNCHED', '') == 'True'
    
    if is_relaunched:
        print("Running in relaunched 32-bit Python process")
        return True
    return False

def main():
    # Check if we're in a relaunched process
    is_relaunched = check_and_cleanup_relaunch()

    # Get folder path FIRST before any imports
    data_folder = get_folder_from_user()
    
    if not data_folder:
        print("No folder selected, exiting")
        return
    
    # NOW import package modules
    from mseqauto.utils import setup_logger # type: ignore
    from mseqauto.config import MseqConfig # type: ignore
    from mseqauto.core import OSCompatibilityManager, FileSystemDAO, MseqAutomation, FolderProcessor # type: ignore
    
    # Get the script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "logs")

    # Setup logger
    logger = setup_logger("ind_auto_mseq", log_dir=log_dir)
    logger.info("Starting IND auto mSeq...")
    logger.info(f"Using folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)
    
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
    
    try:
        # Get folders to process
        bio_folders = file_dao.get_folders(data_folder, r'bioi-\d+')
        logger.info(f"Found {len(bio_folders)} BioI folders")
        
        immediate_orders = file_dao.get_folders(data_folder, r'bioi-\d+_.+_\d+')
        logger.info(f"Found {len(immediate_orders)} immediate order folders")
        
        pcr_folders = file_dao.get_folders(data_folder, r'fb-pcr\d+_\d+')
        logger.info(f"Found {len(pcr_folders)} PCR folders")
        
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
        # # Get folders to process
        # bio_folders = file_dao.get_folders(data_folder, r'bioi-\d+')
        # logger.info(f"Found {len(bio_folders)} BioI folders")
        
        # immediate_orders = file_dao.get_folders(data_folder, r'bioi-\d+_.+_\d+')
        # logger.info(f"Found {len(immediate_orders)} immediate order folders")
        
        # pcr_folders = file_dao.get_folders(data_folder, r'fb-pcr\d+_\d+')
        # logger.info(f"Found {len(pcr_folders)} PCR folders")
        
        # # Process BioI folders
        # for i, folder in enumerate(bio_folders):
        #     logger.info(f"Processing BioI folder {i+1}/{len(bio_folders)}: {os.path.basename(folder)}")
        #     processor.process_bio_folder(folder)
        
        # # Check if processing IND Not Ready folder
        # is_ind_not_ready = os.path.basename(data_folder) == config.IND_NOT_READY_FOLDER
        # logger.info(f"Is IND Not Ready folder: {is_ind_not_ready}")
        
        # # Process immediate orders
        # for i, folder in enumerate(immediate_orders):
        #     logger.info(f"Processing order folder {i+1}/{len(immediate_orders)}: {os.path.basename(folder)}")
        #     processor.process_order_folder(folder, data_folder)
        
        # # Process PCR folders
        # for i, folder in enumerate(pcr_folders):
        #     logger.info(f"Processing PCR folder {i+1}/{len(pcr_folders)}: {os.path.basename(folder)}")
        #     processor.process_pcr_folder(folder)
        
        # logger.info("All processing completed")
        # print("\nALL DONE")
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
        # Clean up temp file after successful run
        try:
            temp_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp", "selected_folder.txt")
            if os.path.exists(temp_file):
                os.remove(temp_file)
                print("Cleaned up temporary folder selection file")
        except Exception as e:
            print(f"Error cleaning up temp file: {e}")

if __name__ == "__main__":
    main()