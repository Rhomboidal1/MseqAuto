# ind_process_all.py
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
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
        title="Select today's data folder for complete IND processing",
        mustexist=True
    )
    
    root.destroy()
    return folder_path

def show_error_and_exit(step_name, error_msg):
    """Show error popup and exit"""
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showerror(
        "Processing Error", 
        f"An error occurred during {step_name}.\n\n{error_msg}\n\nNow exiting."
    )
    
    root.destroy()
    sys.exit(1)

def run_sort_files(data_folder):
    """Run the file sorting step"""
    try:
        from mseqauto.config import MseqConfig # type: ignore
        from mseqauto.core import FileSystemDAO, FolderProcessor # type: ignore
        from mseqauto.utils import setup_logger # type: ignore
        from datetime import datetime
        import subprocess
        
        # Get the script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(script_dir, "logs")

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
            subprocess.run(config.BATCH_FILE_PATH, shell=True, check=True)
            logger.info("Batch file completed successfully")
        except subprocess.CalledProcessError:
            logger.error(f"Batch file {config.BATCH_FILE_PATH} failed to run")
            raise Exception(f"Batch file {config.BATCH_FILE_PATH} failed to run")
        
        # Store the selected folder in the processor for later reference
        processor.current_data_folder = data_folder

        # Get today's I numbers and BioI folders
        i_numbers, bio_folders = file_dao.get_folders_with_inumbers(data_folder)
        logger.info(f"Found {len(i_numbers)} I numbers and {len(bio_folders)} BioI folders")
        
        # Get recent I numbers for order key filtering
        recent_inumbers = file_dao.collect_active_inumbers(
            paths=['G:\\Lab\\Spreadsheets\\Individual Uploaded to ABI', 'G:\\Lab\\Spreadsheets'],
            min_inum=file_dao.get_most_recent_inumber('P:\\Data\\Individuals')
        )
        
        # Load order key and adjust characters
        order_key = file_dao.load_order_key(config.KEY_FILE_PATH)
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
            logger.info(f"Processing folder {i+1}/{len(bio_folders)}: {os.path.basename(folder)}")
            processor.sort_ind_folder(folder, reinject_list, order_key)
        
        # Final cleanup pass for the entire data folder
        processor.final_cleanup(data_folder)

        logger.info("All folders processed")
        return True
        
    except Exception as e:
        raise Exception(f"File sorting failed: {str(e)}")

def run_mseq_processing(data_folder):
    """Run the mSeq processing step"""
    try:
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
        data_folder = os.path.normpath(data_folder)
        
        # Skip the 32-bit Python check in the combined script context
        # The individual script handles this, but we don't want it in the combined version
        # OSCompatibilityManager.py32_check(script_path=__file__, logger=logger)
        
        # Log OS information
        OSCompatibilityManager.log_environment_info(logger)
        
        # Initialize components
        config = MseqConfig()
        file_dao = FileSystemDAO(config)
        ui_automation = MseqAutomation(config, logger)
        processor = FolderProcessor(file_dao, ui_automation, config, logger=logger.info)
        
        try:
            # Get folders to process
            bio_folders = file_dao.get_folders(data_folder, r'bioi-\d+')
            logger.info(f"Found {len(bio_folders)} BioI folders")
            
            immediate_orders = file_dao.get_folders(data_folder, r'bioi-\d+_.+_\d+')
            logger.info(f"Found {len(immediate_orders)} immediate order folders")
            
            pcr_folders = file_dao.get_folders(data_folder, r'fb-pcr\d+_.*')
            logger.info(f"Found {len(pcr_folders)} PCR folders")
            
            # If no folders found at all
            if not bio_folders and not immediate_orders and not pcr_folders:
                logger.info("No folders found to process")
                return True
                
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
            return True
            
        finally:
            # Close mSeq application
            try:
                if ui_automation is not None:
                    logger.info("Closing mSeq application")
                    ui_automation.close()
            except Exception as e:
                logger.error(f"Error closing mSeq: {e}")
                
    except Exception as e:
        raise Exception(f"mSeq processing failed: {str(e)}")

def run_zip_files(data_folder):
    """Run the file zipping step"""
    try:
        from mseqauto.config import MseqConfig # type: ignore
        from mseqauto.core import FileSystemDAO, FolderProcessor # type: ignore
        from mseqauto.utils import setup_logger # type: ignore
        
        # Get the script directory
        script_dir = os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(script_dir, "logs")

        # Setup logger
        logger = setup_logger("ind_zip_files", log_dir=log_dir)
        logger.info("Starting IND zip files...")
        
        # Initialize components
        config = MseqConfig()
        logger.info("Config loaded")
        REGEX = config.REGEX_PATTERNS
        file_dao = FileSystemDAO(config, logger=logger)
        logger.info("FileSystemDAO initialized")
        processor = FolderProcessor(file_dao, None, config, logger=logger.info)
        logger.info("Folder processor initialized")
        logger.info(f"Using folder: {data_folder}")
        
        # Create zip dump folder
        zip_dump_folder = os.path.join(data_folder, config.ZIP_DUMP_FOLDER)
        if not os.path.exists(zip_dump_folder):
            os.makedirs(zip_dump_folder)
            logger.info(f"Created zip dump folder: {zip_dump_folder}")

        bio_folders = file_dao.get_folders(data_folder, pattern=REGEX['bioi_folder'].pattern)
        logger.info(f"Found {len(bio_folders)} BioI folders")

        # Get PCR folders - use the new pcr_folder pattern
        pcr_folders = file_dao.get_folders(data_folder, pattern=REGEX['pcr_folder'].pattern)
        logger.info(f"Found {len(pcr_folders)} PCR folders")

        # Copy recent existing zips to the zip dump folder (recovery logic)
        recovered_count = file_dao.copy_recent_zips_to_dump(
            bio_folders + pcr_folders, 
            zip_dump_folder, 
            max_age_minutes=15
        )

        if recovered_count > 0:
            logger.info(f"Recovered {recovered_count} recently created zip files")

        # Process BioI folders
        order_count = 0
        for bio_folder in bio_folders:
            # Get order folders
            order_folders = file_dao.get_folders(bio_folder, pattern=REGEX['order_folder'].pattern)

            # Filter out reinject folders
            order_folders = [folder for folder in order_folders 
                            if not REGEX['reinject'].search(os.path.basename(folder).lower())]

            logger.info(f"Found {len(order_folders)} order folders in {os.path.basename(bio_folder)}")
            
            for order_folder in order_folders:
                # Check if order already has a zip file
                if file_dao.check_for_zip(order_folder):
                    logger.info(f"Skipping {os.path.basename(order_folder)} - already has zip file")
                    continue
                
                # Zip the order folder
                logger.info(f"Zipping {os.path.basename(order_folder)}")
                zip_path = processor.zip_order_folder(order_folder, include_txt=True)
                
                if zip_path:
                    # Copy zip to dump folder
                    logger.info(f"Copying zip to dump folder: {os.path.basename(zip_path)}")
                    file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
                    
                    order_count += 1
                    logger.info(f"Successfully processed {os.path.basename(order_folder)}")
                else:
                    logger.warning(f"Failed to zip {os.path.basename(order_folder)}")
        
        # Process PCR folders
        for pcr_folder in pcr_folders:
            # Check if PCR folder already has a zip file
            if file_dao.check_for_zip(pcr_folder):
                logger.info(f"Skipping {os.path.basename(pcr_folder)} - already has zip file")
                continue
            
            # Zip the PCR folder
            logger.info(f"Zipping {os.path.basename(pcr_folder)}")
            zip_path = processor.zip_order_folder(pcr_folder, include_txt=True)
            
            if zip_path:
                # Copy zip to dump folder
                logger.info(f"Copying zip to dump folder: {os.path.basename(zip_path)}")
                file_dao.copy_zip_to_dump(zip_path, zip_dump_folder)
                
                order_count += 1
                logger.info(f"Successfully processed {os.path.basename(pcr_folder)}")
            else:
                logger.warning(f"Failed to zip {os.path.basename(pcr_folder)}")
        
        # Calculate total processed files
        total_processed = order_count + recovered_count
        
        # Remove empty zip dump folder if nothing was processed
        if total_processed == 0 and os.path.exists(zip_dump_folder) and not os.listdir(zip_dump_folder):
            os.rmdir(zip_dump_folder)
            logger.info("Removed empty zip dump folder")
        
        # Log the total count
        logger.info(f"Total files processed: {total_processed} (New: {order_count}, Recovered: {recovered_count})")
        
        return True
        
    except Exception as e:
        raise Exception(f"File zipping failed: {str(e)}")

def check_relaunch_status():
    """Check if we're after a relaunch"""
    return os.environ.get('MSEQ_COMBINED_RELAUNCHED', '') == 'True'

def main():
    """Main function to run all IND processing steps in sequence"""
    # Check if we're in a relaunched process
    is_relaunched = check_relaunch_status()
    
    # Import just what we need for the 32-bit check (minimal imports to avoid dialog issues)
    import sys
    
    # Check if we're in 32-bit Python FIRST, before folder selection
    if not is_relaunched and sys.maxsize > 2 ** 32:
        print("Detected 64-bit Python, need to relaunch in 32-bit for mSeq automation...")
        
        # Set environment variable before relaunching
        os.environ['MSEQ_COMBINED_RELAUNCHED'] = 'True'
        
        # Import and use the relaunch logic
        sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
        from mseqauto.core import OSCompatibilityManager  # type: ignore
        
        # This will relaunch in 32-bit Python and exit current process
        OSCompatibilityManager.py32_check(
            script_path=__file__,
            logger=None  # No logger yet since we haven't selected folder
        )
        return  # Should not reach here as py32_check exits
    
    # If we get here, we're either:
    # 1. Already in 32-bit Python from the start, OR 
    # 2. In the relaunched 32-bit process
    
    if is_relaunched:
        print("Running in relaunched 32-bit Python process")
    else:
        print("Already running in 32-bit Python")
    
    # NOW get folder path (only once, in the correct Python version)
    data_folder = get_folder_from_user()
    
    if not data_folder:
        print("No folder selected, exiting")
        return
    
    print(f"Selected folder: {data_folder}")
    
    # NOW import package modules (after folder selection and 32-bit check)
    sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
    from mseqauto.utils import setup_logger  # type: ignore
    
    # Setup logger
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_dir = os.path.join(script_dir, "logs")
    logger = setup_logger("ind_process_all", log_dir=log_dir)
    
    logger.info("Running in 32-bit Python, proceeding with processing")
    
    print(f"\nSelected folder: {data_folder}")
    print("Starting complete IND processing pipeline...")
    print("Steps: Sort Files → mSeq Processing → Zip Files")
    
    # Define the processing steps
    steps = [
        ("File Sorting", run_sort_files),
        ("mSeq Processing", run_mseq_processing), 
        ("File Zipping", run_zip_files)
    ]
    
    # Run each step in sequence
    for i, (step_name, step_func) in enumerate(steps, 1):
        print(f"\n{'='*50}")
        print(f"STEP {i}/3: {step_name.upper()}")
        print(f"{'='*50}")
        
        try:
            success = step_func(data_folder)
            if success:
                print(f"✓ {step_name} completed successfully")
            else:
                show_error_and_exit(step_name, "Step returned failure status")
                
        except Exception as e:
            show_error_and_exit(step_name, str(e))
    
    # All steps completed successfully
    print(f"\n{'='*60}")
    print("🎉 ALL IND PROCESSING COMPLETE! 🎉")
    print(f"{'='*60}")
    print(f"Processed folder: {data_folder}")
    
    # Show completion popup
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Processing Complete", "All IND processing steps completed successfully!")
    root.destroy()

if __name__ == "__main__":
    main()