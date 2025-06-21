# plate_sort_complete.py
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import logging
from pathlib import Path
import shutil
from datetime import datetime

# Add parent directory to PYTHONPATH for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Check if running under GUI
GUI_MODE = os.environ.get('MSEQAUTO_GUI_MODE', 'False') == 'True'

def setup_logger(name):
    """Set up a logger with file and console output"""
    # Create logs directory if it doesn't exist
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
    os.makedirs(log_dir, exist_ok=True)
    
    # Create logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # File handler
    log_file = os.path.join(log_dir, f"{name}_{datetime.now().strftime('%Y%m%d')}.log")
    file_handler = logging.FileHandler(log_file)
    file_handler.setLevel(logging.DEBUG)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # Formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    # Add handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

def get_folder_from_user():
    """Open a dialog to select today's data folder"""
    # Check for GUI-provided folder first

    if GUI_MODE and 'MSEQAUTO_DATA_FOLDER' in os.environ:
        return os.environ['MSEQAUTO_DATA_FOLDER']
    
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force an update to ensure dialog shows
    
    folder_path = filedialog.askdirectory(
        title="Select today's data folder to process",
        mustexist=True
    )
    
    root.destroy()
    return folder_path

def get_raw_plate_folders(data_folder):
    """Find raw plate data folders matching plate formats and exclude individual folders"""
    raw_folders = []
    
    # Pattern for valid plate folders:
    # - Starts with YYYY-MM-DD date format
    # - Followed by underscore and some plate identifiers
    # - Ends with _1
    plate_pattern = re.compile(r'^\d{4}-\d{2}-\d{2}_.*_1$')
    
    # Pattern to exclude individual folders
    exclude_pattern = re.compile(r'.*[Bb]io[Ii]-\d+.*')
    
    for item in os.listdir(data_folder):
        item_path = os.path.join(data_folder, item)
        if os.path.isdir(item_path) and plate_pattern.match(item):
            # Skip if it matches the exclude pattern (BioI folders)
            if exclude_pattern.match(item):
                continue
            
            raw_folders.append(item_path)
            
    return raw_folders

def extract_plate_prefixes(folder_path):
    """Extract unique plate prefixes from all .ab1 files in the folder"""
    prefixes = set()
    
    for file in os.listdir(folder_path):
        if file.endswith('.ab1'):
            # Extract prefix (text before first underscore)
            parts = file.split('_', 1)
            if len(parts) > 1:
                prefixes.add(parts[0])
    
    return prefixes

def clean_filename_by_prefix(filename, prefix):
    """Remove the prefix and underscore from the filename"""
    if filename.startswith(prefix + '_'):
        return filename[len(prefix) + 1:]  # +1 for the underscore
    return filename

def is_control_file(file_name, controls):
    """Check if file is a control based on name patterns"""
    clean_name = file_name.lower()
    
    for control in controls:
        if control.lower() in clean_name:
            return True
    
    return False

def is_blank_file(file_name):
    """Check if file is a blank sample using various patterns"""
    # # Pattern for plate blanks: 01A__.ab1, 02B__.ab1
    # if re.match(r'^\d{2}[A-H]__.ab1$', file_name, re.IGNORECASE):
    #     return True
    
    # Alternative pattern sometimes used
    if re.match(r'^.*__.ab1$', file_name, re.IGNORECASE):
        return True
        
    return False

def is_cross_plate_file(file_name):
    """
    Check if a file has cross-plate information in braces
    Example: P37952a_11A_12A_18Apr2025_72.02_YSD_plate3_ysd3_seq.F{P37955}.ab1
    """
    # Pattern for plate number in braces at the end of the filename
    # Matches {P#####} or {p#####} or similar variations before the .ab1 extension
    cross_plate_pattern = re.compile(r'.*{[Pp]\d+[a-zA-Z]?}\.ab1$')
    
    if cross_plate_pattern.match(file_name):
        return True
    return False

def extract_destination_plate(file_name):
    """
    Extract the destination plate number from a cross-plate file
    Example: "P37952a_11A_12A_sample{P37955}.ab1" returns "P37955"
    """
    # Extract the content between the last { and }
    match = re.search(r'{([^{}]+)}\.ab1$', file_name)
    if match:
        return match.group(1)
    return None

def remove_braces(file_path):
    """Remove anything in braces from the filename"""
    if '{' not in file_path and '}' not in file_path:
        return file_path

    dir_name = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    new_name = re.sub(r'{.*?}', '', base_name)

    new_path = os.path.join(dir_name, new_name)

    try:
        if os.path.exists(file_path):
            os.rename(file_path, new_path)
            return new_path
    except Exception as e:
        print(f"Error renaming file {file_path}: {e}")

    return file_path

def batch_create_plate_folders(raw_folder, parent_folder, logger):
    """
    Create plate folders based on file prefixes and organize files
    
    Args:
        raw_folder: Path to raw data folder containing .ab1 files
        parent_folder: Path to parent folder where plate folders should be created
        logger: Logger object
        
    Returns:
        dict: Map of created plate folder names to their paths
    """
    # Get unique plate prefixes
    prefixes = extract_plate_prefixes(raw_folder)
    logger.info(f"Found {len(prefixes)} unique plate prefixes: {', '.join(prefixes)}")
    
    # Create folders for each prefix in the parent directory
    created_folders = {}
    for prefix in prefixes:
        plate_folder = os.path.join(parent_folder, prefix)
        os.makedirs(plate_folder, exist_ok=True)
        created_folders[prefix] = plate_folder
        logger.info(f"Created folder: {plate_folder}")
    
    # Process files
    for file in os.listdir(raw_folder):
        if not file.endswith('.ab1'):
            continue
            
        # Find matching prefix
        matching_prefix = None
        for prefix in prefixes:
            if file.startswith(prefix + '_'):
                matching_prefix = prefix
                break
                
        if not matching_prefix:
            logger.warning(f"No matching prefix found for file: {file}")
            continue
            
        # Clean filename and move to appropriate folder
        new_filename = clean_filename_by_prefix(file, matching_prefix)
        source_path = os.path.join(raw_folder, file)
        dest_path = os.path.join(created_folders[matching_prefix], new_filename)
        
        try:
            shutil.move(source_path, dest_path)
            logger.debug(f"Moved {file} to {matching_prefix}/{new_filename}")
        except Exception as e:
            logger.error(f"Error moving file {file}: {e}")
    
    return created_folders

def sort_plate_files(plate_folder, logger, all_plate_folders=None, parent_folder=None):
    """
    Sort files within a plate folder into Controls and Blank subfolders,
    and handle files with braces
    
    Args:
        plate_folder: Path to plate folder
        logger: Logger object
        all_plate_folders: Dictionary of {plate_prefix: folder_path} for cross-plate handling
        parent_folder: Parent directory for creating new plate folders for cross-plate files
    """
    # Define control patterns
    controls = [
        'pgem_m13f-20',
        'pgem_m13r-27',
        'pgem_t7promoter',
        'pgem_sp6promoter',
        'water_m13f-20',
        'water_m13r-27',
        'water_t7promoter',
        'water_sp6promoter'
    ]
    
    # Create subfolders if needed
    controls_folder = os.path.join(plate_folder, "Controls")
    blank_folder = os.path.join(plate_folder, "Blank")
    
    # If parent_folder not provided, use the containing directory of the plate folder
    if parent_folder is None:
        parent_folder = os.path.dirname(plate_folder)
    
    # Process each file
    file_count = 0
    control_count = 0
    blank_count = 0
    brace_count = 0
    cross_plate_count = 0
    
    for file in os.listdir(plate_folder):
        file_path = os.path.join(plate_folder, file)
        
        # Skip folders and non-AB1 files
        if os.path.isdir(file_path) or not file.endswith('.ab1'):
            continue
            
        file_count += 1
            
        # Check if it's a control file
        if is_control_file(file, controls):
            os.makedirs(controls_folder, exist_ok=True)
            dest_path = os.path.join(controls_folder, file)
            try:
                shutil.move(file_path, dest_path)
                control_count += 1
                logger.debug(f"Moved control file: {file}")
            except Exception as e:
                logger.error(f"Error moving control file {file}: {e}")
            continue
            
        # Check if it's a blank file
        if is_blank_file(file):
            os.makedirs(blank_folder, exist_ok=True)
            dest_path = os.path.join(blank_folder, file)
            try:
                shutil.move(file_path, dest_path)
                blank_count += 1
                logger.debug(f"Moved blank file: {file}")
            except Exception as e:
                logger.error(f"Error moving blank file {file}: {e}")
            continue
            
        # Check if it's a cross-plate file
        if is_cross_plate_file(file):
            cross_plate_count += 1
            dest_plate = extract_destination_plate(file)
            
            if dest_plate:
                # Always create a folder in the parent directory with the plate name from braces
                dest_folder = os.path.join(parent_folder, dest_plate)
                os.makedirs(dest_folder, exist_ok=True)
                
                try:
                    # Remove braces before moving to destination plate
                    new_filename = re.sub(r'{[^{}]+}', '', file)
                    dest_path = os.path.join(dest_folder, new_filename)
                    shutil.move(file_path, dest_path)
                    logger.info(f"Moved cross-plate file to destination plate folder: {file} to {dest_plate}/{new_filename}")
                except Exception as e:
                    logger.error(f"Error moving cross-plate file {file}: {e}")
            else:
                logger.warning(f"Could not extract plate name from cross-plate file: {file}")
                # Just remove braces in place if we can't extract a plate name
                try:
                    remove_braces(file_path)
                    logger.debug(f"Removed braces from cross-plate file: {file}")
                except Exception as e:
                    logger.error(f"Error removing braces from {file}: {e}")
            continue
            
        # Remove braces from remaining filenames
        if '{' in file or '}' in file:
            try:
                remove_braces(file_path)
                brace_count += 1
                logger.debug(f"Removed braces from: {file}")
            except Exception as e:
                logger.error(f"Error removing braces from {file}: {e}")
    
    logger.info(f"Processed {file_count} files in {os.path.basename(plate_folder)}: "
                f"{control_count} controls, {blank_count} blanks, "
                f"{cross_plate_count} cross-plate, {brace_count} with braces")

def process_raw_folder(raw_folder, data_folder, logger):
    """Process a single raw plate data folder"""
    folder_name = os.path.basename(raw_folder)
    logger.info(f"Processing raw folder: {folder_name}")
    
    # Check if this is a single plate folder
    # Format: YYYY-MM-DD_P#####_1
    parts = folder_name.split('_')
    is_single_plate = False
    
    if len(parts) == 3 and '-' not in parts[1]:
        # This is a single plate folder - directly rename it
        logger.info(f"Detected single plate folder: {folder_name}")
        is_single_plate = True
        
        # Extract plate name (P##### part)
        plate_name = parts[1]
        
        # Create destination path in parent folder
        dest_folder = os.path.join(data_folder, plate_name)
        
        if os.path.exists(dest_folder):
            logger.warning(f"Destination folder already exists: {dest_folder}")
            # Add a timestamp to avoid conflicts
            timestamp = datetime.now().strftime("%H%M%S")
            dest_folder = os.path.join(data_folder, f"{plate_name}_{timestamp}")
            logger.info(f"Using alternate destination: {dest_folder}")
        
        # Move all files to new destination
        os.makedirs(dest_folder, exist_ok=True)
        
        # Move all files
        for file in os.listdir(raw_folder):
            source_path = os.path.join(raw_folder, file)
            if os.path.isfile(source_path) and file.endswith('.ab1'):
                dest_path = os.path.join(dest_folder, file)
                try:
                    shutil.move(source_path, dest_path)
                    logger.debug(f"Moved {file} to {plate_name}/")
                except Exception as e:
                    logger.error(f"Error moving file {file}: {e}")
        
        # Create a fake plate_folders dict with just the renamed folder
        plate_folders = {plate_name: dest_folder}
    else:
        # Regular multi-plate folder - create plate folders and move files
        logger.info(f"Detected multi-plate folder: {folder_name}")
        plate_folders = batch_create_plate_folders(raw_folder, data_folder, logger)
        logger.info(f"Created {len(plate_folders)} plate folders")
    
    # Step 2: Sort files within each plate folder
    # First, collect all available plate folders for cross-plate handling
    all_plate_prefixes = {}
    for prefix, folder_path in plate_folders.items():
        # Use uppercase prefix as key for case-insensitive matching
        normalized_prefix = prefix.upper()
        all_plate_prefixes[normalized_prefix] = folder_path
        # Also add without the letter suffix for flexibility (P37952a -> P37952)
        base_prefix = re.sub(r'([Pp]\d+)[a-zA-Z].*', r'\1', prefix)
        normalized_base = base_prefix.upper()
        if normalized_base != normalized_prefix:
            all_plate_prefixes[normalized_base] = folder_path
    
    # Now process each plate folder with access to all other plate folders
    for prefix, folder_path in plate_folders.items():
        logger.info(f"Sorting files in plate folder: {prefix}")
        sort_plate_files(folder_path, logger, all_plate_prefixes, data_folder)
    
    # Check if raw folder is now empty (except for created plate folders)
    if not is_single_plate:  # Only for multi-plate folders, single plate folders are handled differently
        remaining_files = [f for f in os.listdir(raw_folder) 
                          if os.path.isfile(os.path.join(raw_folder, f))]
        
        if not remaining_files:
            logger.info(f"Raw folder {folder_name} is now empty of files")
        else:
            logger.warning(f"Raw folder {folder_name} still has {len(remaining_files)} files")
            for file in remaining_files:
                logger.debug(f"Remaining file: {file}")
    
    logger.info(f"Completed processing raw folder: {folder_name}")

def main():
    # Setup logger
    logger = setup_logger("plate_sort_complete")
    logger.info("Starting plate sorting process")
    
    # Get today's data folder
    data_folder = get_folder_from_user()
    
    if not data_folder:
        logger.warning("No folder selected, exiting")
        return
    
    logger.info(f"Selected data folder: {data_folder}")
    
    # Find raw plate data folders
    raw_folders = get_raw_plate_folders(data_folder)
    logger.info(f"Found {len(raw_folders)} raw plate data folders")
    
    if not raw_folders:
        logger.warning("No raw plate data folders found, exiting")
        msg = "No valid plate data folders were found.\n\n"
        msg += "Valid folders should:\n"
        msg += "- Start with a date (YYYY-MM-DD)\n"
        msg += "- End with _1\n"
        msg += "- Not contain 'BioI'\n\n"
        msg += "Please check that you selected the correct data folder."
        messagebox.showwarning("No Raw Folders Found", msg)
        return
    
    # Group folders by type for display
    single_plates = []
    multi_plates = []
    
    for folder in raw_folders:
        folder_name = os.path.basename(folder)
        parts = folder_name.split('_')
        if len(parts) == 3 and '-' not in parts[1]:
            single_plates.append(folder_name)
        else:
            multi_plates.append(folder_name)
    
    # Display confirmation dialog with category information
    msg = f"Found {len(raw_folders)} raw plate data folders to process:\n\n"
    
    if single_plates:
        msg += f"SINGLE PLATE FOLDERS ({len(single_plates)}):\n"
        msg += "\n".join(single_plates)
        msg += "\n\n"
    
    if multi_plates:
        msg += f"MULTI-PLATE FOLDERS ({len(multi_plates)}):\n"
        msg += "\n".join(multi_plates)
        msg += "\n\n"
    
    msg += "Continue?"
    
    if not messagebox.askyesno("Confirm Processing", msg):
        logger.info("User cancelled operation")
        return
    
    # Process each raw folder
    for i, raw_folder in enumerate(raw_folders, 1):
        logger.info(f"Processing folder {i}/{len(raw_folders)}: {os.path.basename(raw_folder)}")
        try:
            process_raw_folder(raw_folder, data_folder, logger)
        except Exception as e:
            logger.error(f"Error processing folder {os.path.basename(raw_folder)}: {e}")
    
    # Check for empty raw folders that could be deleted
    empty_folders = []
    empty_folder_paths = []
    for raw_folder in raw_folders:
        # Check if the folder only contains subdirectories and no files
        has_files = False
        for item in os.listdir(raw_folder):
            if os.path.isfile(os.path.join(raw_folder, item)):
                has_files = True
                break
        
        if not has_files:
            empty_folders.append(os.path.basename(raw_folder))
            empty_folder_paths.append(raw_folder)
    
    # Report on empty folders and offer to delete them
    if empty_folders:
        empty_msg = f"The following {len(empty_folders)} raw folders are now empty of files:\n\n"
        empty_msg += "\n".join(empty_folders)
        empty_msg += "\n\nWould you like to delete these empty folders?"
        logger.info(f"Found {len(empty_folders)} empty raw folders")
        
        if messagebox.askyesno("Delete Empty Folders?", empty_msg):
            # User confirmed deletion
            deleted_count = 0
            for folder_path in empty_folder_paths:
                try:
                    # Remove any subdirectories first (recursive)
                    shutil.rmtree(folder_path)
                    deleted_count += 1
                    logger.info(f"Deleted empty folder: {os.path.basename(folder_path)}")
                except Exception as e:
                    logger.error(f"Error deleting folder {os.path.basename(folder_path)}: {e}")
            
            # Confirm deletion to user
            messagebox.showinfo("Folders Deleted", 
                              f"Successfully deleted {deleted_count} of {len(empty_folders)} empty folders.")
    
    # Show completion message
    success_msg = f"Successfully processed {len(raw_folders)} raw plate data folders."
    logger.info(success_msg)
    messagebox.showinfo("Processing Complete", success_msg)

if __name__ == "__main__":
    main()