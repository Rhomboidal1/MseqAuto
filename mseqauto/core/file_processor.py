# file_processor.py
import os
import re
from mseqauto.utils.path_utilities import standardize_path, remove_braces_from_string

class FileProcessor:
    """
    Handles operations on individual files, including moving, renaming, and categorization.
    """
    
    def __init__(self, file_dao, config, logger=None):
        """
        Initialize the FileProcessor.
        
        Args:
            file_dao: File system data access object
            config: Configuration object
            logger: Logging function
        """
        self.file_dao = file_dao
        self.config = config
        self.logger = logger or (lambda msg: None)  # No-op logger if none provided

    def sort_control_file(self, file_path, destination_folder=None):
        """
        Sort a control file to the Controls folder.
        
        Args:
            file_path: Path to the control file
            destination_folder: Optional override for destination folder
            
        Returns:
            bool: Success status
        """
        file_path = standardize_path(file_path)
        
        # Get parent folder and file name
        parent_folder = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)
        
        # Determine destination Controls folder
        if destination_folder:
            controls_folder = os.path.join(destination_folder, self.config.CONTROLS_FOLDER)
        else:
            controls_folder = os.path.join(parent_folder, self.config.CONTROLS_FOLDER)
        
        # Create Controls folder if it doesn't exist
        if not os.path.exists(controls_folder):
            os.makedirs(controls_folder)
            self.logger(f"Created Controls folder in {os.path.basename(parent_folder)}")
        
        # Move the file to Controls folder
        target_path = os.path.join(controls_folder, file_name)
        success = self.file_dao.move_file(file_path, target_path)
        
        if success:
            self.logger(f"Moved control file {file_name} to Controls folder")
        else:
            self.logger(f"Failed to move control file {file_name}")
            
        return success

    def sort_blank_file(self, file_path, destination_folder=None):
        """
        Sort a blank file to the Blank folder.
        
        Args:
            file_path: Path to the blank file
            destination_folder: Optional override for destination folder
            
        Returns:
            bool: Success status
        """
        file_path = standardize_path(file_path)
        
        # Get parent folder and file name
        parent_folder = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)
        
        # Determine destination Blank folder
        if destination_folder:
            blank_folder = os.path.join(destination_folder, self.config.BLANK_FOLDER)
        else:
            blank_folder = os.path.join(parent_folder, self.config.BLANK_FOLDER)
        
        # Create Blank folder if it doesn't exist
        if not os.path.exists(blank_folder):
            os.makedirs(blank_folder)
            self.logger(f"Created Blank folder in {os.path.basename(parent_folder)}")
        
        # Move the file to Blank folder
        target_path = os.path.join(blank_folder, file_name)
        success = self.file_dao.move_file(file_path, target_path)
        
        if success:
            self.logger(f"Moved blank file {file_name} to Blank folder")
        else:
            self.logger(f"Failed to move blank file {file_name}")
            
        return success

    def sort_pcr_file(self, file_path, pcr_number, destination_folder=None, reinject_list=None, raw_reinject_list=None):
        """
        Sort a PCR file to the appropriate folder.
        
        Args:
            file_path: Path to the PCR file
            pcr_number: PCR number
            destination_folder: Optional override for destination folder
            reinject_list: List of reinjection names
            raw_reinject_list: List of raw reinjection names
            
        Returns:
            bool: Success status
        """
        file_path = standardize_path(file_path)
        file_name = os.path.basename(file_path)
        
        # Get day data folder (parent of parent)
        if destination_folder:
            day_data_path = destination_folder
        else:
            day_data_path = os.path.dirname(os.path.dirname(file_path))
        
        # Create PCR folder name and path
        pcr_folder_name = f"FB-PCR{pcr_number}"
        pcr_folder_path = os.path.join(day_data_path, pcr_folder_name)
        
        # Create PCR folder if it doesn't exist
        if not os.path.exists(pcr_folder_path):
            os.makedirs(pcr_folder_path)
            self.logger(f"Created PCR folder: {pcr_folder_name}")
        
        # Use standard normalization for matching
        normalized_name = self.file_dao.standardize_filename_for_matching(file_name)
        
        # Create the destination path
        clean_name = remove_braces_from_string(file_name)
        target_path = os.path.join(pcr_folder_path, clean_name)
        
        # Check if it's a reinject (if reinject list provided)
        is_reinject = False
        if reinject_list:
            for i, item in enumerate(reinject_list):
                reinj_norm = self.file_dao.standardize_filename_for_matching(item)
                if normalized_name == reinj_norm:
                    is_reinject = True
                    raw_name = raw_reinject_list[i] if raw_reinject_list else item
                    
                    # Check if it's a preemptive reinject
                    if raw_name and '{!P}' in raw_name:
                        # Preemptive reinject goes to main folder
                        success = self.file_dao.move_file(file_path, target_path)
                        if success:
                            self.logger(f"Moved preemptive reinject {file_name} to main PCR folder")
                        return success
                    else:
                        # Regular reinject goes to Alternate Injections
                        alt_inj_folder = os.path.join(pcr_folder_path, self.config.ALT_INJECTIONS_FOLDER)
                        if not os.path.exists(alt_inj_folder):
                            os.makedirs(alt_inj_folder)
                        
                        alt_file_path = os.path.join(alt_inj_folder, file_name)
                        success = self.file_dao.move_file(file_path, alt_file_path)
                        if success:
                            self.logger(f"Moved reinject {file_name} to Alternate Injections")
                        return success
        
        # Regular file goes to main folder
        success = self.file_dao.move_file(file_path, target_path)
        if success:
            self.logger(f"Moved PCR file {file_name} to main folder")
        return success

    def move_to_order_folder(self, file_path, destination_folder, normalized_name=None, reinject_list=None, raw_reinject_list=None):
        """
        Move a file to an order folder with appropriate handling for reinjections.
        
        Args:
            file_path: Path to the file
            destination_folder: Destination order folder
            normalized_name: Normalized file name for reinject matching
            reinject_list: List of reinjection names
            raw_reinject_list: List of raw reinjection names
            
        Returns:
            bool: Success status
        """
        file_path = standardize_path(file_path)
        destination_folder = standardize_path(destination_folder)
        
        file_name = os.path.basename(file_path)
        
        # Normalize the name if not provided
        if normalized_name is None:
            normalized_name = self.file_dao.standardize_filename_for_matching(file_name)
        
        # Clean filename for destination
        clean_name = remove_braces_from_string(file_name)
        target_path = os.path.join(destination_folder, clean_name)
        
        # Check if file already exists at destination
        if os.path.exists(target_path):
            # File already exists, put in alternate injections
            alt_inj_folder = os.path.join(destination_folder, self.config.ALT_INJECTIONS_FOLDER)
            if not os.path.exists(alt_inj_folder):
                os.makedirs(alt_inj_folder)
            
            alt_file_path = os.path.join(alt_inj_folder, file_name)
            success = self.file_dao.move_file(file_path, alt_file_path)
            if success:
                self.logger(f"File already exists, moved {file_name} to alternate injections")
            return success
        
        # Check if it's a reinject
        if reinject_list and normalized_name in reinject_list:
            idx = reinject_list.index(normalized_name)
            raw_name = raw_reinject_list[idx] if raw_reinject_list and idx < len(raw_reinject_list) else None
            
            # Check for preemptive reinject
            if raw_name and '{!P}' in raw_name:
                # Preemptive reinject goes to main folder
                success = self.file_dao.move_file(file_path, target_path)
                if success:
                    self.logger(f"Preemptive reinject moved to main folder: {file_name}")
                return success
            else:
                # Regular reinject goes to alternate injections
                alt_inj_folder = os.path.join(destination_folder, self.config.ALT_INJECTIONS_FOLDER)
                if not os.path.exists(alt_inj_folder):
                    os.makedirs(alt_inj_folder)
                
                alt_file_path = os.path.join(alt_inj_folder, file_name)
                success = self.file_dao.move_file(file_path, alt_file_path)
                if success:
                    self.logger(f"Reinject moved to alternate injections: {file_name}")
                return success
        
        # Regular file, put in main folder
        success = self.file_dao.move_file(file_path, target_path)
        if success:
            self.logger(f"File moved to main folder: {file_name}")
        return success

    def remove_braces_from_filename(self, file_path):
        """
        Remove braces and their contents from a filename.
        
        Args:
            file_path: Path to the file
            
        Returns:
            str: New file path if renamed, original path if not changed or error
        """
        file_path = standardize_path(file_path)
        
        # Check if file exists
        if not os.path.isfile(file_path):
            self.logger(f"File not found: {file_path}")
            return file_path
            
        dir_name = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)
        
        # Check if filename contains braces
        if '{' not in file_name and '}' not in file_name:
            return file_path
            
        # Create new filename
        new_name = remove_braces_from_string(file_name)
        if new_name == file_name:
            return file_path
            
        new_path = os.path.join(dir_name, new_name)
        
        try:
            os.rename(file_path, new_path)
            self.logger(f"Renamed: {file_name} -> {new_name}")
            return new_path
        except Exception as e:
            self.logger(f"Failed to rename {file_name}: {str(e)}")
            return file_path