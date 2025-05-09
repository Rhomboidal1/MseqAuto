# folder_processor.py
import re
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

print(sys.path)

from mseqauto.config import MseqConfig

config = MseqConfig()

class FolderProcessor:
    def __init__(self, file_dao, ui_automation, config, logger=None):
        self.file_dao = file_dao
        self.ui_automation = ui_automation
        self.config = config
        
        # Create a unified logging interface regardless of input type
        import logging
        if logger is None:
            # No logger provided, use default
            self._logger = logging.getLogger(__name__)
            self.log = self._logger.info
        elif isinstance(logger, logging.Logger):
            # It's a Logger object
            self._logger = logger
            self.log = logger.info
        else:
            # Assume it's a callable function
            self._logger = None
            self.log = logger
            
        # Keep original reference for backward compatibility
        self.logger = logger
        
        # Initialize other attributes
        self.order_key_index = None
        self.reinject_list = []
        self.raw_reinject_list = []

    def build_order_key_index(self, order_key):
        """Build lookup index for faster order key searches"""
        if self.order_key_index is not None:
            return  # Already built

        self.order_key_index = {}

        # Process each entry in the order key
        for entry in order_key:
            i_num, acct_name, order_num, sample_name = entry[0:4]
            normalized_name = self.file_dao.normalize_filename(sample_name, remove_extension=False)

            # Create entry in index (handle multiple entries with same normalized name)
            if normalized_name not in self.order_key_index:
                self.order_key_index[normalized_name] = []
            self.order_key_index[normalized_name].append((i_num, acct_name, order_num))

        self.log(f"Built order key index with {len(self.order_key_index)} unique entries")

    def sort_customer_file(self, file_path, order_key):
        """Sort a customer file based on order key using the index"""
        # Build index if not already done
        if self.order_key_index is None:
            self.build_order_key_index(order_key)

        file_name = os.path.basename(file_path)
        # Only log once, not for each transformation step
        self.log(f"Processing customer file: {file_name}")

        # Normalize the filename for matching
        normalized_name = self.file_dao.normalize_filename(file_name)

        # Check if we have this filename in our index
        if normalized_name in self.order_key_index:
            matches = self.order_key_index[normalized_name]

            # Prioritize matches from current folder's I number
            current_i_num = self.file_dao.get_inumber_from_name(os.path.dirname(file_path))

            for i_num, acct_name, order_num in matches:
                # Prioritize current I number if available
                if current_i_num and i_num == current_i_num:
                    destination_folder = self._create_and_get_order_folder(i_num, acct_name, order_num)
                    return self._move_file_to_destination(file_path, destination_folder, normalized_name)

            # If no match with current I number, use the first match
            i_num, acct_name, order_num = matches[0]
            destination_folder = self._create_and_get_order_folder(i_num, acct_name, order_num)
            return self._move_file_to_destination(file_path, destination_folder, normalized_name)

        # No match found
        self.log(f"No match found in order key for: {normalized_name}")
        return False

    def _create_and_get_order_folder(self, i_num, acct_name, order_num):
        """Create order folder structure and return the path"""
        # Create order folder name 
        order_folder_name = f"BioI-{i_num}_{acct_name}_{order_num}"
        self.log(f"Target order folder: {order_folder_name}")

        # Get parent folder path
        parent_folder = self._get_destination_for_order_by_inum(i_num)
        self.log(f"Parent folder: {parent_folder}")

        # Create BioI folder if it doesn't exist
        bioi_folder_name = f"BioI-{i_num}"
        bioi_folder_path = os.path.join(parent_folder, bioi_folder_name)
        if not os.path.exists(bioi_folder_path):
            os.makedirs(bioi_folder_path)
            self.log(f"Created BioI folder: {bioi_folder_path}")

        # Create full order folder path inside BioI folder
        order_folder_path = os.path.join(bioi_folder_path, order_folder_name)
        self.log(f"Order folder path: {order_folder_path}")

        # Create order folder if it doesn't exist
        if not os.path.exists(order_folder_path):
            os.makedirs(order_folder_path)
            self.log(f"Created order folder: {order_folder_path}")

        return order_folder_path

    def get_order_folders(self, bio_folder):
        """Get order folders within a BioI folder"""
        order_folders = []

        for item in os.listdir(bio_folder):
            item_path = os.path.join(bio_folder, item)
            if os.path.isdir(item_path):
                if re.search(r'bioi-\d+_.+_\d+', item.lower()) and not re.search('reinject', item.lower()):
                    # If item looks like an order folder BioI-20000_YoMama_123456 and not a reinject folder
                    order_folders.append(item_path)

        return order_folders

    def get_destination_for_order(self, order_folder, base_path):
        """Find the correct destination for an order folder based on its I-number"""
        folder_name = os.path.basename(order_folder)
        
        # Extract I-number from folder name
        i_num = self.file_dao.get_inumber_from_name(folder_name)
        self.log(f"Finding destination for order with I-number: {i_num}")
        
        # Search parent folder for exact I-number match
        day_data_path = os.path.dirname(base_path)
        
        # Look for EXACT match of I-number
        for item in os.listdir(day_data_path):
            item_path = os.path.join(day_data_path, item)
            if os.path.isdir(item_path) and f"BioI-{i_num}" in item:
                self.log(f"Found exact I-number match: {item_path}")
                return item_path
        
        # If no match found, create new folder
        new_folder = os.path.join(day_data_path, f"BioI-{i_num}")
        if not os.path.exists(new_folder):
            os.makedirs(new_folder)
            self.log(f"Created new folder: {new_folder}")
        
        return new_folder

    def _get_destination_for_order_by_inum(self, i_num):
        """
        Get the parent folder path for an order based on I number
        
        Raises:
            PermissionError: If unable to write to the intended directory
            ValueError: If no suitable destination can be found
        """
        self.log(f"Finding destination for I number: {i_num}")

        # User-selected folder is stored in the original path
        # Get the current working folder from the processor's context
        user_selected_folder = getattr(self, 'current_data_folder', None)

        # If we have a user-selected folder, use it first
        if user_selected_folder and os.path.exists(user_selected_folder) and os.access(user_selected_folder, os.W_OK):
            self.log(f"Using user-selected folder: {user_selected_folder}")
            return user_selected_folder

        # Preferred path pattern: P:\Data\MM.DD.YY\
        today = datetime.now().strftime('%m.%d.%y')
        preferred_path = os.path.join('P:', 'Data', today)

        # Try preferred path first
        if os.path.exists(preferred_path) and os.access(preferred_path, os.W_OK):
            # Search for matching I number folder within preferred path
            for item in self.file_dao.get_directory_contents(preferred_path):
                if (self.file_dao.regex_patterns['bioi_folder'].search(item) and
                        not re.search('reinject', item.lower())):
                    dest_path = os.path.join(preferred_path, item)
                    self.log(f"Found matching folder in preferred path: {dest_path}")
                    return dest_path

            self.log(f"No matching folder found, using preferred path: {preferred_path}")
            return preferred_path

        # If preferred path doesn't work, try other methods
        search_paths = [
            os.path.join('P:', 'Data', 'Individuals'),
            os.path.join('P:', 'Data'),
            os.path.dirname(os.getcwd())
        ]

        for search_path in search_paths:
            if os.path.exists(search_path) and os.access(search_path, os.W_OK):
                # Search for matching I number folder
                for item in self.file_dao.get_directory_contents(search_path):
                    if (self.file_dao.regex_patterns['inumber'].search(item) and
                            not re.search('reinject', item.lower())):
                        dest_path = os.path.join(search_path, item)
                        self.log(f"Found matching folder in alternative path: {dest_path}")
                        return dest_path

                # If no matching folder, use this path
                self.log(f"Using alternative path: {search_path}")
                return search_path

        # If no suitable path found, raise an error
        error_msg = f"Unable to find a writable destination for I number {i_num}"
        self.log(error_msg)
        raise ValueError(error_msg)

    def _move_file_to_destination(self, file_path, destination_folder, normalized_name):
        """Handle file placement including reinject logic"""
        file_name = os.path.basename(file_path)

        # Clean filename for destination (remove braces)
        clean_brace_file_name = re.sub(r'{.*?}', '', file_name)
        target_file_path = os.path.join(destination_folder, clean_brace_file_name)

        # Check if file is a reinject
        is_reinject = False
        if hasattr(self, 'reinject_list') and self.reinject_list:
            is_reinject = normalized_name in [self.file_dao.standardize_filename_for_matching(r) for r in
                                              self.reinject_list]
        self.log(f"Is reinject: {is_reinject}")

        # Handle file placement
        if os.path.exists(target_file_path):
            # File already exists, put in alternate injections
            alt_inj_folder = os.path.join(destination_folder, "Alternate Injections")
            if not os.path.exists(alt_inj_folder):
                os.makedirs(alt_inj_folder)

            alt_file_path = os.path.join(alt_inj_folder, file_name)
            self.file_dao.move_file(file_path, alt_file_path)
            self.log(f"File already exists, moved to alternate injections")
            return True

        elif is_reinject:
            # Handle reinjections
            raw_name_idx = self.reinject_list.index(normalized_name)
            raw_name = self.raw_reinject_list[raw_name_idx]

            # Check for preemptive reinject
            if '{!P}' in raw_name:
                # Preemptive reinject goes to main folder
                self.file_dao.move_file(file_path, target_file_path)
                self.log(f"Preemptive reinject moved to main folder")
            else:
                # Regular reinject goes to alternate injections
                alt_inj_folder = os.path.join(destination_folder, "Alternate Injections")
                if not os.path.exists(alt_inj_folder):
                    os.makedirs(alt_inj_folder)

                alt_file_path = os.path.join(alt_inj_folder, file_name)
                self.file_dao.move_file(file_path, alt_file_path)
                self.log(f"Reinject moved to alternate injections")
            return True

        else:
            # Regular file, put in main folder
            self.file_dao.move_file(file_path, target_file_path)
            self.log(f"File moved to main folder")
            return True

    def _get_expected_file_count(self, order_number):
        """Get expected number of files for an order based on the order key"""
        # Load the order key file
        order_key = self.file_dao.load_order_key(self.config.KEY_FILE_PATH)
        if order_key is None:
            self.log(f"Warning: Could not load order key file, unable to verify count for order {order_number}")
            return 0

        # Count matching entries for this order number
        count = 0
        for row in order_key:
            if str(row[2]) == str(order_number):
                count += 1

        return count

    def process_bio_folder(self, folder):
        """Process a BioI folder (specialized for IND)"""
        self.log(f"Processing BioI folder: {os.path.basename(folder)}")

        # Get all order folders in this BioI folder
        order_folders = self.get_order_folders(folder)

        for order_folder in order_folders:
            # Skip Andreev's orders for mSeq processing
            if self.config.ANDREEV_NAME in order_folder.lower():
                continue

            order_number = self.get_order_number_from_folder_name(order_folder)
            ab1_files = self.file_dao.get_files_by_extension(order_folder, config.ABI_EXTENSION)

            # Check order status
            was_mseqed, has_braces, has_ab1_files = self.check_order_status(order_folder)

            if not was_mseqed and not has_braces:
                # Process if we have the right number of ab1 files
                expected_count = self._get_expected_file_count(order_number)

                if expected_count > 0 and len(ab1_files) == expected_count:
                    if has_ab1_files:
                        self.ui_automation.process_folder(order_folder)
                        self.log(f"mSeq completed: {os.path.basename(order_folder)}")
                else:
                    # Move to Not Ready folder if incomplete
                    not_ready_path = os.path.join(os.path.dirname(folder), self.config.IND_NOT_READY_FOLDER)
                    self.file_dao.create_folder_if_not_exists(not_ready_path)
                    self.file_dao.move_folder(order_folder, not_ready_path)
                    self.log(f"Order moved to Not Ready: {os.path.basename(order_folder)}")

    def process_order_folder(self, order_folder, data_folder_path):
        """Process an order folder"""
        self.log(f"Processing order folder: {os.path.basename(order_folder)}")

        # Skip Andreev's orders for mSeq processing
        if self.config.ANDREEV_NAME in order_folder.lower():
            order_number = self.get_order_number_from_folder_name(order_folder)
            ab1_files = self.file_dao.get_files_by_extension(order_folder, config.ABI_EXTENSION)
            expected_count = self._get_expected_file_count(order_number)

            # For Andreev's orders, just check if complete to move back if needed
            if expected_count > 0 and len(ab1_files) == expected_count:
                # If we're processing from IND Not Ready, move it back
                if os.path.basename(os.path.dirname(order_folder)) == self.config.IND_NOT_READY_FOLDER:
                    destination = self.get_destination_for_order(order_folder, data_folder_path)
                    self.file_dao.move_folder(order_folder, destination)
                    self.log(f"Andreev's order moved back: {os.path.basename(order_folder)}")
            return

        order_number = self.get_order_number_from_folder_name(order_folder)
        ab1_files = self.file_dao.get_files_by_extension(order_folder, config.ABI_EXTENSION)

        # Check order status
        was_mseqed, has_braces, has_ab1_files = self.check_order_status(order_folder)

        # Process based on status
        if not was_mseqed and not has_braces:
            expected_count = self._get_expected_file_count(order_number)

            if expected_count > 0 and len(ab1_files) == expected_count and has_ab1_files:
                self.ui_automation.process_folder(order_folder)
                self.log(f"mSeq completed: {os.path.basename(order_folder)}")

                # If processing from IND Not Ready, move it back
                if os.path.basename(os.path.dirname(order_folder)) == self.config.IND_NOT_READY_FOLDER:
                    destination = self.get_destination_for_order(order_folder, data_folder_path)
                    self.file_dao.move_folder(order_folder, destination)
            else:
                # Move to Not Ready if incomplete
                not_ready_path = os.path.join(os.path.dirname(data_folder_path), self.config.IND_NOT_READY_FOLDER)
                self.file_dao.create_folder_if_not_exists(not_ready_path)
                self.file_dao.move_folder(order_folder, not_ready_path)
                self.log(f"Order moved to Not Ready: {os.path.basename(order_folder)}")

        # If already mSeqed but in IND Not Ready, move it back
        elif was_mseqed and os.path.basename(os.path.dirname(order_folder)) == self.config.IND_NOT_READY_FOLDER:
            destination = self.get_destination_for_order(order_folder, data_folder_path)
            self.file_dao.move_folder(order_folder, destination)
            self.log(f"Processed order moved back: {os.path.basename(order_folder)}")

    def get_pcr_folder_path(self, pcr_number, base_path):
        """Get proper PCR folder path or create one if needed"""
        pcr_folder_name = f"FB-PCR{pcr_number}"

        # Check for existing folders with this PCR number
        for item in os.listdir(base_path):
            if os.path.isdir(os.path.join(base_path, item)):
                if re.search(f'fb-pcr{pcr_number}', item.lower()):
                    return os.path.join(base_path, item)

        # If no folder exists, create one
        new_folder_path = os.path.join(base_path, pcr_folder_name)
        self.file_dao.create_folder_if_not_exists(new_folder_path)
        return new_folder_path

    def process_pcr_folder(self, pcr_folder):
        """Process all files in a PCR folder - simpler matching logic"""
        self.log(f"Processing PCR folder: {os.path.basename(pcr_folder)}")

        for file in os.listdir(pcr_folder):
            if file.endswith(config.ABI_EXTENSION):
                # Clean the filename for matching with reinject list
                clean_name = self.file_dao.standardize_filename_for_matching(file)

                # Check if it's in the reinject list
                is_reinject = clean_name in self.reinject_list

                if is_reinject:
                    # Make sure Alternate Injections folder exists
                    alt_folder = os.path.join(pcr_folder, "Alternate Injections")
                    if not os.path.exists(alt_folder):
                        os.mkdir(alt_folder)

                    # Move file to Alternate Injections
                    source = os.path.join(pcr_folder, file)
                    dest = os.path.join(alt_folder, file)  # Keep original filename with brackets
                    if os.path.exists(source) and not os.path.exists(dest):
                        self.file_dao.move_file(source, dest)
                        self.log(f"Moved reinject {file} to Alternate Injections")

        # Now process the folder with mSeq if it's ready
        was_mseqed, has_braces, has_ab1_files = self.check_order_status(pcr_folder)
        if not was_mseqed and not has_braces and has_ab1_files:
            self.ui_automation.process_folder(pcr_folder)
            self.log(f"mSeq completed: {os.path.basename(pcr_folder)}")
        else:
            self.log(f"mSeq NOT completed: {os.path.basename(pcr_folder)}")

    def sort_ind_folder(self, folder_path, reinject_list, order_key):
        """Sort all files in a BioI folder using batch processing"""
        self.log(f"Processing folder: {folder_path}")

        # Store reinject lists for use in methods
        self.reinject_list = reinject_list
        self.raw_reinject_list = getattr(self, 'raw_reinject_list', reinject_list)

        # Extract I number from the folder
        i_num = self.file_dao.get_inumber_from_name(folder_path)

        # If no I-number found, try to find it from the ab1 files
        if not i_num:
            ab1_files = self.file_dao.get_files_by_extension(folder_path, ".ab1")
            for file_path in ab1_files:
                parent_dir = os.path.basename(os.path.dirname(file_path))
                i_num = self.file_dao.get_inumber_from_name(parent_dir)
                if i_num:
                    self.log(f"Found I number {i_num} from parent directory of AB1 file")
                    break

        # Create or find the target BioI folder first
        if i_num:
            new_folder_name = f"BioI-{i_num}"
            new_folder_path = os.path.join(os.path.dirname(folder_path), new_folder_name)

            # Create the new folder if it doesn't exist
            if not os.path.exists(new_folder_path):
                os.makedirs(new_folder_path)
                self.log(f"Created new BioI folder: {new_folder_path}")
        else:
            # If no I number found, use the original folder
            new_folder_path = folder_path
            self.log(f"No I number found, using original folder: {folder_path}")

        # Get all .ab1 files in the folder
        ab1_files = self.file_dao.get_files_by_extension(folder_path, ".ab1")
        self.log(f"Found {len(ab1_files)} .ab1 files in folder")

        # Group files by type for batch processing
        pcr_files = {}
        control_files = []
        blank_files = []
        customer_files = []

        # Classify files first - FIXED VERSION with no duplicate check
        for file_path in ab1_files:
            file_name = os.path.basename(file_path)

            # Check for PCR files
            pcr_number = self.file_dao.get_pcr_number(file_name)
            if pcr_number:
                if pcr_number not in pcr_files:
                    pcr_files[pcr_number] = []
                pcr_files[pcr_number].append(file_path)
                continue

            # Check for blank files (check this before control files)
            if self.file_dao.is_blank_file(file_name):
                self.log(f"Identified blank file: {file_name}")
                blank_files.append(file_path)
                continue

            # Check for control files
            if self.file_dao.is_control_file(file_name, self.config.CONTROLS):
                self.log(f"Identified control file: {file_name}")
                control_files.append(file_path)
                continue

            # Must be a customer file
            customer_files.append(file_path)

        # Detailed logging for debugging
        self.log(
            f"Classified {len(pcr_files)} PCR numbers, {len(control_files)} controls, {len(blank_files)} blanks, {len(customer_files)} customer files")

        # Log all blank files for verification
        if blank_files:
            self.log("Blank files identified:")
            for file_path in blank_files:
                self.log(f"  - {os.path.basename(file_path)}")
        else:
            self.log("No blank files were identified in this folder")

        # Process each group
        # Process PCR files by PCR number
        for pcr_number, files in pcr_files.items():
            self.log(f"Processing {len(files)} files for PCR number {pcr_number}")
            for file_path in files:
                self._sort_pcr_file(file_path, pcr_number)

        # Process controls - Now placing in the new BioI folder
        if control_files:
            self.log(f"Processing {len(control_files)} control files")
            controls_folder = os.path.join(new_folder_path, "Controls")
            if not os.path.exists(controls_folder):
                os.makedirs(controls_folder)

            for file_path in control_files:
                target_path = os.path.join(controls_folder, os.path.basename(file_path))
                moved = self.file_dao.move_file(file_path, target_path)
                if moved:
                    self.log(f"Moved control file {os.path.basename(file_path)} to {controls_folder}")
                else:
                    self.log(f"Failed to move control file {os.path.basename(file_path)}")

        # Process blanks - Now placing in the new BioI folder
        if blank_files:
            self.log(f"Processing {len(blank_files)} blank files")
            blank_folder = os.path.join(new_folder_path, "Blank")
            if not os.path.exists(blank_folder):
                os.makedirs(blank_folder)

            for file_path in blank_files:
                target_path = os.path.join(blank_folder, os.path.basename(file_path))
                moved = self.file_dao.move_file(file_path, target_path)
                if moved:
                    self.log(f"Moved blank file {os.path.basename(file_path)} to {blank_folder}")
                else:
                    self.log(f"Failed to move blank file {os.path.basename(file_path)}")

        # Process customer files (with optimized order key lookup)
        if customer_files:
            self.log(f"Processing {len(customer_files)} customer files")
            # Build order key index first
            if self.order_key_index is None:
                self.build_order_key_index(order_key)

            for file_path in customer_files:
                self.sort_customer_file(file_path, order_key)

        # Enhanced cleanup: Check if the original folder is empty or can be safely deleted
        try:
            self._cleanup_original_folder(folder_path, new_folder_path)
        except Exception as e:
            self.log(f"Error during folder cleanup: {e}")

        return new_folder_path

    def _cleanup_original_folder(self, original_folder: str, new_folder: str):
        """
        Enhanced cleanup to remove the original folder if all files have been processed
        """
        if original_folder == new_folder:
            self.log("Original folder is the same as new folder, no cleanup needed")
            return

        # Force refresh directory cache first
        self.file_dao.get_directory_contents(original_folder, refresh=True)

        # Check if any .ab1 files remain in the original folder
        ab1_files = self.file_dao.get_files_by_extension(original_folder, ".ab1")
        if ab1_files:
            self.log(f"Found {len(ab1_files)} remaining .ab1 files in original folder, cannot delete")

            # Log the names of remaining files for debugging
            for file_path in ab1_files:
                self.log(f"Remaining file: {os.path.basename(file_path)}")

            return

        # Refresh directory contents
        remaining_items = self.file_dao.get_directory_contents(original_folder, refresh=True)

        # If completely empty, delete the folder
        if not remaining_items:
            try:
                os.rmdir(original_folder)
                self.log(f"Deleted empty original folder: {original_folder}")
                return
            except Exception as e:
                self.log(f"Failed to delete original folder: {e}")
                return

        # Try to move any remaining Control/Blank folders to the new location
        moved_all = True
        item: str
        for item in list(remaining_items):  # Create a copy of the list to avoid iteration issues
            item_path = os.path.join(original_folder, str(item))  # Explicit conversion to string

            if item in ["Controls", "Blank", "Alternate Injections"]:
                # Try to move this folder to the new location if it's not empty
                if os.path.isdir(item_path) and os.listdir(item_path):
                    try:
                        # Create target folder in new location if needed
                        target_path = os.path.join(new_folder, item)  # Explicit conversion to string
                        if not os.path.exists(target_path):
                            os.makedirs(target_path)

                        # Move all files from old to new location
                        for subitem in os.listdir(item_path):
                            old_file = os.path.join(item_path, subitem)  # Explicit conversion to string
                            new_file = os.path.join(target_path, subitem)  # Explicit conversion to string
                            if os.path.isfile(old_file):
                                self.file_dao.move_file(old_file, new_file)
                                self.log(f"Moved remaining file {subitem} to {target_path}")

                        # Check if folder is now empty and can be deleted
                        if not os.listdir(item_path):
                            os.rmdir(item_path)
                            self.log(f"Deleted now-empty folder: {item}")
                    except Exception as e:
                        self.log(f"Failed to move remaining files from {item}: {e}")
                        moved_all = False
                        continue
                elif os.path.isdir(item_path) and not os.listdir(item_path):
                    # Empty folder - delete it
                    try:
                        os.rmdir(item_path)
                        self.log(f"Deleted empty folder: {item}")
                    except Exception as e:
                        self.log(f"Failed to delete empty folder {item}: {e}")
                        moved_all = False
            else:
                # Non-standard item
                self.log(f"Found non-standard item in folder: {item}")
                moved_all = False
        if moved_all:
            self.log("Successfully moved all special folders to new location")
        else:
            self.log("Some items could not be moved to the new location")

        # Refresh contents one more time after all operations
        remaining_items = self.file_dao.get_directory_contents(original_folder, refresh=True)

        # If we've successfully cleaned everything up, try to delete the original folder
        if not remaining_items:
            try:
                os.rmdir(original_folder)
                self.log(f"Deleted original folder after cleaning: {original_folder}")
            except Exception as e:
                self.log(f"Failed to delete original folder: {e}")
        else:
            self.log(f"Unable to clean up original folder. {len(remaining_items)} items remain.")

    def _sort_pcr_file(self, file_path, pcr_number):
        """Sort a PCR file to the appropriate folder"""
        file_name = os.path.basename(file_path)
        self.log(f"Processing PCR file: {file_name} with PCR Number: {pcr_number}")

        # Get the day data folder
        day_data_path = os.path.dirname(os.path.dirname(file_path))

        # Create PCR folder name and path
        pcr_folder_path = self.get_pcr_folder_path(pcr_number, day_data_path)

        # Create PCR folder if it doesn't exist
        if not os.path.exists(pcr_folder_path):
            os.makedirs(pcr_folder_path)

        # Special handling for PCR files
        # Use standard normalization first (which removes ab1 extension)
        normalized_name = self.file_dao.standardize_filename_for_matching(file_name)

        # Debug output for troubleshooting
        self.log(f"PCR file normalized name: {normalized_name}")

        # For PCR files, manually check against reinject list
        is_reinject = False
        raw_name = None

        if hasattr(self, 'reinject_list') and self.reinject_list:
            for i, item in enumerate(self.reinject_list):
                reinj_norm = self.file_dao.standardize_filename_for_matching(item)
                if normalized_name == reinj_norm:
                    is_reinject = True
                    raw_name = self.raw_reinject_list[i] if hasattr(self, 'raw_reinject_list') else item
                    self.log(f"Found PCR file in reinject list: {item}")
                    break

        # Create the destination path
        clean_name = re.sub(r'{.*?}', '', file_name)
        target_path = os.path.join(pcr_folder_path, clean_name)

        # Handle reinjects
        if is_reinject:
            # Check for preemptive reinject
            if raw_name and '{!P}' in raw_name:
                # Preemptive reinject goes in main folder
                self.log(f"Preemptive reinject moved to main folder")
                return self.file_dao.move_file(file_path, target_path)
            else:
                # Regular reinject goes to alternate injections
                alt_inj_folder = os.path.join(pcr_folder_path, "Alternate Injections")
                if not os.path.exists(alt_inj_folder):
                    os.makedirs(alt_inj_folder)

                alt_file_path = os.path.join(alt_inj_folder, file_name)
                self.log(f"Reinject moved to alternate injections")
                return self.file_dao.move_file(file_path, alt_file_path)
        else:
            # Regular file goes in main folder
            self.log(f"File moved to main folder")
            return self.file_dao.move_file(file_path, target_path)

    def _sort_control_file(self, file_path):
        """Sort a control file to the Controls folder"""
        parent_folder = os.path.dirname(file_path)
        controls_folder = os.path.join(parent_folder, "Controls")

        # Create Controls folder if it doesn't exist
        if not os.path.exists(controls_folder):
            os.makedirs(controls_folder)

        # Just move the file directly (no special handling needed)
        target_path = os.path.join(controls_folder, os.path.basename(file_path))
        return self.file_dao.move_file(file_path, target_path)

    def _sort_blank_file(self, file_path):
        """Sort a blank file to the Blank folder"""
        # Get parent BioI folder first, not just the immediate parent folder
        immediate_parent = os.path.dirname(file_path)
        file_name = os.path.basename(file_path)

        # Get or create the BioI folder
        i_num = self.file_dao.get_inumber_from_name(immediate_parent)
        if i_num:
            bioi_folder_name = f"BioI-{i_num}"
            bioi_folder_path = os.path.join(os.path.dirname(immediate_parent), bioi_folder_name)

            # Ensure BioI folder exists
            if not os.path.exists(bioi_folder_path):
                os.makedirs(bioi_folder_path)

            # Create Blank folder inside BioI folder
            blank_folder = os.path.join(bioi_folder_path, "Blank")
            if not os.path.exists(blank_folder):
                os.makedirs(blank_folder)

            # Move file to Blank folder
            target_path = os.path.join(blank_folder, file_name)
            return self.file_dao.move_file(file_path, target_path)
        else:
            # Fallback to original logic if no I number found
            parent_folder = immediate_parent
            blank_folder = os.path.join(parent_folder, "Blank")
            if not os.path.exists(blank_folder):
                os.makedirs(blank_folder)
            target_path = os.path.join(blank_folder, file_name)
            return self.file_dao.move_file(file_path, target_path)

    def _rename_processed_folder(self, folder_path):
        """Rename the folder after processing"""
        i_num = self.file_dao.get_inumber_from_name(folder_path)
        if i_num:
            new_folder_name = f"BioI-{i_num}"
            new_folder_path = os.path.join(os.path.dirname(folder_path), new_folder_name)

            # Only rename if the new name doesn't already exist
            if not os.path.exists(new_folder_path) and os.path.basename(folder_path) != new_folder_name:
                try:
                    os.rename(folder_path, new_folder_path)
                    self.log(f"Renamed folder to: {new_folder_name}")
                    return True
                except Exception as e:
                    self.log(f"Failed to rename folder: {e}")

        return False

    def get_todays_inumbers_from_folder(self, path):
        """Get I numbers and folder paths from the selected folder"""
        i_numbers = []
        bioi_folders = []

        for item in self.file_dao.get_directory_contents(path):
            item_path = os.path.join(path, item)
            if os.path.isdir(item_path):
                # Check if it's a BioI folder
                if self.file_dao.regex_patterns['inumber'].search(item):
                    i_num = self.file_dao.get_inumber_from_name(item)
                    if i_num and i_num not in i_numbers:
                        i_numbers.append(i_num)

                # Get BioI folders before sorting (avoid reinject folders)
                if (self.file_dao.regex_patterns['bioi_folder'].search(item) and
                        'reinject' not in item.lower()):
                    bioi_folders.append(item_path)

        return i_numbers, bioi_folders

    def get_order_number_from_folder_name(self, folder_path):
        """Extract order number from folder name"""
        folder_name = os.path.basename(folder_path)
        match = re.search(r'_\d+$', folder_name)
        if match:
            order_number = re.search(r'\d+', match.group(0)).group(0)
            return order_number
        return None

    def get_reinject_list(self, i_numbers, reinject_path=None):
        """Get list of reactions that are reinjects - optimized version"""
        import re
        import os
        import numpy as np

        reinject_list = []
        raw_reinject_list = []

        # Cache of processed text files to avoid reprocessing
        processed_files = set()

        # Process text files only once
        spreadsheets_path = r'G:\Lab\Spreadsheets'
        abi_path = os.path.join(spreadsheets_path, 'Individual Uploaded to ABI')

        # Build a list of potential reinject files first
        reinject_files = []

        # Scan directories once
        if os.path.exists(spreadsheets_path):
            for file in self.file_dao.get_directory_contents(spreadsheets_path):
                if 'reinject' in file.lower() and file.endswith('.txt'):
                    # Check if any of our I numbers are in the filename
                    if any(i_num in file for i_num in i_numbers):
                        reinject_files.append(os.path.join(spreadsheets_path, file))

        if os.path.exists(abi_path):
            for file in self.file_dao.get_directory_contents(abi_path):
                if 'reinject' in file.lower() and file.endswith('.txt'):
                    # Check if any of our I numbers are in the filename
                    if any(i_num in file for i_num in i_numbers):
                        reinject_files.append(os.path.join(abi_path, file))

        # Process each found reinject file
        for file_path in reinject_files:
            try:
                if file_path in processed_files:
                    continue

                processed_files.add(file_path)
                data = np.loadtxt(file_path, dtype=str, delimiter='\t')

                # Parse rows 5-101 (B6:B101 in Excel terms)
                for j in range(5, min(101, data.shape[0])):
                    if j < data.shape[0] and data.shape[1] > 1:
                        raw_reinject_list.append(data[j, 1])
                        cleaned_name = self.file_dao.standardize_filename_for_matching(data[j, 1])
                        reinject_list.append(cleaned_name)
            except Exception as e:
                self.log(f"Error processing reinject file {file_path}: {e}")

        # If a reinject_path is provided, also check that Excel file
        if reinject_path and os.path.exists(reinject_path):
            try:
                import pylightxl as xl

                # Read the Excel file
                db = xl.readxl(reinject_path)
                sheet = db.ws('Sheet1')

                # Get reinject entries from Excel
                reinject_prep_list = []
                for row in range(1, sheet.maxrow + 1):
                    if sheet.maxcol >= 2:  # Ensure we have at least 2 columns
                        sample = sheet.index(row, 1)  # Use index instead of address
                        primer = sheet.index(row, 2)  # Use index instead of address
                        if sample and primer:
                            reinject_prep_list.append(sample + primer)

                # Convert to numpy array for easier searching
                reinject_prep_array = np.array(reinject_prep_list)

                # Check for partial plate reinjects by comparing against reinject_prep_array
                for i_num in i_numbers:
                    # Check all txt files for this I number
                    # Fix: Change glob pattern to valid regex pattern
                    regex_pattern = f'.*{i_num}.*txt'
                    txt_files = []

                    for file in self.file_dao.get_directory_contents(spreadsheets_path):
                        if re.search(regex_pattern, file, re.IGNORECASE) and not 'reinject' in file:
                            txt_files.append(os.path.join(spreadsheets_path, file))

                    if os.path.exists(abi_path):
                        for file in self.file_dao.get_directory_contents(abi_path):
                            if re.search(regex_pattern, file, re.IGNORECASE) and not 'reinject' in file:
                                txt_files.append(os.path.join(abi_path, file))

                    # Check each txt file for matches against reinject_prep_array
                    for file_path in txt_files:
                        try:
                            data = np.loadtxt(file_path, dtype=str, delimiter='\t')
                            for j in range(5, min(101, data.shape[0])):
                                if j < data.shape[0] and data.shape[1] > 1:
                                    # Check if this sample is in the reinject list
                                    sample_name = data[j, 1][5:] if len(data[j, 1]) > 5 else data[j, 1]
                                    indices = np.where(reinject_prep_array == sample_name)[0]

                                    if len(indices) > 0:
                                        raw_reinject_list.append(data[j, 1])
                                        cleaned_name = self.file_dao.standardize_filename_for_matching(data[j, 1])
                                        reinject_list.append(cleaned_name)
                        except Exception as e:
                            self.log(f"Error processing txt file {file_path}: {e}")

            except Exception as e:
                self.log(f"Error processing reinject Excel file: {e}")

        # Store the raw_reinject_list for reference in _move_file_to_destination
        self.raw_reinject_list = raw_reinject_list
        return reinject_list

    def test_specific_pcr_sorting(self, pcr_number, folder_path=None):
        """Test PCR file sorting for a specific PCR number"""
        # Set up file output
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"pcr_test_{pcr_number}_{timestamp}.txt"

        # Function to log both to console and file
        def log(message):
            print(message)
            with open(output_file, "a") as f:
                f.write(message + "\n")

        log(f"\n=== Testing PCR {pcr_number} Sorting ===")
        log(f"Output saved to: {output_file}")

        # Load the reinject list
        if not hasattr(self, 'reinject_list') or not self.reinject_list:
            log("Loading reinject list...")
            i_numbers = ['21888', '21889', '21890', '21891']  # Hardcode I numbers for testing
            reinject_path = f"P:\\Data\\Reinjects\\Reinject List_{datetime.now().strftime('%m-%d-%Y')}.xlsx"
            self.reinject_list = self.get_reinject_list(i_numbers, reinject_path)
            log(f"Loaded {len(self.reinject_list)} reinjections")

        if folder_path:
            log(f"\nListing ALL files in folder: {folder_path}")
            log("=" * 50)

            # Get all files in this folder recursively
            all_files = []
            for root, dirs, files in os.walk(folder_path):
                for file in files:
                    all_files.append(os.path.join(root, file))

            # Group by file extension
            file_by_ext = {}
            for file_path in all_files:
                ext = os.path.splitext(file_path)[1].lower()
                if ext not in file_by_ext:
                    file_by_ext[ext] = []
                file_by_ext[ext].append(file_path)

            # Print summary of file types
            log(f"Found {len(all_files)} total files:")
            for ext, files in file_by_ext.items():
                log(f"  {ext}: {len(files)} files")

            # Print all .ab1 files
            log("\nAll .ab1 files:")
            ab1_files = file_by_ext.get(config.ABI_EXTENSION, [])
            for i, file_path in enumerate(ab1_files):
                log(f"  {i + 1}. {os.path.relpath(file_path, folder_path)}")

            # Check file contents for PCR number
            pcr_files = []
            for file_path in ab1_files:
                file_content = os.path.basename(file_path)
                if f"PCR{pcr_number}" in file_content:
                    pcr_files.append(file_path)

            log(f"\nFound {len(pcr_files)} PCR {pcr_number} files")
            if pcr_files:
                log("PCR files found:")
                for i, file_path in enumerate(pcr_files):
                    log(f"  {i + 1}. {os.path.basename(file_path)}")

        else:
            log("No folder specified, using test files")
            pcr_files = [
                "{07E}{06G}940.9.H446_940R{PCR2961exp1}{2_28}{I-21889}.ab1",
                "{06G}940.9.H446_940R{PCR2961exp1}.ab1"
            ]

        # Check reinject list for entries that might match these files
        log("\nChecking reinject list for matches:")
        matches_found = 0

        if hasattr(self, 'reinject_list') and self.reinject_list and pcr_files:
            for file_path in pcr_files:
                file_name = os.path.basename(file_path) if os.path.exists(file_path) else file_path
                log(f"\nFile: {file_name}")

                # Standard normalization (with extension removal)
                std_norm = self.file_dao.standardize_filename_for_matching(file_name)
                log(f"Standard normalized: {std_norm}")

                # PCR normalization (with manual extension removal)
                pcr_norm = self.file_dao.standardize_filename_for_matching(file_name)
                if pcr_norm.endswith(config.ABI_EXTENSION):
                    pcr_norm = pcr_norm[:-4]
                log(f"PCR normalized: {pcr_norm}")

                # Check for matches in reinject list
                found_match = False
                for i, item in enumerate(self.reinject_list):
                    reinj_std_norm = self.file_dao.standardize_filename_for_matching(item)

                    if std_norm == reinj_std_norm:
                        log(f"MATCH found with reinject #{i + 1}: {item}")
                        log(f"  Normalized to: {reinj_std_norm}")
                        found_match = True
                        matches_found += 1
                        break

                if not found_match:
                    log("No matches found in reinject list")
        else:
            log("No reinject list available for testing or no PCR files found")

        log(f"\nFound {matches_found} files that match entries in the reinject list")
        log("===================================")

        # Display all entries in reinject list for reference
        log("\nFull reinject list for reference:")
        if hasattr(self, 'reinject_list') and self.reinject_list:
            for i, item in enumerate(self.reinject_list):
                log(f"{i + 1}. {item}")
        else:
            log("No reinject list available")

        log(f"\nTest completed. Results saved to {output_file}")

    def check_order_status(self, folder_path):
        """
        Check the processing status of a folder

        Args:
            folder_path (str): Path to the folder to check

        Returns:
            tuple: (was_mseqed, has_braces, has_ab1_files)
                - was_mseqed: True if folder has been processed by mSeq
                - has_braces: True if any files have brace tags (e.g., {tag})
                - has_ab1_files: True if folder contains .ab1 files
        """
        was_mseqed = False
        has_braces = False
        has_ab1_files = False

        # Get all files in the folder
        folder_contents = self.file_dao.get_directory_contents(folder_path)

        # Check for mSeq directory structure
        mseq_set = {'chromat_dir', 'edit_dir', 'phd_dir', 'mseq4.ini'}
        current_proj = [item for item in folder_contents if item in mseq_set]

        # A folder has been mSeqed if it contains all the required directories/files
        if set(current_proj) == mseq_set:
            was_mseqed = True

        # Check for the 5 txt files as an additional verification
        txt_file_count = 0
        for item in folder_contents:
            item_path = os.path.join(folder_path, item)
            if os.path.isfile(item_path):
                if (item.endswith('.raw.qual.txt') or
                        item.endswith('.raw.seq.txt') or
                        item.endswith('.seq.info.txt') or
                        item.endswith('.seq.qual.txt') or
                        item.endswith('.seq.txt')):
                    txt_file_count += 1

        # If all 5 txt files are present, that's another indicator it was mSeqed
        if txt_file_count == 5:
            was_mseqed = True

        # Check for .ab1 files and braces
        for item in folder_contents:
            if item.endswith(config.ABI_EXTENSION):
                has_ab1_files = True
                if '{' in item or '}' in item:
                    has_braces = True

        return was_mseqed, has_braces, has_ab1_files

    def zip_order_folder(self, folder_path, include_txt=True):
        """
        Zip the contents of an order folder

        Args:
            folder_path (str): Path to the order folder
            include_txt (bool): Whether to include text files in zip

        Returns:
            str: Path to created zip file, or None if failed
        """
        try:
            folder_name = os.path.basename(folder_path)
            zip_filename = f"{folder_name}.zip"
            zip_path = os.path.join(folder_path, zip_filename)

            # Determine which files to include
            file_extensions = [config.ABI_EXTENSION]
            if include_txt:
                file_extensions.extend(self.config.TEXT_FILES)

            # Create zip file
            self.log(f"Creating zip file: {zip_filename}")
            success = self.file_dao.zip_files(
                source_folder=folder_path,
                zip_path=zip_path,
                file_extensions=file_extensions,
                exclude_extensions=None
            )

            if success:
                self.log(f"Successfully created zip file: {zip_path}")
                return zip_path
            else:
                self.log(f"Failed to create zip file: {zip_path}")
                return None
        except Exception as e:
            self.log(f"Error creating zip file for {folder_path}: {e}")
            return None

    def find_zip_file(self, folder_path):
        """Find zip file in a folder"""
        for item in os.listdir(folder_path):
            if item.endswith('.zip') and os.path.isfile(os.path.join(folder_path, item)):
                return os.path.join(folder_path, item)
        return None


    def get_zip_mod_time(self, worksheet, order_number):
        """Get the modification time of a zip file from the summary worksheet"""
        for row in range(2, worksheet.max_row + 1):
            if worksheet.cell(row=row, column=2).value == order_number:
                return worksheet.cell(row=row, column=8).value
        return None

    def validate_zip_contents(self, zip_path, i_number, order_number, order_key):
        """
        Validate zip file contents against order key

        Args:
            zip_path (str): Path to the zip file
            i_number (str): I number
            order_number (str): Order number
            order_key (array): Order key data

        Returns:
            dict: Validation results
        """
        import zipfile
        import numpy as np

        # Results to return
        validation_result = {
            'match_count': 0,
            'mismatch_count': 0,
            'txt_count': 0,
            'expected_count': 0,
            'matches': [],
            'mismatches_in_zip': [],
            'mismatches_in_order': [],
            'txt_files': []
        }

        # Get expected files from order key for this order
        order_items = []

        # Handle NumPy array properly
        if isinstance(order_key, np.ndarray):
            # Use NumPy's boolean indexing for efficient filtering
            mask = (order_key[:, 0] == i_number) & (order_key[:, 2] == order_number)
            matching_rows = order_key[mask]

            for row in matching_rows:
                raw_name = row[3]
                adjusted_name = self.file_dao.normalize_filename(raw_name)
                order_items.append({'raw_name': raw_name, 'adjusted_name': adjusted_name})
        else:
            # Fallback for non-NumPy arrays
            for entry in order_key:
                if entry[0] == i_number and entry[2] == order_number:
                    raw_name = entry[3]
                    adjusted_name = self.file_dao.normalize_filename(raw_name)
                    order_items.append({'raw_name': raw_name, 'adjusted_name': adjusted_name})

        validation_result['expected_count'] = len(order_items)

        # Get actual files from zip
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_contents = zip_ref.namelist()

        # Check for text files
        txt_extensions = self.config.TEXT_FILES

        for item in zip_contents:
            for ext in txt_extensions:
                if item.endswith(ext):
                    validation_result['txt_count'] += 1
                    validation_result['txt_files'].append(ext)
                    break

        # Check for matches
        zip_contents_to_check = list(zip_contents)  # Create a copy to modify safely

        for order_item in order_items:
            adjusted_name = order_item['adjusted_name']
            raw_name = order_item['raw_name']

            found_match = False
            for zip_item in zip_contents_to_check[:]:  # Iterate over a copy
                if zip_item.endswith(config.ABI_EXTENSION):
                    # Remove braces and extension for comparison
                    clean_zip_item = re.sub(r'{.*?}', '', zip_item)[:-4]  # Remove .ab1

                    if clean_zip_item == adjusted_name:
                        validation_result['matches'].append({
                            'raw_name': raw_name,
                            'file_name': zip_item
                        })
                        validation_result['match_count'] += 1
                        found_match = True
                        zip_contents_to_check.remove(zip_item)  # Remove to avoid duplicates
                        break

            if not found_match:
                validation_result['mismatches_in_order'].append({
                    'raw_name': raw_name
                })
                validation_result['mismatch_count'] += 1

        # Remaining unmatched AB1 files in zip
        for item in zip_contents_to_check:
            if item.endswith(config.ABI_EXTENSION):
                validation_result['mismatches_in_zip'].append(item)
                validation_result['mismatch_count'] += 1

        return validation_result

if __name__ == "__main__":
    # Simple test if run directly
    from mseqauto.config import MseqConfig
    from file_system_dao import FileSystemDAO
    from datetime import datetime

    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    processor = FolderProcessor(file_dao, None, config)

    # Test folder processing
    test_path = os.getcwd()
    print(f"Testing with folder: {test_path}")

    # Test getting I number from a folder name
    test_folder_name = "BioI-12345_Customer_67890"
    i_num = processor.file_dao.get_inumber_from_name(test_folder_name)
    print(f"I number from {test_folder_name}: {i_num}")

    # Add this simple reinject list test
    import tkinter as tk
    from tkinter import filedialog

    # Create a simple dialog to select folder
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Select folder to check reinjections")

    if folder_path:
        print(f"Testing reinject detection for: {folder_path}")

        # Get I numbers from the folder
        i_numbers, _ = processor.get_todays_inumbers_from_folder(folder_path)
        print(f"Found I numbers: {i_numbers}")

        # Get reinject list
        reinject_list = processor.get_reinject_list(i_numbers)

        # Print the results
        print(f"\nReinject List ({len(reinject_list)} entries):")
        for i, item in enumerate(reinject_list):
            print(f"{i + 1}. {item}")
    processor.test_specific_pcr_sorting("2961", folder_path)
    processor.test_specific_pcr_sorting("2872", folder_path)

'''        # Test file with period normalization
        test_filename = "{06G}940.9.H446_940R{PCR2961exp1}"
        normalized_test = file_dao.normalize_filename(test_filename)
        print(f"\nNormalization test:")
        print(f"Original: {test_filename}")
        print(f"Normalized: {normalized_test}")
        
        # Check if this would be found in the reinject list
        is_in_list = False
        for item in reinject_list:
            normalized_item = file_dao.normalize_filename(item)
            if normalized_item == normalized_test:
                is_in_list = True
                print(f"MATCH FOUND: {item} normalizes to {normalized_item}")
                break
        
        if not is_in_list:
            print("NO MATCH FOUND in reinject list")
            
            # Let's see what might be close matches
            print("\nChecking for close matches:")
            for item in reinject_list:
                normalized_item = file_dao.normalize_filename(item)
                if "940" in normalized_item and "H446" in normalized_item:
                    print(f"Possible match: {item}")
                    print(f"  Normalized to: {normalized_item}")
    # Test normalize_filename function with different inputs
    print("\nTesting normalize_filename function:")
    test_filenames = [
        "{06G}940.9.H446_940R{PCR2961exp1}",
        "{06G}940.9.H446_940R",
        "940.9.H446_940R",
        "940.9_H446_940R",
        "940_9_H446_940R"
    ]
    # Add test with remove_extension=False
    print("\nTesting normalize_filename with remove_extension=False:")
    test_filenames = [
        "{06G}940.9.H446_940R{PCR2961exp1}",
        "940.9.H446_940R"
    ]
    
    for test_file in test_filenames:
        normalized = file_dao.normalize_filename(test_file, remove_extension=False)
        print(f"Original: '{test_file}'")
        print(f"Normalized: '{normalized}'")
        print()
        '''
