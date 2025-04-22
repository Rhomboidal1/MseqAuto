# folder_processor.py
import re
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

#print(sys.path)
import time
from datetime import datetime

from mseqauto.config import MseqConfig
import warnings
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

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
                         destination_folder = self.create_order_folder(i_num, acct_name, order_num)
                         return self._move_file_to_destination(file_path, destination_folder, normalized_name)

               # If no match with current I number, use the first match
               i_num, acct_name, order_num = matches[0]
               destination_folder = self.create_order_folder(i_num, acct_name, order_num)
               return self._move_file_to_destination(file_path, destination_folder, normalized_name)

          # No match found
          self.log(f"No match found in order key for: {normalized_name}")
          return False

     # def sort_customer_file(self, file_path, order_key):
     #      """Sort a customer file based on order key using the index"""
     #      # Build index if not already done
     #      if self.order_key_index is None:
     #           self.build_order_key_index(order_key)

     #      file_name = os.path.basename(file_path)
     #      # Only log once, not for each transformation step
     #      self.log(f"Processing customer file: {file_name}")

     #      # Normalize the filename for matching
     #      normalized_name = self.file_dao.normalize_filename(file_name)

     #      # Check if we have this filename in our index
     #      if normalized_name in self.order_key_index:
     #           matches = self.order_key_index[normalized_name]

     #           # Prioritize matches from current folder's I number
     #           current_i_num = self.file_dao.get_inumber_from_name(os.path.dirname(file_path))

     #           for i_num, acct_name, order_num in matches:
     #                # Prioritize current I number if available
     #                if current_i_num and i_num == current_i_num:
     #                     destination_folder = self.create_order_folder(i_num, acct_name, order_num)
     #                     return self._move_file_to_destination(file_path, destination_folder, normalized_name)

     #           # If no match with current I number, use the first match
     #           i_num, acct_name, order_num = matches[0]
     #           destination_folder = self.create_order_folder(i_num, acct_name, order_num)
     #           return self._move_file_to_destination(file_path, destination_folder, normalized_name)

     #      # No match found
     #      self.log(f"No match found in order key for: {normalized_name}")
     #      return False

     def create_order_folder(self, i_num, acct_name, order_num, base_path=None):
          """
          Create order folder structure and return the path
          
          Args:
               i_num: I-number for the order
               acct_name: Account name for the order
               order_num: Order number
               base_path: Optional base path to search from
               
          Returns:
               Path to the created order folder
          """
          # Create order folder name 
          order_folder_name = f"BioI-{i_num}_{acct_name}_{order_num}"
          self.log(f"Target order folder: {order_folder_name}")

          # Get parent folder path with consistent naming scheme
          parent_folder = self._get_bioi_folder_path(i_num, base_path)
          self.log(f"Parent folder: {parent_folder}")
          
          # Create full order folder path inside BioI folder
          order_folder_path = os.path.join(parent_folder, order_folder_name)
          self.log(f"Order folder path: {order_folder_path}")

          # Create order folder if it doesn't exist
          if not os.path.exists(order_folder_path):
               os.makedirs(order_folder_path)
               self.log(f"Created order folder: {order_folder_path}")

          return order_folder_path

     def _get_bioi_folder_path(self, i_num, base_path=None):
          """
          Get or create the standardized BioI folder path
          
          Args:
               i_num: I-number for the folder
               base_path: Optional base path to search from
               
          Returns:
               Path to the BioI folder (always in clean format)
          """
          # Determine the data folder
          data_folder = base_path or getattr(self, 'current_data_folder', None)
          
          # If no data folder is specified, use today's date folder
          if not data_folder:
               from datetime import datetime
               today = datetime.now().strftime('%m.%d.%y')
               data_folder = os.path.join('P:', 'Data', today)
               try:
                    if not os.path.exists(data_folder):
                         os.makedirs(data_folder)
                         self.log(f"Created today's data folder: {data_folder}")
               except Exception as e:
                    self.log(f"Error creating date folder: {e}")
                    # Fallback to current directory
                    data_folder = os.path.dirname(os.getcwd())
          
          # Create path to the standardized BioI folder
          bioi_folder_name = f"BioI-{i_num}"
          bioi_folder_path = os.path.join(data_folder, bioi_folder_name)
          
          # Create the folder if it doesn't exist
          if not os.path.exists(bioi_folder_path):
               try:
                    os.makedirs(bioi_folder_path)
                    self.log(f"Created clean BioI folder: {bioi_folder_path}")
               except Exception as e:
                    self.log(f"Error creating BioI folder: {e}")
          
          return bioi_folder_path

     def sort_ind_folder(self, folder_path, reinject_list, order_key):
          """Sort all files in a BioI folder using batch processing with recursive folder scanning"""
          self.log(f"Processing folder: {folder_path}")

          # Store reinject lists for use in methods
          self.reinject_list = reinject_list
          self.raw_reinject_list = getattr(self, 'raw_reinject_list', reinject_list)

          # Extract I number from the folder
          i_num = self.file_dao.get_inumber_from_name(folder_path)

          # If no I-number found, try to find it from the ab1 files
          if not i_num:
               ab1_files = self.file_dao.get_files_by_extension(folder_path, ".ab1", recursive=True)
               for file_path in ab1_files:
                    parent_dir = os.path.basename(os.path.dirname(file_path))
                    i_num = self.file_dao.get_inumber_from_name(parent_dir)
                    if i_num:
                         self.log(f"Found I number {i_num} from parent directory of AB1 file")
                         break

          # Create or find the target BioI folder - always use clean format
          if i_num:
               new_folder_path = self._get_bioi_folder_path(i_num, os.path.dirname(folder_path))
          else:
               # If no I number found, use the original folder
               new_folder_path = folder_path
               self.log(f"No I number found, using original folder: {folder_path}")

          # Recursively get all .ab1 files in the folder and subfolders
          ab1_files = []
          for root, dirs, files in os.walk(folder_path):
               for file in files:
                    if file.endswith(".ab1"):
                         ab1_files.append(os.path.join(root, file))
          
          self.log(f"Found {len(ab1_files)} .ab1 files in folder and subfolders")

          # Group files by type for batch processing
          pcr_files = {}
          control_files = []
          blank_files = []
          customer_files = []
          unmatched_files = []

          # Classify files first
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

               # Check if it's a customer file (has a match in order key)
               # Pre-normalize for matching
               normalized_name = self.file_dao.normalize_filename(file_name)
               
               # Build order key index if needed
               if self.order_key_index is None and order_key is not None:
                    self.build_order_key_index(order_key)
                    
               # Check if in order key
               if self.order_key_index and normalized_name in self.order_key_index:
                    customer_files.append(file_path)
               else:
                    # No match found - unmatched file
                    unmatched_files.append(file_path)

          # Detailed logging
          self.log(
               f"Classified {len(pcr_files)} PCR numbers, {len(control_files)} controls, "
               f"{len(blank_files)} blanks, {len(customer_files)} customer files, "
               f"{len(unmatched_files)} unmatched files"
          )

          # Process each group
          # Process PCR files by PCR number
          for pcr_number, files in pcr_files.items():
               self.log(f"Processing {len(files)} files for PCR number {pcr_number}")
               for file_path in files:
                    self._sort_pcr_file(file_path, pcr_number)

          # Process controls
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

          # Process blanks
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

          # Process customer files
          if customer_files:
               self.log(f"Processing {len(customer_files)} customer files")
               for file_path in customer_files:
                    self.sort_customer_file(file_path, order_key)

          # Process unmatched files - move to BioI folder root
          if unmatched_files and i_num:
               self.log(f"Processing {len(unmatched_files)} unmatched files")
               # Create an 'Unsorted' folder for files with ambiguous names
               unsorted_folder = None
               
               for file_path in unmatched_files:
                    file_name = os.path.basename(file_path)
                    
                    # Check if this is a preemptive but has no match in order key
                    is_preempt = self.is_preemptive(file_name)
                    
                    if is_preempt:
                         # For preemptive files with no match, put in Unsorted folder
                         if unsorted_folder is None:
                              unsorted_folder = os.path.join(new_folder_path, "Unsorted")
                              if not os.path.exists(unsorted_folder):
                                   os.makedirs(unsorted_folder)
                         
                         # Keep original name for Unsorted folder
                         target_path = os.path.join(unsorted_folder, file_name)
                         moved = self.file_dao.move_file(file_path, target_path)
                         if moved:
                              self.log(f"Moved preemptive unmatched file {file_name} to Unsorted folder")
                         else:
                              self.log(f"Failed to move unmatched file {file_name}")
                    else:
                         # Clean filename for destination (remove braces)
                         clean_brace_file_name = re.sub(r'{.*?}', '', file_name)
                         target_path = os.path.join(new_folder_path, clean_brace_file_name)
                         
                         moved = self.file_dao.move_file(file_path, target_path)
                         if moved:
                              self.log(f"Moved unmatched file {file_name} to BioI folder root")
                         else:
                              self.log(f"Failed to move unmatched file {file_name}")

          # Enhanced cleanup: Check if the original folder is empty or can be safely deleted
          try:
               self._cleanup_original_folder(folder_path, new_folder_path)
          except Exception as e:
               self.log(f"Error during folder cleanup: {e}")

          return new_folder_path

     # def sort_ind_folder(self, folder_path, reinject_list, order_key):
     #      """Sort all files in a BioI folder using batch processing"""
     #      self.log(f"Processing folder: {folder_path}")

     #      # Store reinject lists for use in methods
     #      self.reinject_list = reinject_list
     #      self.raw_reinject_list = getattr(self, 'raw_reinject_list', reinject_list)

     #      # Extract I number from the folder
     #      i_num = self.file_dao.get_inumber_from_name(folder_path)

     #      # If no I-number found, try to find it from the ab1 files
     #      if not i_num:
     #           ab1_files = self.file_dao.get_files_by_extension(folder_path, ".ab1")
     #           for file_path in ab1_files:
     #                parent_dir = os.path.basename(os.path.dirname(file_path))
     #                i_num = self.file_dao.get_inumber_from_name(parent_dir)
     #                if i_num:
     #                     self.log(f"Found I number {i_num} from parent directory of AB1 file")
     #                     break

     #      # Create or find the target BioI folder - always use clean format
     #      if i_num:
     #           new_folder_path = self._get_bioi_folder_path(i_num, os.path.dirname(folder_path))
     #      else:
     #           # If no I number found, use the original folder
     #           new_folder_path = folder_path
     #           self.log(f"No I number found, using original folder: {folder_path}")

     #      # Get all .ab1 files in the folder
     #      ab1_files = self.file_dao.get_files_by_extension(folder_path, ".ab1")
     #      self.log(f"Found {len(ab1_files)} .ab1 files in folder")

     #      # Group files by type for batch processing
     #      pcr_files = {}
     #      control_files = []
     #      blank_files = []
     #      customer_files = []
     #      unmatched_files = []

     #      # Classify files first
     #      for file_path in ab1_files:
     #           file_name = os.path.basename(file_path)

     #           # Check for PCR files
     #           pcr_number = self.file_dao.get_pcr_number(file_name)
     #           if pcr_number:
     #                if pcr_number not in pcr_files:
     #                     pcr_files[pcr_number] = []
     #                pcr_files[pcr_number].append(file_path)
     #                continue

     #           # Check for blank files (check this before control files)
     #           if self.file_dao.is_blank_file(file_name):
     #                self.log(f"Identified blank file: {file_name}")
     #                blank_files.append(file_path)
     #                continue

     #           # Check for control files
     #           if self.file_dao.is_control_file(file_name, self.config.CONTROLS):
     #                self.log(f"Identified control file: {file_name}")
     #                control_files.append(file_path)
     #                continue

     #           # Check if it's a customer file (has a match in order key)
     #           # Pre-normalize for matching
     #           normalized_name = self.file_dao.normalize_filename(file_name)
               
     #           # Build order key index if needed
     #           if self.order_key_index is None and order_key is not None:
     #                self.build_order_key_index(order_key)
                    
     #           # Check if in order key
     #           if self.order_key_index and normalized_name in self.order_key_index:
     #                customer_files.append(file_path)
     #           else:
     #                # No match found - unmatched file
     #                unmatched_files.append(file_path)

     #      # Detailed logging
     #      self.log(
     #           f"Classified {len(pcr_files)} PCR numbers, {len(control_files)} controls, "
     #           f"{len(blank_files)} blanks, {len(customer_files)} customer files, "
     #           f"{len(unmatched_files)} unmatched files"
     #      )

     #      # Process each group
     #      # Process PCR files by PCR number
     #      for pcr_number, files in pcr_files.items():
     #           self.log(f"Processing {len(files)} files for PCR number {pcr_number}")
     #           for file_path in files:
     #                self._sort_pcr_file(file_path, pcr_number)

     #      # Process controls
     #      if control_files:
     #           self.log(f"Processing {len(control_files)} control files")
     #           controls_folder = os.path.join(new_folder_path, "Controls")
     #           if not os.path.exists(controls_folder):
     #                os.makedirs(controls_folder)

     #           for file_path in control_files:
     #                target_path = os.path.join(controls_folder, os.path.basename(file_path))
     #                moved = self.file_dao.move_file(file_path, target_path)
     #                if moved:
     #                     self.log(f"Moved control file {os.path.basename(file_path)} to {controls_folder}")
     #                else:
     #                     self.log(f"Failed to move control file {os.path.basename(file_path)}")

     #      # Process blanks
     #      if blank_files:
     #           self.log(f"Processing {len(blank_files)} blank files")
     #           blank_folder = os.path.join(new_folder_path, "Blank")
     #           if not os.path.exists(blank_folder):
     #                os.makedirs(blank_folder)

     #           for file_path in blank_files:
     #                target_path = os.path.join(blank_folder, os.path.basename(file_path))
     #                moved = self.file_dao.move_file(file_path, target_path)
     #                if moved:
     #                     self.log(f"Moved blank file {os.path.basename(file_path)} to {blank_folder}")
     #                else:
     #                     self.log(f"Failed to move blank file {os.path.basename(file_path)}")

     #      # Process customer files
     #      if customer_files:
     #           self.log(f"Processing {len(customer_files)} customer files")
     #           for file_path in customer_files:
     #                self.sort_customer_file(file_path, order_key)

     #      # Process unmatched files - move to BioI folder root
     #      if unmatched_files and i_num:
     #           self.log(f"Processing {len(unmatched_files)} unmatched files")
     #           for file_path in unmatched_files:
     #                file_name = os.path.basename(file_path)
     #                # Clean filename for destination (remove braces)
     #                clean_brace_file_name = re.sub(r'{.*?}', '', file_name)
     #                target_path = os.path.join(new_folder_path, clean_brace_file_name)
                    
     #                moved = self.file_dao.move_file(file_path, target_path)
     #                if moved:
     #                     self.log(f"Moved unmatched file {file_name} to BioI folder root")
     #                else:
     #                     self.log(f"Failed to move unmatched file {file_name}")

     #      # Enhanced cleanup: Check if the original folder is empty or can be safely deleted
     #      try:
     #           self._cleanup_original_folder(folder_path, new_folder_path)
     #      except Exception as e:
     #           self.log(f"Error during folder cleanup: {e}")

     #      return new_folder_path

     def get_destination_for_order(self, order_identifier, base_path=None):
          """
          Find the correct destination for an order based on its I-number.
          
          Args:
               order_identifier: Either an order folder path or an I-number string
               base_path: Optional base path to search from
               
          Returns:
               str: Path to the destination parent folder (not the full target path)
          """
          # Determine if we have a folder path or direct I-number
          if isinstance(order_identifier, str) and os.path.exists(order_identifier):
               # Extract I-number from folder name
               folder_name = os.path.basename(order_identifier)
               i_num = self.file_dao.get_inumber_from_name(folder_name)
          else:
               # Assume direct I-number was provided
               i_num = order_identifier
          
          if not i_num:
               self.log("Warning: Could not extract I number from identifier")
               return None
               
          self.log(f"Extracted I number: {i_num}")
          
          # Always use the standardized folder path method
          return self._get_bioi_folder_path(i_num, base_path)

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

     def _move_file_to_destination(self, file_path, destination_folder, normalized_name):
          """Handle file placement including reinject and NN folder logic"""
          file_name = os.path.basename(file_path)
          
          # First check if file is in a Not Needed folder
          is_from_nn = self.is_in_not_needed_folder(file_path)
          
          # Check if file is in reinject list
          is_reinject = False
          if hasattr(self, 'reinject_list') and self.reinject_list:
               is_reinject = normalized_name in [self.file_dao.standardize_filename_for_matching(r) for r in self.reinject_list]
          
          # Check if file is a preemptive reinject
          is_preempt = self.is_preemptive(file_name)
          
          # Clean filename for destination (remove braces)
          clean_brace_file_name = re.sub(r'{.*?}', '', file_name)
          target_file_path = os.path.join(destination_folder, clean_brace_file_name)
          
          # Determine if file should go to Alternate Injections
          go_to_alt_injections = False
          
          if is_from_nn:
               go_to_alt_injections = True
               self.log(f"File is from NN folder: {file_name}")
          elif is_reinject:
               go_to_alt_injections = True
               self.log(f"File is in reinject list: {file_name}")
          elif is_preempt and not is_reinject:
               # Look for matching files to see if this preemptive should be moved
               matching_files = self.find_matching_files(destination_folder, normalized_name)
               if matching_files:
                    go_to_alt_injections = True
                    self.log(f"Preemptive file has matching files in destination: {file_name}")
               else:
                    self.log(f"Preemptive file has no matching files, will be primary: {file_name}")
          
          # Additional check: if file already exists at destination, use Alternate Injections
          if os.path.exists(target_file_path):
               go_to_alt_injections = True
               self.log(f"File already exists at destination: {file_name}")
          
          # Move the file to the appropriate location
          if go_to_alt_injections:
               alt_inj_folder = os.path.join(destination_folder, "Alternate Injections")
               os.makedirs(alt_inj_folder, exist_ok=True)
               alt_file_path = os.path.join(alt_inj_folder, file_name)  # Keep original name with braces
               return self.file_dao.move_file(file_path, alt_file_path)
          else:
               # Put in main folder
               return self.file_dao.move_file(file_path, target_file_path)

     # def _move_file_to_destination(self, file_path, destination_folder, normalized_name):
     #      """Handle file placement including reinject logic"""
     #      file_name = os.path.basename(file_path)

     #      # Clean filename for destination (remove braces)
     #      clean_brace_file_name = re.sub(r'{.*?}', '', file_name)
     #      target_file_path = os.path.join(destination_folder, clean_brace_file_name)

     #      # Check if file is a reinject by presence in reinject list
     #      is_reinject = False
     #      if hasattr(self, 'reinject_list') and self.reinject_list:
     #           is_reinject = normalized_name in [self.file_dao.standardize_filename_for_matching(r) for r in
     #                                              self.reinject_list]
          
     #      # IMPROVEMENT: Use the pattern from config
     #      # Check for double well pattern (also triggers reinject logic)
     #      if self.config.REGEX_PATTERNS['double_well'].match(file_name):
     #           is_reinject = True
     #           self.log(f"Double well pattern detected in: {file_name}")
          
     #      self.log(f"Is reinject: {is_reinject}")

     #      # Handle file placement - same logic as before but much cleaner
     #      if os.path.exists(target_file_path):
     #           # File already exists, put in alternate injections
     #           alt_inj_folder = os.path.join(destination_folder, "Alternate Injections")
     #           os.makedirs(alt_inj_folder, exist_ok=True)  # Simplified folder creation
     #           alt_file_path = os.path.join(alt_inj_folder, file_name)
     #           return self.file_dao.move_file(file_path, alt_file_path)
     #      elif is_reinject:
     #           # Handle reinjections
     #           if is_reinject and normalized_name in [self.file_dao.standardize_filename_for_matching(r) for r in self.reinject_list]:
     #                # It's in the actual reinject list (not just double well pattern)
     #                raw_name_idx = self.reinject_list.index(normalized_name)
     #                raw_name = self.raw_reinject_list[raw_name_idx] if hasattr(self, 'raw_reinject_list') else self.reinject_list[raw_name_idx]

     #                # Check for preemptive reinject
     #                if '{!P}' in raw_name:
     #                     # Preemptive reinject goes to main folder
     #                     return self.file_dao.move_file(file_path, target_file_path)
               
     #           # Regular reinject or double well pattern goes to alternate injections
     #           alt_inj_folder = os.path.join(destination_folder, "Alternate Injections")
     #           os.makedirs(alt_inj_folder, exist_ok=True)  # Simplified folder creation
     #           alt_file_path = os.path.join(alt_inj_folder, file_name)
     #           return self.file_dao.move_file(file_path, alt_file_path)
     #      else:
     #           # Regular file, put in main folder
     #           return self.file_dao.move_file(file_path, target_file_path)

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

     def process_bio_folder(self, bio_folder):
          """Process a BioI folder containing multiple order folders"""
          folder_name = os.path.basename(bio_folder)
          self.log(f"Processing BioI folder: {folder_name}")
          
          # Get all order folders in this BioI folder
          order_folders = self.get_order_folders(bio_folder)
          
          # Process each order folder individually
          for order_folder in order_folders:
               self.process_order_folder(order_folder, bio_folder)
          
          return

     def process_order_folder(self, order_folder, parent_folder=None):
          """Process an individual order folder"""
          folder_name = os.path.basename(order_folder)
          self.log(f"Processing order folder: {folder_name}")
          
          # Determine if we're in IND Not Ready context
          in_not_ready = os.path.basename(os.path.dirname(order_folder)) == self.config.IND_NOT_READY_FOLDER
          
          # Get order information
          order_number = self.get_order_number_from_folder_name(order_folder)
          ab1_files = self.file_dao.get_files_by_extension(order_folder, self.config.ABI_EXTENSION)
          expected_count = self._get_expected_file_count(order_number)
          is_andreev_order = self.config.ANDREEV_NAME in folder_name.lower()
          
          # Check order completeness
          files_complete = expected_count > 0 and len(ab1_files) == expected_count
          
          # Check order status
          was_mseqed, has_braces, has_ab1_files = self.check_order_status(order_folder)
          
          # Handle based on status and order type
          if not was_mseqed and not has_braces:
               if files_complete and has_ab1_files:
                    # Process with mSeq if complete (but skip for Andreev orders)
                    if not is_andreev_order:
                         self.ui_automation.process_folder(order_folder)
                         self.log(f"mSeq completed: {folder_name}")
                         
                         # IMPORTANT: Close mSeq to release file handles before moving
                         if self.ui_automation:
                              self.ui_automation.close()
                              time.sleep(0.5)  # Wait for handles to be released
                    
                    # If in IND Not Ready, move back to appropriate location
                    if in_not_ready:
                         # Get parent BioI folder path
                         bioi_folder = self.get_destination_for_order(order_folder, parent_folder)
                         
                         if bioi_folder:
                              # Construct the destination path
                              destination = os.path.join(bioi_folder, folder_name)
                              
                              # Check if destination exists
                              if os.path.exists(destination):
                                   self.log(f"Destination already exists: {destination}, leaving in current location")
                              else:
                                   # Use direct shutil.move to avoid nested directory issues
                                   try:
                                        import shutil
                                        shutil.move(order_folder, destination)
                                        self.log(f"Order moved back: {folder_name}")
                                   except Exception as e:
                                        self.warning(f"Error moving folder: {e}")
                                        self.log(f"Failed to move folder, leaving in current location: {folder_name}")
                         else:
                              self.log(f"Could not determine destination BioI folder, leaving in current location: {folder_name}")
               else:
                    # Not complete - move to IND Not Ready
                    if not in_not_ready:
                         not_ready_path = os.path.join(os.path.dirname(parent_folder or order_folder), 
                                                  self.config.IND_NOT_READY_FOLDER)
                         self.file_dao.create_folder_if_not_exists(not_ready_path)
                         
                         target_path = os.path.join(not_ready_path, folder_name)
                         if os.path.exists(target_path):
                              self.log(f"Destination already exists in IND Not Ready: {target_path}")
                         else:
                              try:
                                   import shutil
                                   shutil.move(order_folder, target_path)
                                   self.log(f"Incomplete order moved to Not Ready: {folder_name}")
                              except Exception as e:
                                   self.warning(f"Error moving to IND Not Ready: {e}")
                                   self.log(f"Failed to move folder to IND Not Ready: {folder_name}")
          
          # Already mSeqed but in IND Not Ready - move back
          elif was_mseqed and in_not_ready:
               # Close mSeq to release file handles before moving
               if self.ui_automation:
                    self.ui_automation.close()
                    time.sleep(0.5)  # Wait for handles to be released
                    
               # Get parent BioI folder path
               bioi_folder = self.get_destination_for_order(order_folder, parent_folder)
               
               if bioi_folder:
                    # Construct the destination path
                    destination = os.path.join(bioi_folder, folder_name)
                    
                    # Check if destination exists
                    if os.path.exists(destination):
                         self.log(f"Destination already exists: {destination}, leaving in current location")
                    else:
                         # Use direct shutil.move to avoid nested directory issues
                         try:
                              import shutil
                              shutil.move(order_folder, destination)
                              self.log(f"Processed order moved back: {folder_name}")
                         except Exception as e:
                              self.warning(f"Error moving folder: {e}")
                              self.log(f"Failed to move folder, leaving in current location: {folder_name}")
               else:
                    self.log(f"Could not determine destination BioI folder, leaving in current location: {folder_name}")

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
          """Process all files in a PCR folder - updated matching logic"""
          self.log(f"Processing PCR folder: {os.path.basename(pcr_folder)}")

          for file in os.listdir(pcr_folder):
               if file.endswith(config.ABI_EXTENSION):
                    file_path = os.path.join(pcr_folder, file)
                    
                    # Check if file is in NN folder
                    is_from_nn = self.is_in_not_needed_folder(pcr_folder)
                    
                    # Clean the filename for matching with reinject list
                    clean_name = self.file_dao.standardize_filename_for_matching(file)

                    # Check if it's in the reinject list
                    is_reinject = clean_name in self.reinject_list
                    
                    # Check if it's a preemptive reinject
                    is_preempt = self.is_preemptive(file)
                    
                    # Determine if file should go to Alternate Injections
                    go_to_alt_injections = is_from_nn or is_reinject
                    
                    if is_preempt and not is_reinject:
                         # Check for other files with matching name
                         matching_files = self.find_matching_files(pcr_folder, clean_name)
                         # Remove the current file from consideration
                         matching_files = [f for f in matching_files if f != file]
                         if matching_files:
                              go_to_alt_injections = True

                    if go_to_alt_injections:
                         # Make sure Alternate Injections folder exists
                         alt_folder = os.path.join(pcr_folder, "Alternate Injections")
                         if not os.path.exists(alt_folder):
                              os.mkdir(alt_folder)

                         # Move file to Alternate Injections
                         source = os.path.join(pcr_folder, file)
                         dest = os.path.join(alt_folder, file)  # Keep original filename with brackets
                         if os.path.exists(source) and not os.path.exists(dest):
                              self.file_dao.move_file(source, dest)
                              reason = ""
                              if is_from_nn: reason += "from NN folder "
                              if is_reinject: reason += "in reinject list "
                              if is_preempt and not is_reinject and matching_files: reason += "preemptive with match"
                              self.log(f"Moved file {file} to Alternate Injections ({reason.strip()})")

          # Now process the folder with mSeq if it's ready
          was_mseqed, has_braces, has_ab1_files = self.check_order_status(pcr_folder)
          if not was_mseqed and not has_braces and has_ab1_files:
               self.ui_automation.process_folder(pcr_folder)
               self.log(f"mSeq completed: {os.path.basename(pcr_folder)}")
          else:
               self.log(f"mSeq NOT completed: {os.path.basename(pcr_folder)}")
     # def process_pcr_folder(self, pcr_folder):
     #      """Process all files in a PCR folder - simpler matching logic"""
     #      self.log(f"Processing PCR folder: {os.path.basename(pcr_folder)}")

     #      for file in os.listdir(pcr_folder):
     #           if file.endswith(config.ABI_EXTENSION):
     #                # Clean the filename for matching with reinject list
     #                clean_name = self.file_dao.standardize_filename_for_matching(file)

     #                # Check if it's in the reinject list
     #                is_reinject = clean_name in self.reinject_list

     #                if is_reinject:
     #                     # Make sure Alternate Injections folder exists
     #                     alt_folder = os.path.join(pcr_folder, "Alternate Injections")
     #                     if not os.path.exists(alt_folder):
     #                          os.mkdir(alt_folder)

     #                     # Move file to Alternate Injections
     #                     source = os.path.join(pcr_folder, file)
     #                     dest = os.path.join(alt_folder, file)  # Keep original filename with brackets
     #                     if os.path.exists(source) and not os.path.exists(dest):
     #                          self.file_dao.move_file(source, dest)
     #                          self.log(f"Moved reinject {file} to Alternate Injections")

     #      # Now process the folder with mSeq if it's ready
     #      was_mseqed, has_braces, has_ab1_files = self.check_order_status(pcr_folder)
     #      if not was_mseqed and not has_braces and has_ab1_files:
     #           self.ui_automation.process_folder(pcr_folder)
     #           self.log(f"mSeq completed: {os.path.basename(pcr_folder)}")
     #      else:
     #           self.log(f"mSeq NOT completed: {os.path.basename(pcr_folder)}")

     def _cleanup_original_folder(self, original_folder: str, new_folder: str):
          """
          Enhanced cleanup to remove the original folder if all files have been processed,
          handle nested NN folders, and clean up empty folders
          """
          if original_folder == new_folder:
               self.log("Original folder is the same as new folder, no cleanup needed")
               return

          # Process nested NN folders first
          for root, dirs, files in os.walk(original_folder, topdown=False):  # Process bottom-up
               for dir_name in dirs:
                    dir_path = os.path.join(root, dir_name)
                    dir_lower = dir_name.lower()
                    
                    # Check if this is an NN folder
                    if dir_lower == "nn" or dir_lower == "nn-preemptives" or dir_lower.startswith("nn_") or dir_lower.startswith("nn-preemptives_"):
                         self.log(f"Processing nested NN folder: {dir_path}")
                         
                         # Get all AB1 files in this NN folder
                         nn_ab1_files = self.file_dao.get_files_by_extension(dir_path, ".ab1", recursive=True)
                         
                         if nn_ab1_files:
                              self.log(f"Found {len(nn_ab1_files)} AB1 files in nested NN folder")
                              
                              # Process each file
                              for file_path in nn_ab1_files:
                                   file_name = os.path.basename(file_path)
                                   normalized_name = self.file_dao.normalize_filename(file_name)
                                   
                                   # Find matching entry in order key (if available)
                                   if self.order_key_index and normalized_name in self.order_key_index:
                                        matches = self.order_key_index[normalized_name]
                                        i_num, acct_name, order_num = matches[0]  # Use first match
                                        
                                        # Create destination folder and move file
                                        destination_folder = self.create_order_folder(i_num, acct_name, order_num)
                                        self._move_file_to_destination(file_path, destination_folder, normalized_name)
                                   else:
                                        self.log(f"No order key match for nested NN file: {file_name}")
                              
                         # After processing files, check if NN folder is now empty
                         remaining_nn_items = os.listdir(dir_path)
                         if not remaining_nn_items:
                              try:
                                   os.rmdir(dir_path)
                                   self.log(f"Deleted empty NN folder: {dir_path}")
                              except Exception as e:
                                   self.log(f"Failed to delete empty NN folder {dir_path}: {e}")

          # Force refresh directory cache after processing NN folders
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
               elif os.path.isdir(item_path):
                    # Check if this is a non-standard folder (like NN) and if it's empty
                    dir_lower = item.lower()
                    if (dir_lower == "nn" or dir_lower == "nn-preemptives" or 
                         dir_lower.startswith("nn_") or dir_lower.startswith("nn-preemptives_")) and not os.listdir(item_path):
                         try:
                              os.rmdir(item_path)
                              self.log(f"Deleted empty NN folder: {item}")
                              # Remove from remaining items list
                              remaining_items = [i for i in remaining_items if i != item]
                              continue
                         except Exception as e:
                              self.log(f"Failed to delete empty NN folder {item}: {e}")
                              moved_all = False
                    else:
                         # Non-standard folder
                         self.log(f"Found non-standard folder in folder: {item}")
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
               # Check if all remaining items are empty NN folders
               all_empty_nn = True
               for item in remaining_items:
                    item_path = os.path.join(original_folder, str(item))
                    if not os.path.isdir(item_path):
                         all_empty_nn = False
                         break
                         
                    dir_lower = item.lower()
                    if not (dir_lower == "nn" or dir_lower == "nn-preemptives" or 
                         dir_lower.startswith("nn_") or dir_lower.startswith("nn-preemptives_")):
                         all_empty_nn = False
                         break
                         
                    if os.listdir(item_path):
                         all_empty_nn = False
                         break
               
               if all_empty_nn:
                    # Delete all empty NN folders
                    for item in list(remaining_items):
                         item_path = os.path.join(original_folder, str(item))
                         try:
                              os.rmdir(item_path)
                              self.log(f"Deleted empty NN folder: {item}")
                         except Exception as e:
                              self.log(f"Failed to delete empty NN folder {item}: {e}")
                    
                    # Try to delete the original folder again
                    try:
                         os.rmdir(original_folder)
                         self.log(f"Deleted original folder after cleaning all empty NN folders: {original_folder}")
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

          # Check if file is in a Not Needed folder
          is_from_nn = self.is_in_not_needed_folder(file_path)
          
          # Use standard normalization
          normalized_name = self.file_dao.standardize_filename_for_matching(file_name)

          # Check if file is in reinject list
          is_reinject = False
          if hasattr(self, 'reinject_list') and self.reinject_list:
               for i, item in enumerate(self.reinject_list):
                    reinj_norm = self.file_dao.standardize_filename_for_matching(item)
                    if normalized_name == reinj_norm:
                         is_reinject = True
                         raw_name = self.raw_reinject_list[i] if hasattr(self, 'raw_reinject_list') else item
                         self.log(f"Found PCR file in reinject list: {item}")
                         break
          
          # Check if file is a preemptive reinject
          is_preempt = self.is_preemptive(file_name)
          
          # Create the destination path for main folder
          clean_name = re.sub(r'{.*?}', '', file_name)
          target_path = os.path.join(pcr_folder_path, clean_name)
          
          # Determine if file should go to Alternate Injections
          go_to_alt_injections = False
          
          if is_from_nn:
               go_to_alt_injections = True
               self.log(f"PCR file is from NN folder: {file_name}")
          elif is_reinject:
               go_to_alt_injections = True 
               self.log(f"PCR file is in reinject list: {file_name}")
          elif is_preempt and not is_reinject:
               # Look for matching files
               matching_files = self.find_matching_files(pcr_folder_path, normalized_name)
               if matching_files:
                    go_to_alt_injections = True
                    self.log(f"Preemptive PCR file has matching files in destination: {file_name}")
               else:
                    self.log(f"Preemptive PCR file has no matching files, will be primary: {file_name}")
          
          # Additional check: if file already exists at destination, use Alternate Injections
          if os.path.exists(target_path):
               go_to_alt_injections = True
               self.log(f"PCR file already exists at destination: {file_name}")
          
          # Move the file to the appropriate location
          if go_to_alt_injections:
               alt_inj_folder = os.path.join(pcr_folder_path, "Alternate Injections")
               if not os.path.exists(alt_inj_folder):
                    os.makedirs(alt_inj_folder)
               alt_file_path = os.path.join(alt_inj_folder, file_name)  # Keep original name with braces
               return self.file_dao.move_file(file_path, alt_file_path)
          else:
               # Put in main folder
               return self.file_dao.move_file(file_path, target_path)

     # def _sort_pcr_file(self, file_path, pcr_number):
     #      """Sort a PCR file to the appropriate folder"""
     #      file_name = os.path.basename(file_path)
     #      self.log(f"Processing PCR file: {file_name} with PCR Number: {pcr_number}")

     #      # Get the day data folder
     #      day_data_path = os.path.dirname(os.path.dirname(file_path))

     #      # Create PCR folder name and path
     #      pcr_folder_path = self.get_pcr_folder_path(pcr_number, day_data_path)

     #      # Create PCR folder if it doesn't exist
     #      if not os.path.exists(pcr_folder_path):
     #           os.makedirs(pcr_folder_path)

     #      # Special handling for PCR files
     #      # Use standard normalization first (which removes ab1 extension)
     #      normalized_name = self.file_dao.standardize_filename_for_matching(file_name)

     #      # Debug output for troubleshooting
     #      self.log(f"PCR file normalized name: {normalized_name}")

     #      # NEW: Check if file is a reinject by double well pattern
     #      is_reinject = bool(self.config.REGEX_PATTERNS['double_well'].match(file_name))
          
     #      # For completeness, also check against reinject list
     #      raw_name = None
     #      if not is_reinject and hasattr(self, 'reinject_list') and self.reinject_list:
     #           for i, item in enumerate(self.reinject_list):
     #                reinj_norm = self.file_dao.standardize_filename_for_matching(item)
     #                if normalized_name == reinj_norm:
     #                     is_reinject = True
     #                     raw_name = self.raw_reinject_list[i] if hasattr(self, 'raw_reinject_list') else item
     #                     self.log(f"Found PCR file in reinject list: {item}")
     #                     break

     #      # Create the destination path for main folder
     #      clean_name = re.sub(r'{.*?}', '', file_name)
     #      target_path = os.path.join(pcr_folder_path, clean_name)

     #      # Handle reinjects
     #      if is_reinject:
     #           # Check for preemptive reinject marker
     #           if raw_name and '{!P}' in raw_name:
     #                # Preemptive reinject goes in main folder, but check if file exists first
     #                if os.path.exists(target_path):
     #                     # Create Alternate Injections folder instead to avoid overwriting
     #                     alt_inj_folder = os.path.join(pcr_folder_path, "Alternate Injections")
     #                     if not os.path.exists(alt_inj_folder):
     #                          os.makedirs(alt_inj_folder)
     #                     alt_file_path = os.path.join(alt_inj_folder, file_name)  # Keep original name with braces
     #                     self.log(f"Preemptive reinject moved to alternate injections (to avoid overwrite)")
     #                     return self.file_dao.move_file(file_path, alt_file_path)
     #                else:
     #                     # No conflict, move to main folder
     #                     self.log(f"Preemptive reinject moved to main folder")
     #                     return self.file_dao.move_file(file_path, target_path)
     #           else:
     #                # Regular reinject always goes to alternate injections
     #                alt_inj_folder = os.path.join(pcr_folder_path, "Alternate Injections")
     #                if not os.path.exists(alt_inj_folder):
     #                     os.makedirs(alt_inj_folder)
     #                alt_file_path = os.path.join(alt_inj_folder, file_name)  # Keep original name with braces
     #                self.log(f"Reinject moved to alternate injections")
     #                return self.file_dao.move_file(file_path, alt_file_path)
     #      else:
     #           # Regular file goes in main folder, but check if it already exists
     #           if os.path.exists(target_path):
     #                # File with same name exists - might be an undetected reinject or a duplicate
     #                # Move to alternate injections to be safe
     #                alt_inj_folder = os.path.join(pcr_folder_path, "Alternate Injections")
     #                if not os.path.exists(alt_inj_folder):
     #                     os.makedirs(alt_inj_folder)
     #                alt_file_path = os.path.join(alt_inj_folder, file_name)  # Keep original name with braces
     #                self.log(f"File with same name already exists in main folder, moved to alternate injections")
     #                return self.file_dao.move_file(file_path, alt_file_path)
     #           else:
     #                # No conflict, move to main folder
     #                self.log(f"File moved to main folder")
     #                return self.file_dao.move_file(file_path, target_path)

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

     def get_order_number_from_folder_name(self, folder_path):
          """Extract order number from folder name"""
          folder_name = os.path.basename(folder_path)
          match = re.search(r'_\d+$', folder_name)
          if match:
               order_number = re.search(r'\d+', match.group(0)).group(0)
               return order_number
          return None

     def is_in_not_needed_folder(self, file_path):
          """
          Check if file is in a Not Needed folder at any level of nesting
          """
          # Get the path components
          path_parts = file_path.replace('\\', '/').split('/')
          
          # Check if any directory in the path is an NN folder
          for part in path_parts:
               part_lower = part.lower()
               if part_lower == "nn" or part_lower == "nn-preemptives" or part_lower.startswith("nn_") or part_lower.startswith("nn-preemptives_"):
                    self.log(f"File in NN folder: {file_path}, detected folder: {part}")
                    return True
          
          return False

     def is_preemptive(self, file_name):
          """Check if file has double well pattern indicating it's a preemptive reinject"""
          return bool(self.config.REGEX_PATTERNS['double_well'].match(file_name))

     def find_matching_files(self, folder_path, normalized_name):
          """Find files in folder that match the normalized name"""
          matching_files = []
          for item in os.listdir(folder_path):
               if item.endswith(self.config.ABI_EXTENSION):
                    item_norm = self.file_dao.normalize_filename(item)
                    if item_norm == normalized_name:
                         matching_files.append(item)
          return matching_files

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

     def debug_reinject_detection(self, file_path):
          """Debug helper for all file sorting logic"""
          file_name = os.path.basename(file_path)
          
          # Check all conditions
          is_from_nn = self.is_in_not_needed_folder(file_path)
          
          # Check if in reinject list
          normalized_name = self.file_dao.standardize_filename_for_matching(file_name)
          is_reinject = False
          if hasattr(self, 'reinject_list') and self.reinject_list:
               normalized_reinjects = [self.file_dao.standardize_filename_for_matching(r) for r in self.reinject_list]
               is_reinject = normalized_name in normalized_reinjects
          
          # Check for preemptive pattern
          is_preempt = self.is_preemptive(file_name)
          
          # Determine destination
          go_to_alt_injections = is_from_nn or is_reinject
          if is_preempt and not is_reinject:
               # Would need to check for matching files in actual destination folder
               go_to_alt_injections = "NEEDS DESTINATION FOLDER CHECK"
          
          debug_info = {
               'file_name': file_name,
               'normalized_name': normalized_name,
               'is_from_nn': is_from_nn,
               'is_reinject': is_reinject,
               'is_preemptive': is_preempt,
               'likely_destination': "Alternate Injections" if go_to_alt_injections else "Main Folder"
          }
          
          return debug_info

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
          Zip the contents of an order folder with special handling for Andreev orders
          
          Args:
               folder_path (str): Path to the order folder
               include_txt (bool): Whether to include text files in zip
               
          Returns:
               str: Path to created zip file, or None if failed
          """
          try:
               folder_name = os.path.basename(folder_path)
               is_andreev_order = self.config.ANDREEV_NAME.lower() in folder_name.lower()
               
               # For Andreev orders, use different naming and never include txt files
               if is_andreev_order:
                    # Extract order number and I-number from folder name
                    order_number = self.get_order_number_from_folder_name(folder_path)
                    i_number = self.file_dao.get_inumber_from_name(folder_name)
                    
                    if not order_number or not i_number:
                         self.log(f"Could not extract order number or I-number for Andreev order: {folder_name}")
                         return None
                         
                    # Use Andreev's preferred naming format: "123456_I-20000.zip"
                    zip_filename = f"{order_number}_I-{i_number}.zip"
                    include_txt = False  # Never include txt files for Andreev
               else:
                    # Use standard naming format: "BioI-20000_Customer_123456.zip"
                    zip_filename = f"{folder_name}.zip"
               
               zip_path = os.path.join(folder_path, zip_filename)
               
               # Determine which files to include
               file_extensions = [self.config.ABI_EXTENSION]
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

     def _try_delete_if_empty(self, folder_path, depth=0):
          """
          Recursively check if a folder is empty or contains only empty folders,
          and delete it if so.
          
          Args:
               folder_path (str): Path to the folder to check
               depth (int): Current recursion depth (for logging)
               
          Returns:
               bool: True if the folder was empty or successfully deleted, False otherwise
          """
          if not os.path.exists(folder_path):
               return True
               
          # Get list of items in the folder
          try:
               items = os.listdir(folder_path)
          except Exception as e:
               self.log(f"Error listing directory {folder_path}: {e}")
               return False
               
          # If folder is already empty, try to delete it
          if not items:
               try:
                    os.rmdir(folder_path)
                    self.log(f"Deleted empty folder in final cleanup: {folder_path}")
                    return True
               except Exception as e:
                    self.log(f"Failed to delete empty folder {folder_path}: {e}")
                    return False
                    
          # Check if all items are directories that can be deleted
          all_removable = True
          for item in items:
               item_path = os.path.join(folder_path, item)
               if os.path.isdir(item_path):
                    # Recursively check if this subdirectory can be deleted
                    if not self._try_delete_if_empty(item_path, depth + 1):
                         all_removable = False
               else:
                    # It's a file, so this directory cannot be removed
                    all_removable = False
                    
          # If all subdirectories were removed, try to delete this directory too
          if all_removable:
               # Check again if the folder is now empty
               if not os.listdir(folder_path):
                    try:
                         os.rmdir(folder_path)
                         self.log(f"Deleted empty folder in final cleanup: {folder_path}")
                         return True
                    except Exception as e:
                         self.log(f"Failed to delete empty folder {folder_path}: {e}")
                         return False
                         
          return False

     def final_cleanup(self, root_folder):
          """
          Final cleanup pass to delete any empty folders that might have been missed.
          
          Args:
               root_folder (str): Root folder to start cleanup from
          """
          self.log(f"Running final cleanup on: {root_folder}")
          
          # Get all folders in the root directory
          folders_to_check = []
          for item in os.listdir(root_folder):
               item_path = os.path.join(root_folder, item)
               if os.path.isdir(item_path):
                    folders_to_check.append(item_path)
          
          # Process each folder to see if it's empty or contains only empty folders
          for folder in folders_to_check:
               self._try_delete_if_empty(folder)
          
          self.log("Final cleanup complete")

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
        i_numbers, _ = file_dao.get_folders_with_inumbers(folder_path)
        print(f"Found I numbers: {i_numbers}")

        # Get reinject list
        reinject_list = processor.get_reinject_list(i_numbers)

        # Print the results
        print(f"\nReinject List ({len(reinject_list)} entries):")
        for i, item in enumerate(reinject_list):
            print(f"{i + 1}. {item}")