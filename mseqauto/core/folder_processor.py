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

config = MseqConfig()

class FolderProcessor:
     def __init__(self, file_dao, ui_automation, config, logger=None):
          self.file_dao = file_dao
          self.ui_automation = ui_automation
          self.config = config
          # Add AB1Processor
          self.ab1_processor = AB1Processor(config)

          # Flag to control whether to use direct processing or UI automation
          self.use_direct_processing = True  # Set default here or from config

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

          # Get parent folder path - this is the BioI-nnnnn folder
          parent_folder = self.get_destination_for_order(i_num, base_path)
          self.log(f"Parent folder: {parent_folder}")

          # No need to create BioI folder - it was already created by get_destination_for_order
          
          # Create full order folder path inside BioI folder
          order_folder_path = os.path.join(parent_folder, order_folder_name)
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
          
          # Determine search path based on context
          search_path = None
          
          # Check if we're in IND Not Ready context
          ind_not_ready_context = False
          
          # Check if the order folder is directly in IND Not Ready
          if isinstance(order_identifier, str) and os.path.exists(order_identifier):
               parent_dir = os.path.dirname(order_identifier)
               if os.path.basename(parent_dir) == self.config.IND_NOT_READY_FOLDER:
                    ind_not_ready_context = True
                    # Use the parent of IND Not Ready as our search path
                    search_path = os.path.dirname(parent_dir)
                    self.log(f"Order is in IND Not Ready, using parent folder: {search_path}")
          
          # Check if base_path is IND Not Ready folder
          if base_path and self.config.IND_NOT_READY_FOLDER in os.path.basename(base_path):
               ind_not_ready_context = True
               # Use the parent of IND Not Ready as our search path
               search_path = os.path.dirname(base_path)
               self.log(f"Base path is IND Not Ready, using parent folder: {search_path}")
          
          # If not in IND Not Ready context, follow normal path resolution
          if not ind_not_ready_context:
               # 1. Use explicitly provided base_path if available
               if base_path and os.path.exists(base_path):
                    search_path = base_path
                    self.log(f"Using provided base path: {search_path}")
               
               # 2. Use current_data_folder attribute if it exists
               if not search_path:
                    current_data_folder = getattr(self, 'current_data_folder', None)
                    if current_data_folder and os.path.exists(current_data_folder):
                         search_path = current_data_folder
                         self.log(f"Using current data folder: {search_path}")
               
               # 3. Fall back to today's date folder
               if not search_path:
                    try:
                         from datetime import datetime
                         today = datetime.now().strftime('%m.%d.%y')
                         preferred_path = os.path.join('P:', 'Data', today)
                         
                         # Create today's folder if it doesn't exist
                         if not os.path.exists(preferred_path):
                              try:
                                   os.makedirs(preferred_path)
                                   self.log(f"Created today's data folder: {preferred_path}")
                                   search_path = preferred_path
                              except Exception as e:
                                   self.log(f"Error creating preferred path {preferred_path}: {e}")
                         else:
                              search_path = preferred_path
                    except Exception as e:
                         self.log(f"Error using preferred date path: {e}")
                         
          # By this point, we should have a search_path determined
          if not search_path:
               # Ultimate fallback - current working directory
               search_path = os.path.dirname(os.getcwd())
               self.log(f"Using fallback search path: {search_path}")
          
          # Now search for matching BioI folder in the determined search path
          self.log(f"Searching for matching I number folder in: {search_path}")
          
          # Look for existing BioI folder in the search path
          for item in os.listdir(search_path):
               item_path = os.path.join(search_path, item)
               if os.path.isdir(item_path) and re.search(f"bioi-{i_num}", item.lower()):
                    self.log(f"Found matching folder: {item_path}")
                    return item_path
          
          # If no matching folder found, create a new one in the search path
          new_folder = os.path.join(search_path, f"BioI-{i_num}")
          try:
               if not os.path.exists(new_folder):
                    os.makedirs(new_folder)
                    self.log(f"Created new folder: {new_folder}")
               return new_folder
          except Exception as e:
               self.log(f"Error creating folder {new_folder}: {e}")
               return None

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

          # Skip Andreev's orders for mSeq processing - same as original
          if is_andreev_order in order_folder.lower():
               return
          
          # Check if we should use direct processing or UI automation
          if self.use_direct_processing:
               # Use AB1Processor for direct processing
               was_processed, has_braces, has_ab1_files = self.ab1_processor.check_process_status(order_folder)
          else:
               # Use original mSeq detection approach
               was_processed, has_braces, has_ab1_files = self.check_order_status(order_folder)

          # Process based on status - similar logic to original
          if not was_processed and not has_braces:
               expected_count = self._get_expected_file_count(order_number)
               
               if expected_count > 0 and len(ab1_files) == expected_count and has_ab1_files:
                    if self.use_direct_processing:
                         # Process directly with AB1Processor
                         success = self.ab1_processor.process_folder(order_folder)
                         if success:
                              self.log(f"Direct processing completed: {os.path.basename(order_folder)}")
                    else:
                         # Use UI automation as before
                         self.ui_automation.process_folder(order_folder)
                         self.log(f"mSeq completed: {os.path.basename(order_folder)}")
                    
                    # Handle IND Not Ready folder - same logic as original
                    if os.path.basename(os.path.dirname(order_folder)) == self.config.IND_NOT_READY_FOLDER:
                         destination = self.get_destination_for_order(order_folder, parent_folder)
                         self.file_dao.move_folder(order_folder, destination)
               else:
                    # Move to Not Ready folder if incomplete - same logic as original
                    not_ready_path = os.path.join(os.path.dirname(parent_folder), 
                                                  self.config.IND_NOT_READY_FOLDER)
                    self.file_dao.create_folder_if_not_exists(not_ready_path)
                    self.file_dao.move_folder(order_folder, not_ready_path)
                    self.log(f"Order moved to Not Ready: {os.path.basename(order_folder)}")

          # Check order completeness
          files_complete = expected_count > 0 and len(ab1_files) == expected_count
          
          # Check order status
          was_processed, has_braces, has_ab1_files = self.check_order_status(order_folder)
          
          # Handle based on status and order type
          if not was_processed and not has_braces:
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
          elif was_processed and in_not_ready:
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
          """Process a PCR folder using either direct processing or UI automation"""
          self.log(f"Processing PCR folder: {os.path.basename(pcr_folder)}")
          
          # Handle reinjects first - same logic as original
          for file in os.listdir(pcr_folder):
               if file.endswith(self.config.ABI_EXTENSION):
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
          
          # Check order status
          if self.use_direct_processing:
               # Use AB1Processor for direct processing
               was_processed, has_braces, has_ab1_files = self.ab1_processor.check_process_status(pcr_folder)
          else:
               # Use original mSeq detection approach
               was_processed, has_braces, has_ab1_files = self.check_order_status(pcr_folder)
          
          # Process the folder if needed
          if not was_processed and not has_braces and has_ab1_files:
               if self.use_direct_processing:
                    # Process directly with AB1Processor
                    success = self.ab1_processor.process_folder(pcr_folder)
                    if success:
                         self.log(f"Direct processing completed: {os.path.basename(pcr_folder)}")
               else:
                    # Use UI automation as before
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

          # NEW: Check if file is a reinject by looking for double well location pattern
          # Pattern: {XXX}{YYY}Sample_Name{PCR####exp#}{dilution} 
          is_reinject = bool(re.match(r'^{\d+[A-Z]}{\d+[A-Z]}', file_name))
          
          if is_reinject:
               self.log(f"Identified as reinject based on double well location pattern")
          
          # For completeness, also check against reinject list
          raw_name = None
          if not is_reinject and hasattr(self, 'reinject_list') and self.reinject_list:
               for i, item in enumerate(self.reinject_list):
                    reinj_norm = self.file_dao.standardize_filename_for_matching(item)
                    if normalized_name == reinj_norm:
                         is_reinject = True
                         raw_name = self.raw_reinject_list[i] if hasattr(self, 'raw_reinject_list') else item
                         self.log(f"Found PCR file in reinject list: {item}")
                         break

          # Create the destination path for main folder
          clean_name = re.sub(r'{.*?}', '', file_name)
          target_path = os.path.join(pcr_folder_path, clean_name)

          # Handle reinjects
          if is_reinject:
               # Check for preemptive reinject marker
               if raw_name and '{!P}' in raw_name:
                    # Preemptive reinject goes in main folder, but check if file exists first
                    if os.path.exists(target_path):
                         # Create Alternate Injections folder instead to avoid overwriting
                         alt_inj_folder = os.path.join(pcr_folder_path, "Alternate Injections")
                         if not os.path.exists(alt_inj_folder):
                              os.makedirs(alt_inj_folder)
                         alt_file_path = os.path.join(alt_inj_folder, file_name)  # Keep original name with braces
                         self.log(f"Preemptive reinject moved to alternate injections (to avoid overwrite)")
                         return self.file_dao.move_file(file_path, alt_file_path)
                    else:
                         # No conflict, move to main folder
                         self.log(f"Preemptive reinject moved to main folder")
                         return self.file_dao.move_file(file_path, target_path)
               else:
                    # Regular reinject always goes to alternate injections
                    alt_inj_folder = os.path.join(pcr_folder_path, "Alternate Injections")
                    if not os.path.exists(alt_inj_folder):
                         os.makedirs(alt_inj_folder)
                    alt_file_path = os.path.join(alt_inj_folder, file_name)  # Keep original name with braces
                    self.log(f"Reinject moved to alternate injections")
                    return self.file_dao.move_file(file_path, alt_file_path)
          else:
               # Regular file goes in main folder, but check if it already exists
               if os.path.exists(target_path):
                    # File with same name exists - might be an undetected reinject or a duplicate
                    # Move to alternate injections to be safe
                    alt_inj_folder = os.path.join(pcr_folder_path, "Alternate Injections")
                    if not os.path.exists(alt_inj_folder):
                         os.makedirs(alt_inj_folder)
                    alt_file_path = os.path.join(alt_inj_folder, file_name)  # Keep original name with braces
                    self.log(f"File with same name already exists in main folder, moved to alternate injections")
                    return self.file_dao.move_file(file_path, alt_file_path)
               else:
                    # No conflict, move to main folder
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

     def process_pcr_folder(self, pcr_folder):
          """Process a PCR folder using either direct processing or UI automation"""
          self.log(f"Processing PCR folder: {os.path.basename(pcr_folder)}")
          
          # Handle reinjects first - same logic as original
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
          
          # Check order status
          if self.use_direct_processing:
               # Use AB1Processor for direct processing
               was_processed, has_braces, has_ab1_files = self.ab1_processor.check_process_status(pcr_folder)
          else:
               # Use original mSeq detection approach
               was_processed, has_braces, has_ab1_files = self.check_order_status(pcr_folder)
          
          # Process the folder if needed
          if not was_processed and not has_braces and has_ab1_files:
               if self.use_direct_processing:
                    # Process directly with AB1Processor
                    success = self.ab1_processor.process_folder(pcr_folder)
                    if success:
                         self.log(f"Direct processing completed: {os.path.basename(pcr_folder)}")
               else:
                    # Use UI automation as before
                    self.ui_automation.process_folder(pcr_folder)
                    self.log(f"mSeq completed: {os.path.basename(pcr_folder)}")
          else:
               self.log(f"mSeq NOT completed: {os.path.basename(pcr_folder)}")

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