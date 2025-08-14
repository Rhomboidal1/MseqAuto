# folder_processor.py
import re
# Add parent directory to PYTHONPATH for imports
import os
import sys
import numpy as np
from pathlib import Path
sys.path.append(str(Path(__file__).parents[2]))

#print(sys.path)
import time
from datetime import datetime

from mseqauto.config import MseqConfig # type: ignore
import warnings
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

config = MseqConfig()

class FolderProcessor:
     def __init__(self, file_dao, ui_automation, config, logger=None):
          self.file_dao = file_dao
          self.ui_automation = ui_automation
          self.config = config

          # Create a unified logging interface with support for different levels
          import logging
          if logger is None:
               # No logger provided, use default
               self._logger = logging.getLogger(__name__)
               self.log = self._logger.info
               self.debug = self._logger.debug
               self.warning = self._logger.warning
               self.error = self._logger.error
          elif isinstance(logger, logging.Logger):
               # It's a Logger object
               self._logger = logger
               self.log = logger.info
               self.debug = logger.debug
               self.warning = logger.warning
               self.error = logger.error
          else:
               # Assume it's a callable function - in this case, we don't have level differentiation
               self._logger = None
               self.log = logger
               self.debug = logger
               self.warning = logger
               self.error = logger

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
               # Use customer-specific normalization for order key
               normalized_name = self.file_dao.standardize_for_customer_files(sample_name, remove_extension=False)

               # Create entry in index (handle multiple entries with same normalized name)
               if normalized_name not in self.order_key_index:
                    self.order_key_index[normalized_name] = []
               self.order_key_index[normalized_name].append((i_num, acct_name, order_num))

          self.log(f"Built order key index with {len(self.order_key_index)} unique entries")

     def sort_customer_file(self, file_path, order_key):
          """Sort a customer file based on order key using the index and considering embedded order numbers"""
          # Build index if not already done
          if self.order_key_index is None:
               self.build_order_key_index(order_key)

          file_name = Path(file_path).name
          self.log(f"Processing customer file: {file_name}")

          # Extract order number from filename if present
          embedded_order_number = self.file_dao.extract_order_number_from_filename(file_name)

          # Use customer-specific normalization (removes well locations completely)
          base_normalized_name = self.file_dao.standardize_for_customer_files(file_name, remove_extension=True)

          if base_normalized_name in self.order_key_index:
               matches = self.order_key_index[base_normalized_name] #type: ignore

               # First, try to find a match with the embedded order number
               if embedded_order_number:
                    for i_num, acct_name, order_num in matches:
                         if order_num == embedded_order_number:
                              self.log(f"Found exact match with embedded order number: {embedded_order_number}")
                              destination_folder = self.create_order_folder(i_num, acct_name, order_num)
                              return self._place_customer_file(file_path, destination_folder, base_normalized_name)

               # Prioritize matches from current folder's I number
               current_i_num = self.file_dao.get_inumber_from_name(str(Path(file_path).parent))

               for i_num, acct_name, order_num in matches:
                    # Prioritize current I number if available
                    if current_i_num and i_num == current_i_num:
                         destination_folder = self.create_order_folder(i_num, acct_name, order_num)
                         return self._place_customer_file(file_path, destination_folder, base_normalized_name)

               # If no match with current I number, use the first match
               i_num, acct_name, order_num = matches[0]
               destination_folder = self.create_order_folder(i_num, acct_name, order_num)
               return self._place_customer_file(file_path, destination_folder, base_normalized_name)

          # No match found
          self.log(f"No match found in order key for: {base_normalized_name}")
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

          # Get parent folder path with consistent naming scheme
          parent_folder = Path(self._get_bioi_folder_path(i_num, base_path))
          self.log(f"Parent folder: {parent_folder}")

          # Create full order folder path inside BioI folder
          order_folder_path = parent_folder / order_folder_name
          self.log(f"Order folder path: {order_folder_path}")

          # Create order folder if it doesn't exist
          if not order_folder_path.exists():
               order_folder_path.mkdir(parents=True, exist_ok=True)
               self.log(f"Created order folder: {order_folder_path}")

          return str(order_folder_path)

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
               data_folder_path = Path('P:') / 'Data' / today
               try:
                    if not data_folder_path.exists():
                         data_folder_path.mkdir(parents=True, exist_ok=True)
                         self.log(f"Created today's data folder: {data_folder_path}")
                    data_folder = str(data_folder_path)
               except Exception as e:
                    self.log(f"Error creating date folder: {e}")
                    # Fallback to current directory
                    data_folder = str(Path.cwd().parent)

          # Create path to the standardized BioI folder
          bioi_folder_name = f"BioI-{i_num}"
          bioi_folder_path = Path(data_folder) / bioi_folder_name

          # Create the folder if it doesn't exist
          if not bioi_folder_path.exists():
               try:
                    bioi_folder_path.mkdir(parents=True, exist_ok=True)
                    self.log(f"Created clean BioI folder: {bioi_folder_path}")
               except Exception as e:
                    self.log(f"Error creating BioI folder: {e}")

          return str(bioi_folder_path)

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
                    parent_dir = Path(file_path).parent.name
                    i_num = self.file_dao.get_inumber_from_name(parent_dir)
                    if i_num:
                         self.log(f"Found I number {i_num} from parent directory of AB1 file")
                         break

          # Create or find the target BioI folder - always use clean format
          if i_num:
               new_folder_path = self._get_bioi_folder_path(i_num, str(Path(folder_path).parent))
          else:
               # If no I number found, use the original folder
               new_folder_path = folder_path
               self.log(f"No I number found, using original folder: {folder_path}")

          # Recursively get all .ab1 files in the folder and subfolders
          ab1_files = []
          for ab1_file in Path(folder_path).rglob("*.ab1"):
               ab1_files.append(str(ab1_file))

          self.log(f"Found {len(ab1_files)} .ab1 files in folder and subfolders")

          # Group files by type for batch processing
          pcr_files = {}
          control_files = []
          blank_files = []
          customer_files = []
          unmatched_files = []

          # Classify files first
          for file_path in ab1_files:
               file_name = Path(file_path).name

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
               controls_folder = Path(new_folder_path) / "Controls"
               controls_folder.mkdir(exist_ok=True)

               for file_path in control_files:
                    target_path = controls_folder / Path(file_path).name
                    moved = self.file_dao.move_file(file_path, str(target_path))
                    if moved:
                         self.log(f"Moved control file {Path(file_path).name} to {controls_folder}")
                    else:
                         self.log(f"Failed to move control file {Path(file_path).name}")

          # Process blanks
          if blank_files:
               self.log(f"Processing {len(blank_files)} blank files")
               blank_folder = Path(new_folder_path) / "Blank"
               blank_folder.mkdir(exist_ok=True)

               for file_path in blank_files:
                    target_path = blank_folder / Path(file_path).name
                    moved = self.file_dao.move_file(file_path, str(target_path))
                    if moved:
                         self.log(f"Moved blank file {Path(file_path).name} to {blank_folder}")
                    else:
                         self.log(f"Failed to move blank file {Path(file_path).name}")

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
                    file_name = Path(file_path).name

                    # Check if this is a preemptive but has no match in order key
                    is_preempt = self.is_preemptive(file_name)

                    if is_preempt:
                         # For preemptive files with no match, put in Unsorted folder
                         if unsorted_folder is None:
                              unsorted_folder = Path(new_folder_path) / "Unsorted"
                              unsorted_folder.mkdir(exist_ok=True)

                         # Keep original name for Unsorted folder
                         target_path = unsorted_folder / file_name
                         moved = self.file_dao.move_file(file_path, str(target_path))
                         if moved:
                              self.log(f"Moved preemptive unmatched file {file_name} to Unsorted folder")
                         else:
                              self.log(f"Failed to move unmatched file {file_name}")
                    else:
                         # Clean filename for destination (remove braces)
                         clean_brace_file_name = re.sub(r'{.*?}', '', file_name)
                         target_path = Path(new_folder_path) / clean_brace_file_name

                         moved = self.file_dao.move_file(file_path, str(target_path))
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

     def sort_plate_folder(self, folder_path):
          """Sort all files in a plate folder - much simpler than ind folder processing"""
          self.log(f"Processing plate folder: {folder_path}")

          # Get all .ab1 files in the folder (no recursive search needed for plates)
          ab1_files = []
          for file in Path(folder_path).iterdir():
               if file.suffix == ".ab1":
                    ab1_files.append(str(file))

          self.log(f"Found {len(ab1_files)} .ab1 files in folder")

          # Group files by type for batch processing
          control_files = []
          blank_files = []
          brace_files = []

          # Classify files
          for file_path in ab1_files:
               file_name = Path(file_path).name

               # Check for control files (use PLATE_CONTROLS)
               if self.file_dao.is_control_file(file_name, self.config.PLATE_CONTROLS):
                    self.log(f"Identified control file: {file_name}")
                    control_files.append(file_path)
                    continue

               # Check for blank files
               if self.file_dao.is_blank_file(file_name):
                    self.log(f"Identified blank file: {file_name}")
                    blank_files.append(file_path)
                    continue

               # Check for files with braces (need cleaning)
               if '{' in file_name or '}' in file_name:
                    brace_files.append(file_path)

          # Detailed logging
          self.log(f"Classified {len(control_files)} controls, {len(blank_files)} blanks, "
                    f"{len(brace_files)} files with braces")

          # Process controls
          if control_files:
               self.log(f"Processing {len(control_files)} control files")
               controls_folder = Path(folder_path) / "Controls"
               controls_folder.mkdir(exist_ok=True)

               for file_path in control_files:
                    target_path = controls_folder / Path(file_path).name
                    moved = self.file_dao.move_file(file_path, str(target_path))
                    if moved:
                         self.log(f"Moved control file {Path(file_path).name} to Controls")
                    else:
                         self.log(f"Failed to move control file {Path(file_path).name}")

          # Process blanks
          if blank_files:
               self.log(f"Processing {len(blank_files)} blank files")
               blank_folder = Path(folder_path) / "Blank"
               blank_folder.mkdir(exist_ok=True)

               for file_path in blank_files:
                    target_path = blank_folder / Path(file_path).name
                    moved = self.file_dao.move_file(file_path, str(target_path))
                    if moved:
                         self.log(f"Moved blank file {Path(file_path).name} to Blank")
                    else:
                         self.log(f"Failed to move blank file {Path(file_path).name}")

          # Remove braces from remaining files
          if brace_files:
               self.log(f"Removing braces from {len(brace_files)} files")
               braces_count = 0
               for file_path in brace_files:
                    # Skip if file was already moved to Controls or Blank folders
                    if not Path(file_path).exists():
                         continue

                    new_path = self.file_dao.rename_file_without_braces(file_path)
                    if new_path != file_path:
                         braces_count += 1
                         self.log(f"Removed braces from {Path(file_path).name}")

               self.log(f"Successfully removed braces from {braces_count} files")

          self.log(f"Completed processing plate folder: {Path(folder_path).name}")
          return folder_path

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
          if isinstance(order_identifier, str) and Path(order_identifier).exists():
               # Extract I-number from folder name
               folder_name = Path(order_identifier).name
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

          for item in Path(bio_folder).iterdir():
               if item.is_dir():
                    if re.search(r'bioi-\d+_.+_\d+', item.name.lower()) and not re.search('reinject', item.name.lower()):
                         # If item looks like an order folder BioI-20000_YoMama_123456 and not a reinject folder
                         order_folders.append(str(item))

          return order_folders

     def _get_well_locations(self, filename):
          """Extract well locations from a filename"""
          well_locations = []
          matches = re.finditer(r'{(\d+[A-H])}', filename)
          for match in matches:
               well_locations.append(match.group(1))
          return well_locations

     def _place_customer_file(self, file_path, destination_folder, normalized_name):
          """Place customer file in main folder or Alternate Injections based on file characteristics"""

          if self._should_use_alternate_injections(file_path, destination_folder, normalized_name):
               return self._move_to_alternate_injections(file_path, destination_folder)
          else:
               return self._move_to_main_folder(file_path, destination_folder)

     def _should_use_alternate_injections(self, file_path, destination_folder, normalized_name):
          """Check if customer file should be placed in Alternate Injections folder"""
          file_name = Path(file_path).name
          # Use existing DAO function to get base name without order number
          base_name = self.file_dao.standardize_filename_for_matching(file_name, preserve_order_number=False)

          placement_reasons = []

          if self.is_in_not_needed_folder(file_path):
               placement_reasons.append("from NN folder")

          if self._is_customer_file_in_reinject_list(file_name, base_name):
               placement_reasons.append("in reinject list")

          if self._has_preemptive_conflicts(file_name, destination_folder, base_name):
               placement_reasons.append("preemptive with conflicts")

          if self._would_overwrite_existing_file(file_path, destination_folder):
               placement_reasons.append("would overwrite existing")

          if placement_reasons:
               reason_text = ', '.join(placement_reasons)
               self.log(f"Customer file goes to Alternate Injections ({reason_text}): {file_name}")
               return True

          return False

     def _is_customer_file_in_reinject_list(self, file_name, base_normalized_name):
          """Check if customer file is in reinject list using well location matching"""
          if not (hasattr(self, 'reinject_list') and self.reinject_list and hasattr(self, 'raw_reinject_list')):
               return False

          current_wells = self._get_well_locations(file_name)
          if not current_wells:
               return False

          current_well = current_wells[0]

          # Customer files use complex well-based matching (different from PCR files)
          for raw_entry in self.raw_reinject_list:
               reinject_wells = self._get_well_locations(raw_entry)
               if len(reinject_wells) >= 2 and reinject_wells[1] == current_well:
                    # Use customer-specific normalization consistently
                    reinj_base = self.file_dao.standardize_for_customer_files(raw_entry, remove_extension=True)

                    if base_normalized_name == reinj_base:
                         return True

          return False

     def _has_preemptive_conflicts(self, file_name, destination_folder, base_normalized_name):
          """Check if preemptive file would conflict with existing files"""
          if not self.is_preemptive(file_name):
               return False

          matching_files = self.find_matching_files(destination_folder, base_normalized_name)
          return len(matching_files) > 0

     def _would_overwrite_existing_file(self, file_path, destination_folder):
          """Check if cleaned filename would overwrite existing file"""
          file_name = Path(file_path).name
          # Use existing DAO function instead of manual regex
          clean_name = self.file_dao.clean_braces_format(file_name)
          target_path = Path(destination_folder) / clean_name
          return target_path.exists()

     def _move_to_alternate_injections(self, file_path, destination_folder):
          """Move file to Alternate Injections folder preserving original name"""
          file_name = Path(file_path).name
          alt_folder = Path(destination_folder) / "Alternate Injections"
          alt_folder.mkdir(exist_ok=True)

          alt_path = alt_folder / file_name  # Keep original name with braces
          success = self.file_dao.move_file(file_path, str(alt_path))
          if success:
               self.log(f"Moved to Alternate Injections: {file_name}")
          return success

     def _move_to_main_folder(self, file_path, destination_folder):
          """Move file to main folder with cleaned filename (braces removed only)"""
          file_name = Path(file_path).name
          # Remove only braces, keep suffixes like _Premixed and _RTI
          clean_name = re.sub(r'{.*?}', '', file_name)

          target_path = Path(destination_folder) / clean_name
          success = self.file_dao.move_file(file_path, str(target_path))
          if success:
               self.log(f"Moved to main folder: {file_name}")
          return success

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
          """Process an individual I number plate folder - BioI folder containing multiple order folders"""
          folder_name = Path(bio_folder).name
          self.log(f"Processing BioI folder: {folder_name}")

          # Get all order folders in this BioI folder
          order_folders = self.get_order_folders(bio_folder)

          # Process each order folder individually
          for order_folder in order_folders:
               self.process_order_folder(order_folder, bio_folder)

          return

     def process_order_folder(self, order_folder, parent_folder=None):
          """Process an individual order folder"""
          folder_name = Path(order_folder).name
          self.log(f"Processing order folder: {folder_name}")

          # Check for special case: re-sequenced requests (folders ending with underscore and single digit)
          # These represent secondary uploads, additional requests, etc.
          import re
          is_resequenced = bool(re.search(r'_\d$', folder_name))

          # Determine if we're in IND Not Ready context
          in_not_ready = Path(order_folder).parent.name == self.config.IND_NOT_READY_FOLDER

          # Get order information
          order_number = self.get_order_number_from_folder_name(order_folder)
          ab1_files = self.file_dao.get_files_by_extension(order_folder, self.config.ABI_EXTENSION)
          is_andreev_order = self.config.ANDREEV_NAME in folder_name.lower()

          # Check order completeness
          if is_resequenced:
               # For re-sequenced requests, bypass normal file count checking
               # These are often partial uploads with only a few reactions
               files_complete = len(ab1_files) > 0  # Just need to have some files
               self.log(f"Re-sequenced request detected: {folder_name}, bypassing file count validation")
          else:
               # Normal file count validation
               expected_count = self._get_expected_file_count(order_number)
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
                              time.sleep(0.3)  # Wait for handles to be released

                    # If in IND Not Ready, move back to appropriate location
                    if in_not_ready:
                         # Get parent BioI folder path
                         bioi_folder = self.get_destination_for_order(order_folder, parent_folder)

                         if bioi_folder:
                              # Construct the destination path
                              destination = Path(bioi_folder) / folder_name

                              # Check if destination exists
                              if destination.exists():
                                   self.log(f"Destination already exists: {destination}, leaving in current location")
                              else:
                                   # Use direct shutil.move to avoid nested directory issues
                                   try:
                                        import shutil
                                        shutil.move(order_folder, str(destination))
                                        self.log(f"Order moved back: {folder_name}")
                                   except Exception as e:
                                        self.warning(f"Error moving folder: {e}")
                                        self.log(f"Failed to move folder, leaving in current location: {folder_name}")
                         else:
                              self.log(f"Could not determine destination BioI folder, leaving in current location: {folder_name}")
               else:
                    # Not complete - move to IND Not Ready
                    if not in_not_ready:
                         not_ready_path = Path(parent_folder or order_folder).parent / self.config.IND_NOT_READY_FOLDER
                         self.file_dao.create_folder_if_not_exists(str(not_ready_path))

                         target_path = not_ready_path / folder_name
                         if target_path.exists():
                              self.log(f"Destination already exists in IND Not Ready: {target_path}")
                         else:
                              try:
                                   import shutil
                                   shutil.move(order_folder, str(target_path))
                                   self.log(f"Incomplete order moved to Not Ready: {folder_name}")
                              except Exception as e:
                                   self.warning(f"Error moving to IND Not Ready: {e}")
                                   self.log(f"Failed to move folder to IND Not Ready: {folder_name}")

          # Already mSeqed but in IND Not Ready - move back
          elif was_mseqed and in_not_ready:
               # Close mSeq to release file handles before moving
               if self.ui_automation:
                    self.ui_automation.close()
                    time.sleep(0.3)  # Wait for handles to be released

               # Get parent BioI folder path
               bioi_folder = self.get_destination_for_order(order_folder, parent_folder)

               if bioi_folder:
                    # Construct the destination path
                    destination = Path(bioi_folder) / folder_name

                    # Check if destination exists
                    if destination.exists():
                         self.log(f"Destination already exists: {destination}, leaving in current location")
                    else:
                         # Use direct shutil.move to avoid nested directory issues
                         try:
                              import shutil
                              shutil.move(order_folder, str(destination))
                              self.log(f"Processed order moved back: {folder_name}")
                         except Exception as e:
                              self.warning(f"Error moving folder: {e}")
                              self.log(f"Failed to move folder, leaving in current location: {folder_name}")
               else:
                    self.log(f"Could not determine destination BioI folder, leaving in current location: {folder_name}")

     def get_pcr_folder_path(self, pcr_number, base_path):
          """
          Get proper PCR folder path or create one if needed.

          Args:
               pcr_number: PCR number (without the 'PCR' prefix)
               base_path: The base folder path to search in

          Returns:
               str: Path to the PCR folder
          """
          self.log(f"Looking for PCR folder with PCR{pcr_number} in {base_path}")

          # Search for any folder containing the PCR number
          pcr_pattern = re.compile(f'pcr{pcr_number}', re.IGNORECASE)
          found_folders = []

          for item in Path(base_path).iterdir():
               if item.is_dir() and pcr_pattern.search(item.name.lower()):
                    found_folders.append(str(item))

          if found_folders:
               # If multiple matches found, prioritize ones with order numbers
               order_pattern = re.compile(f'pcr{pcr_number}_\\d+', re.IGNORECASE)
               for folder in found_folders:
                    if order_pattern.search(Path(folder).name.lower()):
                         self.log(f"Found PCR folder with order number: {folder}")
                         return folder

               # If no folder with order number, use the first match
               self.log(f"Found PCR folder: {found_folders[0]}")
               return found_folders[0]

          # No matching folder found, create a new one
          pcr_folder_name = f"FB-PCR{pcr_number}"
          new_folder_path = Path(base_path) / pcr_folder_name
          self.log(f"Creating new PCR folder: {new_folder_path}")
          self.file_dao.create_folder_if_not_exists(str(new_folder_path))
          return str(new_folder_path)

     def process_pcr_folder(self, pcr_folder):
          """Process all files in a PCR folder - updated matching logic"""
          self.log(f"Processing PCR folder: {Path(pcr_folder).name}")

          for file in Path(pcr_folder).iterdir():
               if file.suffix == self.config.ABI_EXTENSION:
                    file_path = str(file)
                    file_name = Path(file_path).name

                    # Check if file is in NN folder
                    is_from_nn = self.is_in_not_needed_folder(pcr_folder)

                    # Clean the filename for matching with reinject list
                    clean_name = self.file_dao.standardize_filename_for_matching(file_name)

                    # Check if it's in the reinject list
                    is_reinject = clean_name in self.reinject_list

                    # Check if it's a preemptive reinject
                    is_preempt = self.is_preemptive(file_name)

                    # Determine if file should go to Alternate Injections
                    go_to_alt_injections = is_from_nn or is_reinject

                    if is_preempt and not is_reinject:
                         # Check for other files with matching name
                         matching_files = self.find_matching_files(pcr_folder, clean_name)
                         # Remove the current file from consideration
                         matching_files = [f for f in matching_files if f != file_name]
                         if matching_files:
                              go_to_alt_injections = True

                    if go_to_alt_injections:
                         # Make sure Alternate Injections folder exists
                         alt_folder = Path(pcr_folder) / "Alternate Injections"
                         alt_folder.mkdir(exist_ok=True)

                         # Move file to Alternate Injections
                         source = file_path
                         dest = alt_folder / file_name  # Keep original filename with brackets
                         if Path(source).exists() and not dest.exists():
                              self.file_dao.move_file(source, str(dest))
                              reason = ""
                              if is_from_nn: reason += "from NN folder "
                              if is_reinject: reason += "in reinject list "
                              if is_preempt and not is_reinject and matching_files: reason += "preemptive with match"
                              self.log(f"Moved file {file_name} to Alternate Injections ({reason.strip()})")

          # Now process the folder with mSeq if it's ready
          was_mseqed, has_braces, has_ab1_files = self.check_order_status(pcr_folder)
          if not was_mseqed and not has_braces and has_ab1_files:
               self.ui_automation.process_folder(pcr_folder)
               self.log(f"mSeq completed: {Path(pcr_folder).name}")
          else:
               self.log(f"mSeq NOT completed: {Path(pcr_folder).name}")

     def _cleanup_original_folder(self, original_folder: str, new_folder: str):
          """
          Enhanced cleanup to remove the original folder if all files have been processed,
          handle nested NN folders, and clean up empty folders
          """
          if original_folder == new_folder:
               self.log("Original folder is the same as new folder, no cleanup needed")
               return

          # Process nested NN folders first
          # Walk the directory tree from bottom-up to process subdirectories first
          for dir_path in reversed(list(Path(original_folder).rglob("*"))):
               if dir_path.is_dir():
                    dir_lower = dir_path.name.lower()

                    # Check if this is an NN folder
                    if dir_lower == "nn" or dir_lower == "nn-preemptives" or dir_lower.startswith("nn_") or dir_lower.startswith("nn-preemptives_"):
                         self.log(f"Processing nested NN folder: {dir_path.name}")

                         # Get all AB1 files in this NN folder
                         nn_ab1_files = self.file_dao.get_files_by_extension(str(dir_path), ".ab1", recursive=True)

                         if nn_ab1_files:
                              self.log(f"Found {len(nn_ab1_files)} AB1 files in nested NN folder")

                              # Process each file
                              for file_path in nn_ab1_files:
                                   file_name = Path(file_path).name
                                   normalized_name = self.file_dao.normalize_filename(file_name)

                                   # Find matching entry in order key (if available)
                                   if self.order_key_index and normalized_name in self.order_key_index:
                                        matches = self.order_key_index[normalized_name]
                                        i_num, acct_name, order_num = matches[0]  # Use first match

                                        # Create destination folder and move file
                                        destination_folder = self.create_order_folder(i_num, acct_name, order_num)
                                        self._place_customer_file(file_path, destination_folder, normalized_name)
                                   else:
                                        self.log(f"No order key match for nested NN file: {file_name}")

                         # After processing files, check if NN folder is now empty
                         if not any(dir_path.iterdir()):
                              try:
                                   dir_path.rmdir()
                                   self.log(f"Deleted empty NN folder: {dir_path.name}")
                              except Exception as e:
                                   self.log(f"Failed to delete empty NN folder {dir_path.name}: {e}")

          # Force refresh directory cache after processing NN folders
          self.file_dao.get_directory_contents(original_folder, refresh=True)

          # Check if any .ab1 files remain in the original folder
          ab1_files = self.file_dao.get_files_by_extension(original_folder, ".ab1")
          if ab1_files:
               self.log(f"Found {len(ab1_files)} remaining .ab1 files in original folder, cannot delete")

               # Log the names of remaining files for debugging
               for file_path in ab1_files:
                    self.log(f"Remaining file: {Path(file_path).name}")

               return

          # Refresh directory contents
          remaining_items = self.file_dao.get_directory_contents(original_folder, refresh=True)

          # If completely empty, delete the folder
          if not remaining_items:
               try:
                    Path(original_folder).rmdir()
                    self.log(f"Deleted empty original folder: {original_folder}")
                    return
               except Exception as e:
                    self.log(f"Failed to delete original folder: {e}")
                    return

          # Try to move any remaining Control/Blank folders to the new location
          moved_all = True
          item: str
          for item in list(remaining_items):  # Create a copy of the list to avoid iteration issues
               item_path = Path(original_folder) / str(item)

               if item in ["Controls", "Blank", "Alternate Injections"]:
                    # Try to move this folder to the new location if it's not empty
                    if item_path.is_dir() and any(item_path.iterdir()):
                         try:
                              # Create target folder in new location if needed
                              target_path = Path(new_folder) / item
                              target_path.mkdir(exist_ok=True)

                              # Move all files from old to new location
                              for subitem in item_path.iterdir():
                                   if subitem.is_file():
                                        new_file = target_path / subitem.name
                                        self.file_dao.move_file(str(subitem), str(new_file))
                                        self.log(f"Moved remaining file {subitem.name} to {target_path}")

                              # Check if folder is now empty and can be deleted
                              if not any(item_path.iterdir()):
                                   item_path.rmdir()
                                   self.log(f"Deleted now-empty folder: {item}")
                         except Exception as e:
                              self.log(f"Failed to move remaining files from {item}: {e}")
                              moved_all = False
                              continue

                    elif item_path.is_dir() and not any(item_path.iterdir()):
                         # Empty folder - delete it
                         try:
                              item_path.rmdir()
                              self.log(f"Deleted empty folder: {item}")
                         except Exception as e:
                              self.log(f"Failed to delete empty folder {item}: {e}")
                              moved_all = False
               elif item_path.is_dir():
                    # Check if this is a non-standard folder (like NN) and if it's empty
                    dir_lower = item.lower()
                    if (dir_lower == "nn" or dir_lower == "nn-preemptives" or
                         dir_lower.startswith("nn_") or dir_lower.startswith("nn-preemptives_")) and not any(item_path.iterdir()):
                         try:
                              item_path.rmdir()
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
                    Path(original_folder).rmdir()
                    self.log(f"Deleted original folder after cleaning: {original_folder}")
               except Exception as e:
                    self.log(f"Failed to delete original folder: {e}")
          else:
               # Check if all remaining items are empty NN folders
               all_empty_nn = True
               for item in remaining_items:
                    item_path = Path(original_folder) / str(item)
                    if not item_path.is_dir():
                         all_empty_nn = False
                         break

                    dir_lower = item.lower()
                    if not (dir_lower == "nn" or dir_lower == "nn-preemptives" or
                         dir_lower.startswith("nn_") or dir_lower.startswith("nn-preemptives_")):
                         all_empty_nn = False
                         break

                    if any(item_path.iterdir()):
                         all_empty_nn = False
                         break

               if all_empty_nn:
                    # Delete all empty NN folders
                    for item in list(remaining_items):
                         item_path = Path(original_folder) / str(item)
                         try:
                              item_path.rmdir()
                              self.log(f"Deleted empty NN folder: {item}")
                         except Exception as e:
                              self.log(f"Failed to delete empty NN folder {item}: {e}")

                    # Try to delete the original folder again
                    try:
                         Path(original_folder).rmdir()
                         self.log(f"Deleted original folder after cleaning all empty NN folders: {original_folder}")
                    except Exception as e:
                         self.log(f"Failed to delete original folder: {e}")
               else:
                    self.log(f"Unable to clean up original folder. {len(remaining_items)} items remain.")

     def _sort_pcr_file(self, file_path, pcr_number):
          """Sort a PCR file to the appropriate folder"""
          file_name = Path(file_path).name
          self.log(f"Processing PCR file: {file_name} with PCR Number: {pcr_number}")

          # Check if file is in a Not Needed folder
          is_from_nn = self.is_in_not_needed_folder(file_path)

          # IMPORTANT: Always use the user-selected folder for consistency
          # This is the key change - we always use the same base path
          # If this is undefined, use the current_data_folder attribute which should be set
          if hasattr(self, 'current_data_folder') and self.current_data_folder: #type: ignore
               day_data_path = self.current_data_folder #type: ignore
          else:
               # Fallback to parent folder logic - should not normally be used
               current_folder = str(Path(file_path).parent)
               # For NN folder files, go up two levels instead of one
               if is_from_nn:
                    day_data_path = str(Path(current_folder).parent.parent)
               else:
                    day_data_path = str(Path(current_folder).parent)

          self.log(f"Using day data path: {day_data_path}")

          # Find or create PCR folder
          pcr_folder_path = self.get_pcr_folder_path(pcr_number, day_data_path)
          self.log(f"Target PCR folder: {pcr_folder_path}")

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
          target_path = Path(pcr_folder_path) / clean_name

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
          if target_path.exists():
               go_to_alt_injections = True
               self.log(f"PCR file already exists at destination: {file_name}")

          # Move the file to the appropriate location
          if go_to_alt_injections:
               alt_inj_folder = Path(pcr_folder_path) / "Alternate Injections"
               alt_inj_folder.mkdir(exist_ok=True)
               alt_file_path = alt_inj_folder / file_name  # Keep original name with braces
               return self.file_dao.move_file(file_path, str(alt_file_path))
          else:
               # Put in main folder
               return self.file_dao.move_file(file_path, str(target_path))

     def _sort_control_file(self, file_path):
          """Sort a control file to the Controls folder"""
          parent_folder = Path(file_path).parent
          controls_folder = parent_folder / "Controls"
          controls_folder.mkdir(exist_ok=True)

          # Just move the file directly (no special handling needed)
          target_path = controls_folder / Path(file_path).name
          return self.file_dao.move_file(file_path, str(target_path))

     def _sort_blank_file(self, file_path):
          """Sort a blank file to the Blank folder"""
          # Get parent BioI folder first, not just the immediate parent folder
          immediate_parent = Path(file_path).parent
          file_name = Path(file_path).name

          # Get or create the BioI folder
          i_num = self.file_dao.get_inumber_from_name(str(immediate_parent))
          if i_num:
               bioi_folder_name = f"BioI-{i_num}"
               bioi_folder_path = immediate_parent.parent / bioi_folder_name
               bioi_folder_path.mkdir(exist_ok=True)

               # Create Blank folder inside BioI folder
               blank_folder = bioi_folder_path / "Blank"
               blank_folder.mkdir(exist_ok=True)

               # Move file to Blank folder
               target_path = blank_folder / file_name
               return self.file_dao.move_file(file_path, str(target_path))
          else:
               # Fallback to original logic if no I number found
               parent_folder = immediate_parent
               blank_folder = parent_folder / "Blank"
               blank_folder.mkdir(exist_ok=True)
               target_path = blank_folder / file_name
               return self.file_dao.move_file(file_path, str(target_path))

     def _rename_processed_folder(self, folder_path):
          """Rename the folder after processing"""
          i_num = self.file_dao.get_inumber_from_name(folder_path)
          if i_num:
               new_folder_name = f"BioI-{i_num}"
               folder_pathobj = Path(folder_path)
               new_folder_path = folder_pathobj.parent / new_folder_name

               # Only rename if the new name doesn't already exist
               if not new_folder_path.exists() and folder_pathobj.name != new_folder_name:
                    try:
                         folder_pathobj.rename(new_folder_path)
                         self.log(f"Renamed folder to: {new_folder_name}")
                         return True
                    except Exception as e:
                         self.log(f"Failed to rename folder: {e}")

          return False

     def get_order_number_from_folder_name(self, folder_path):
          """Extract order number from folder name"""
          folder_name = Path(folder_path).name
          match = re.search(r'_\d+$', folder_name)
          if match:
               order_number = re.search(r'\d+', match.group(0)).group(0) #type: ignore
               return order_number
          return None

     def is_in_not_needed_folder(self, file_path):
          """
          Check if file is in a Not Needed folder at any level of nesting
          """
          # Ensure file_path is a string to avoid Path.replace() vs str.replace() confusion
          file_path_str = str(file_path)

          # Get the path components
          path_parts = file_path_str.replace('\\', '/').split('/')

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
          # Remove order number suffix if present
          base_normalized_name = normalized_name
          if '#' in normalized_name:
               base_normalized_name = normalized_name.split('#', 1)[0]

          matching_files = []
          for item in Path(folder_path).iterdir():
               if item.suffix == self.config.ABI_EXTENSION:
                    item_norm = self.file_dao.normalize_filename(item.name)
                    # Also handle potential order numbers in the normalized matched files
                    if '#' in item_norm:
                         item_norm = item_norm.split('#', 1)[0]

                    if item_norm == base_normalized_name:
                         matching_files.append(item.name)
          return matching_files

     def get_reinject_list(self, i_numbers, reinject_path=None):
          """Get list of reactions that are reinjects - optimized version"""
          from pathlib import Path

          # Helper function to check if entry is just a well location
          def is_valid_entry(raw_name):
               if not raw_name or not isinstance(raw_name, str):
                    return False
               # Skip entries that are just well locations in braces
               if re.match(r'^\{\d+[A-H]\}$', raw_name):
                    return False
               # Skip empty entries
               if not raw_name.strip():
                    return False
               return True

          reinject_list = []
          raw_reinject_list = []
          processed_files = set()

          # Process text files only once
          spreadsheets_path = Path(r'G:\Lab\Spreadsheets')
          abi_path = spreadsheets_path / 'Individual Uploaded to ABI'
          reinject_files = []

          # Build list of reinject files
          if spreadsheets_path.exists():
               for file_path in self.file_dao.get_directory_contents(spreadsheets_path):
                    # file_path is a Path object
                    if 'reinject' in file_path.name.lower() and file_path.suffix == '.txt':
                         if any(i_num in file_path.name for i_num in i_numbers):
                              reinject_files.append(str(file_path))  # Convert to string for numpy

          if abi_path.exists():
               for file_path in self.file_dao.get_directory_contents(abi_path):
                    # file_path is a Path object
                    if 'reinject' in file_path.name.lower() and file_path.suffix == '.txt':
                         if any(i_num in file_path.name for i_num in i_numbers):
                              reinject_files.append(str(file_path))  # Convert to string for numpy

          # Process each found reinject file
          for file_path in reinject_files:
               try:
                    if file_path in processed_files:
                         continue

                    processed_files.add(file_path)
                    self.log(f"Processing reinject file: {file_path}")
                    data = np.loadtxt(file_path, dtype=str, delimiter='\t', skiprows=5)

                    # Parse rows (data already has headers skipped via skiprows=5)
                    for j in range(0, min(96, data.shape[0])):  # Adjusted range since we skipped headers
                         if j < data.shape[0] and data.shape[1] > 1:
                              raw_name = data[j, 1]
                              if is_valid_entry(raw_name):
                                   cleaned_name = self.file_dao.standardize_filename_for_matching(raw_name)
                                   raw_reinject_list.append(raw_name)
                                   reinject_list.append(cleaned_name)
               except Exception as e:
                    self.log(f"Error processing reinject file {file_path}: {e}")

          # Excel file processing
          if reinject_path and Path(reinject_path).exists():
               try:
                    import pylightxl as xl
                    db = xl.readxl(reinject_path)
                    sheet = db.ws('Sheet1')

                    # Note: Excel file exists but we process all txt files for the I-numbers
                    # rather than filtering by Excel contents

                    # Check for partial plate reinjects by comparing against reinject_prep_array
                    for i_num in i_numbers:
                         # Check all txt files for this I number
                         regex_pattern = f'.*{i_num}.*txt'
                         txt_files = []

                         for file_path in self.file_dao.get_directory_contents(spreadsheets_path):
                              # file_path is now a Path object
                              if re.search(regex_pattern, file_path.name, re.IGNORECASE) and 'reinject' not in file_path.name:
                                   txt_files.append(str(file_path))  # Convert to string for numpy

                         if abi_path.exists():
                              for file_path in self.file_dao.get_directory_contents(abi_path):
                                   # file_path is now a Path object
                                   if re.search(regex_pattern, file_path.name, re.IGNORECASE) and 'reinject' not in file_path.name:
                                        txt_files.append(str(file_path))  # Convert to string for numpy

                         # Process each txt file
                         for file_path in txt_files:
                              try:
                                   data = np.loadtxt(file_path, dtype=str, delimiter='\t', skiprows=5)
                                   for j in range(0, min(96, data.shape[0])):  # Adjusted range here too
                                        if j < data.shape[0] and data.shape[1] > 1:
                                             raw_name = data[j, 1]
                                             if is_valid_entry(raw_name):
                                                  raw_reinject_list.append(raw_name)
                                                  cleaned_name = self.file_dao.standardize_filename_for_matching(raw_name)
                                                  reinject_list.append(cleaned_name)
                              except Exception as e:
                                   self.log(f"Error processing txt file {file_path}: {e}")

               except Exception as e:
                    self.log(f"Error processing reinject Excel file: {e}")

          # Store the raw_reinject_list for reference
          self.raw_reinject_list = raw_reinject_list
          return reinject_list

     def debug_reinject_detection(self, file_path):
          """Debug helper for all file sorting logic"""
          file_name = Path(file_path).name

          # Check all conditions
          is_from_nn = self.is_in_not_needed_folder(file_path)

          # Check if in reinject list - FIXED: Don't double-normalize
          normalized_name = self.file_dao.standardize_filename_for_matching(file_name)
          is_reinject = False
          if hasattr(self, 'reinject_list') and self.reinject_list:
               # reinject_list already contains normalized entries, so compare directly
               is_reinject = normalized_name in self.reinject_list

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

          folder_path = Path(folder_path)
          folder_contents = list(self.file_dao.get_directory_contents(folder_path))

          # Check for mSeq directory structure
          mseq_set = {'chromat_dir', 'edit_dir', 'phd_dir', 'mseq4.ini'}
          current_proj = [item.name if hasattr(item, "name") else str(item) for item in folder_contents if (item.name if hasattr(item, "name") else str(item)) in mseq_set]

          if set(current_proj) == mseq_set:
               was_mseqed = True

          # Check for the 5 txt files as an additional verification
          txt_file_count = 0
          for item in folder_contents:
               item_name = item.name if hasattr(item, "name") else str(item)
               item_path = folder_path / item_name
               if item_path.is_file():
                    if (item_name.endswith('.raw.qual.txt') or
                              item_name.endswith('.raw.seq.txt') or
                              item_name.endswith('.seq.info.txt') or
                              item_name.endswith('.seq.qual.txt') or
                              item_name.endswith('.seq.txt')):
                         txt_file_count += 1

          if txt_file_count == 5:
               was_mseqed = True

          # Check for .ab1 files and braces
          for item in folder_contents:
               item_name = item.name if hasattr(item, "name") else str(item)
               if item_name.endswith(self.config.ABI_EXTENSION):
                    has_ab1_files = True
                    if '{' in item_name or '}' in item_name:
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
               folder_path_obj = Path(folder_path)
               folder_name = folder_path_obj.name
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

               zip_path = folder_path_obj / zip_filename

               # Check if this folder contains FSA files
               has_fsa_files = self.file_dao.contains_file_type(folder_path, self.config.FSA_EXTENSION)

               # Determine which files to include based on file types present
               if has_fsa_files:
                    # FSA folders: include only .fsa files
                    file_extensions = [self.config.FSA_EXTENSION]
                    self.log(f"FSA folder detected, including only .fsa files: {zip_filename}")
               else:
                    # Regular folders: include .ab1 files and optionally .txt files
                    file_extensions = [self.config.ABI_EXTENSION]
                    if include_txt:
                         file_extensions.extend(self.config.TEXT_FILES)
                    self.log(f"Regular folder, including .ab1 files{' and .txt files' if include_txt else ''}: {zip_filename}")

               # Create zip file
               self.log(f"Creating zip file: {zip_filename}")
               success = self.file_dao.zip_files(
                    source_folder=str(folder_path_obj),  # Convert Path object to string
                    zip_path=str(zip_path), # Convert Path object back to string
                    file_extensions=file_extensions,
                    exclude_extensions=None
               )

               if success:
                    self.log(f"Successfully created zip file: {zip_path}")
                    return str(zip_path) # Convert Path object back to string
               else:
                    self.log(f"Failed to create zip file: {zip_path}")
                    return None
          except Exception as e:
               self.log(f"Error creating zip file for {folder_path}: {e}")
               return None

     def zip_full_plasmid_order_folder(self, folder_path, order_key=None, order_number=None, use_7zip=False, compression_level=6):
          """
          Zip the contents of a full plasmid sequencing order folder.

          Args:
               folder_path (str): Path to the order folder
               order_key (array, optional): Order key data to extract sample names
               order_number (str, optional): Order number to filter order key entries
               use_7zip (bool, optional): Whether to try using 7-Zip for compression (default: False)
               compression_level (int, optional): Compression level 0-9, where 9 is maximum (default: 6)

          Returns:
               str: Path to created zip file, or None if failed
          """
          try:
               folder_name = Path(folder_path).name
               self.log(f"Creating zip for full plasmid order: {folder_name}")
               self.log(f"Compression settings: Method={'7-Zip' if use_7zip else 'Python zipfile'}, Level={compression_level}")

               # Set up the zip filename
               zip_filename = f"{folder_name}.zip"
               zip_path = Path(folder_path) / zip_filename

               # Define file extensions to include
               plasmid_extensions = [
                    '.annotations.bed',
                    '.annotations.gbk',
                    '.assembly_stats.tsv',
                    '.final.fasta',
                    '.final.fastq',
                    '.html'  # Include the validation report HTML
               ]

               # Collect all files to zip
               all_files = []
               folder_pathobj = Path(folder_path)
               for item in folder_pathobj.iterdir():
                    if item.is_file() and any(item.name.endswith(ext) for ext in plasmid_extensions):
                         all_files.append(item.name)

               if not all_files:
                    self.log(f"No matching files found in {folder_path}")
                    return None

               self.log(f"Found {len(all_files)} matching files to include in zip")

               # Try to use 7-Zip if requested and available
               if use_7zip:
                    seven_zip_exe = r"G:\Lab\DATA Processing\7-Zip\7z.exe"

                    self.log(f"Checking for 7-Zip at: {seven_zip_exe}")
                    if Path(seven_zip_exe).exists():
                         self.log(f"Found 7-Zip at: {seven_zip_exe}")
                         self.log("USING 7-ZIP FOR COMPRESSION")

                         # Create a list file for 7-Zip
                         list_file_path = folder_pathobj / "files_to_zip.txt"

                         with open(list_file_path, "w") as f:
                              for item in all_files:
                                   f.write(f"{item}\n")

                         # Build 7-Zip command with the specified compression level
                         command = [
                              seven_zip_exe,
                              "a",                      # Add files to archive
                              "-tzip",                  # ZIP format
                              f"-mx={compression_level}", # Use the specified compression level
                              str(zip_path),            # Output ZIP file
                              "@" + str(list_file_path) # Input list of files
                         ]

                         self.log(f"7-Zip command: {' '.join(command)}")

                         # Run 7-Zip
                         import subprocess
                         self.log("Executing 7-Zip subprocess...")

                         try:
                              process = subprocess.run(
                              command,
                              cwd=str(folder_pathobj),  # Set working directory to folder_path
                              capture_output=True,
                              text=True,
                              check=False  # Don't raise exception on non-zero return
                              )

                              self.log(f"7-Zip process return code: {process.returncode}")

                              if process.returncode != 0:
                                   self.log(f"7-Zip error: {process.stderr}")
                                   self.log("Falling back to Python zipfile due to 7-Zip error")
                              else:
                                   # Verify the zip file was created
                                   if zip_path.exists():
                                        self.log(f"Successfully created zip file using 7-Zip: {zip_path}")

                                        # Clean up list file
                                        try:
                                             list_file_path.unlink()
                                        except Exception as e:
                                             self.log(f"Warning: Failed to remove list file: {e}")

                                        return str(zip_path)
                                   else:
                                        self.log(f"7-Zip reported success but zip file not found at: {zip_path}")
                         except Exception as sub_e:
                              self.log(f"Exception during 7-Zip subprocess execution: {sub_e}")
                    else:
                         self.log(f"7-Zip not found at: {seven_zip_exe}")
               else:
                    self.log("7-Zip compression not requested, using Python zipfile")

               # Use Python zipfile (either by choice or as fallback)
               self.log("USING PYTHON ZIPFILE FOR COMPRESSION")
               import zipfile
               from datetime import datetime

               # Get current date/time for all files (Windows compatible format)
               now = datetime.now()
               date_time = (now.year, now.month, now.day, now.hour, now.minute, now.second)

               # Ensure compression level is within valid range
               python_compression = max(0, min(9, compression_level))

               with zipfile.ZipFile(str(zip_path), 'w', zipfile.ZIP_DEFLATED, compresslevel=python_compression) as zipf:
                    for item in all_files:
                         file_path = folder_pathobj / item

                         # Create ZipInfo object for more control
                         zipinfo = zipfile.ZipInfo(item, date_time)

                         # Set Windows-compatible attributes (Archive bit set)
                         zipinfo.external_attr = 0x20  # Archive bit for Windows

                         # Set creation system to Windows
                         zipinfo.create_system = 0  # 0 = Windows

                         # Add the file with custom ZipInfo
                         with open(file_path, 'rb') as f:
                              file_data = f.read()
                              zipf.writestr(zipinfo, file_data, zipfile.ZIP_DEFLATED)

               self.log(f"Successfully created zip file using Python zipfile: {zip_path}")
               return str(zip_path)

          except Exception as e:
               self.log(f"Error creating zip file for {folder_path}: {e}")
               import traceback
               self.log(traceback.format_exc())
               return None

     def find_zip_file(self, folder_path):
          """Find zip file in a folder"""
          folder_pathobj = Path(folder_path)
          for item in folder_pathobj.iterdir():
               if item.suffix == '.zip' and item.is_file():
                    return str(item)
          return None

     def get_zip_mod_time(self, worksheet, order_number):
          """Get the modification time of a zip file from the summary worksheet"""
          for row in range(2, worksheet.max_row + 1):
               if worksheet.cell(row=row, column=2).value == order_number:
                    return worksheet.cell(row=row, column=8).value
          return None

     def get_order_folders_for_validation(self, data_folder):
          """
          Get both BioI order folders and immediate order folders for validation

          Args:
               data_folder (str): Path to data folder

          Returns:
               list: List of tuples (order_folder_path, i_number)
          """
          order_folders = []

          # Get BioI folders
          bio_folders = self.file_dao.get_folders(data_folder, pattern=self.config.REGEX_PATTERNS['bioi_folder'].pattern)

          for bio_folder in bio_folders:
               folder_name = Path(bio_folder).name
               i_number = self.file_dao.get_inumber_from_name(folder_name)

               if not i_number:
                    continue

               # Get order folders within this BioI folder
               bio_folder_path = Path(bio_folder)
               for item in bio_folder_path.iterdir():
                    if (item.is_dir() and
                         self.config.REGEX_PATTERNS['order_folder'].search(item.name.lower()) and
                         not self.config.REGEX_PATTERNS['reinject'].search(item.name.lower())):
                         order_folders.append((str(item), i_number))

          # Get immediate order folders (orders directly in data folder)
          immediate_orders = self.file_dao.get_folders(data_folder, pattern=self.config.REGEX_PATTERNS['order_folder'].pattern)
          # Filter out reinject folders
          immediate_orders = [folder for folder in immediate_orders
                              if not self.config.REGEX_PATTERNS['reinject'].search(Path(folder).name.lower())]

          for order_folder in immediate_orders:
               # Extract I-number from folder name for immediate orders
               i_number = self.file_dao.get_inumber_from_name(Path(order_folder).name)
               if i_number:
                    order_folders.append((order_folder, i_number))

          return order_folders


     def validate_zip_contents(self, zip_path, i_number, order_number, order_key):
          """
          Validate zip file contents against order key

          Args:
               zip_path (str): Path to the zip file
               i_number (str): I number for the order
               order_number (str): Order number
               order_key (array): Order key data

          Returns:
               dict: Validation results with match/mismatch details
          """
          import zipfile
          import numpy as np

          # Results to return
          validation_result = {
               'match_count': 0,
               'mismatch_count': 0,
               'txt_count': 0,
               'expected_count': 0,
               'extra_ab1_count': 0,
               'matches': [],
               'mismatches_in_zip': [],
               'mismatches_in_order': [],
               'txt_files': []
          }

          try:
               # Get expected files from order key for this order
               order_items = []

               # Handle NumPy array properly
               if isinstance(order_key, np.ndarray):
                    # Use NumPy's boolean indexing for efficient filtering
                    mask = (order_key[:, 0] == i_number) & (order_key[:, 2] == order_number)
                    matching_rows = order_key[mask]

                    for row in matching_rows:
                         raw_name = row[3]
                         # Use customer-specific normalization (same as order key)
                         adjusted_name = self.file_dao.standardize_for_customer_files(raw_name, remove_extension=True)
                         order_items.append({'raw_name': raw_name, 'adjusted_name': adjusted_name})
                         self.log(f"DEBUG: Order key entry - Raw: '{raw_name}' -> Adjusted: '{adjusted_name}'")
               else:
                    # Fallback for non-NumPy arrays
                    for entry in order_key:
                         if entry[0] == i_number and entry[2] == order_number:
                              raw_name = entry[3]
                              adjusted_name = self.file_dao.standardize_for_customer_files(raw_name, remove_extension=True)
                              order_items.append({'raw_name': raw_name, 'adjusted_name': adjusted_name})
                              self.log(f"DEBUG: Order key entry - Raw: '{raw_name}' -> Adjusted: '{adjusted_name}'")

               validation_result['expected_count'] = len(order_items)

               # Get actual files from zip
               with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_contents = zip_ref.namelist()

               # Get only AB1 files from zip for matching
               zip_ab1_files = [f for f in zip_contents if f.endswith('.ab1')]

               self.log(f"DEBUG: Starting match process for {len(order_items)} expected items vs {len(zip_ab1_files)} AB1 files in zip")
               self.log(f"DEBUG: Expected items: {[item['adjusted_name'] for item in order_items]}")
               self.log(f"DEBUG: Zip AB1 files: {zip_ab1_files}")

               # Track which items have been matched
               matched_order_indices = set()
               matched_zip_indices = set()

               # Find matches using index-based approach to avoid list modification issues
               for order_idx, order_item in enumerate(order_items):
                    if order_idx in matched_order_indices:
                         continue

                    adjusted_name = order_item['adjusted_name']
                    raw_name = order_item['raw_name']
                    self.log(f"DEBUG: Looking for match for '{adjusted_name}'...")

                    for zip_idx, zip_item in enumerate(zip_ab1_files):
                         if zip_idx in matched_zip_indices:
                              continue

                         # Use customer-specific normalization consistently
                         clean_zip_item = self.file_dao.standardize_for_customer_files(zip_item, remove_extension=True)

                         self.log(f"DEBUG:   Checking against: '{zip_item}' -> '{clean_zip_item}'")

                         if clean_zip_item == adjusted_name:  # Match found
                              self.log(f"DEBUG:   ! MATCH FOUND! '{raw_name}' matches '{zip_item}'")
                              validation_result['matches'].append({
                                   'raw_name': raw_name,
                                   'file_name': zip_item
                              })
                              validation_result['match_count'] += 1

                              # Mark both items as matched
                              matched_order_indices.add(order_idx)
                              matched_zip_indices.add(zip_idx)
                              break

               # Record unmatched order items as mismatches
               for order_idx, order_item in enumerate(order_items):
                    if order_idx not in matched_order_indices:
                         self.log(f"DEBUG: X NO MATCH found for '{order_item['raw_name']}' (adjusted: '{order_item['adjusted_name']}')")
                         validation_result['mismatches_in_order'].append({
                              'raw_name': order_item['raw_name']
                         })
                         validation_result['mismatch_count'] += 1

               # Record unmatched zip files as extra files
               for zip_idx, zip_item in enumerate(zip_ab1_files):
                    if zip_idx not in matched_zip_indices:
                         self.log(f"DEBUG: Found extra AB1 file in zip: {zip_item}")
                         validation_result['mismatches_in_zip'].append(zip_item)
                         validation_result['extra_ab1_count'] += 1
                         validation_result['mismatch_count'] += 1

               self.log(f"DEBUG: Final match summary:")
               self.log(f"DEBUG: - Expected files: {len(order_items)}")
               self.log(f"DEBUG: - Zip AB1 files: {len(zip_ab1_files)}")
               self.log(f"DEBUG: - Successful matches: {len(matched_order_indices)}")
               self.log(f"DEBUG: - Unmatched order items: {len(order_items) - len(matched_order_indices)}")
               self.log(f"DEBUG: - Extra zip files: {len(zip_ab1_files) - len(matched_zip_indices)}")

               # Check for text files
               txt_extensions = self.config.TEXT_FILES
               for txt_ext in txt_extensions:
                    for zip_item in zip_contents:
                         if zip_item.endswith(txt_ext):
                              validation_result['txt_count'] += 1
                              validation_result['txt_files'].append(txt_ext)
                              break  # Only count each extension once

               return validation_result

          except Exception as e:
               self.log(f"Error validating zip file {zip_path}: {e}")
               return None

     def process_fb_pcr_zip(self, zip_path, pcr_number, order_number, version):
          """
          Process FB-PCR zip file to count the total number of files

          Args:
               zip_path (str): Path to the FB-PCR zip file
               pcr_number (str): PCR number extracted from filename
               order_number (str): Order number extracted from filename
               version (str): Version number (e.g., '1', '2', etc.)

          Returns:
               dict: Results with file count and other metadata
          """
          import zipfile

          # Results to return
          fb_pcr_result = {
               'pcr_number': pcr_number,
               'order_number': order_number,
               'version': version,
               'total_files': 0,
               'ab1_count': 0,
               'txt_count': 0,
               'file_types': {},
               'file_names': [],
               'zip_name': Path(zip_path).name
          }

          try:
               # Get all files from zip
               with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_contents = zip_ref.namelist()

               fb_pcr_result['total_files'] = len(zip_contents)
               fb_pcr_result['file_names'] = zip_contents.copy()

               # Count files by extension for detailed breakdown
               for file_name in zip_contents:
                    ext = Path(file_name).suffix.lower()
                    if ext:
                         fb_pcr_result['file_types'][ext] = fb_pcr_result['file_types'].get(ext, 0) + 1
                         # Count .ab1 files specifically
                         if ext == '.ab1':
                              fb_pcr_result['ab1_count'] += 1
                         # Count text files
                         elif ext == '.txt':
                              fb_pcr_result['txt_count'] += 1
                    else:
                         fb_pcr_result['file_types']['(no extension)'] = fb_pcr_result['file_types'].get('(no extension)', 0) + 1

               self.log(f"FB-PCR zip processed: {fb_pcr_result['zip_name']} - {fb_pcr_result['total_files']} files total, {fb_pcr_result['ab1_count']} .ab1 files")
               self.log(f"File types: {fb_pcr_result['file_types']}")

               return fb_pcr_result

          except Exception as e:
               self.log(f"Error processing FB-PCR zip file {zip_path}: {e}")
               return None

     def process_plate_zip(self, zip_path, plate_number, description):
          """
          Process plate folder zip file to count the total number of files

          Args:
               zip_path (str): Path to the plate folder zip file
               plate_number (str): Plate number extracted from filename
               description (str): Description extracted from filename

          Returns:
               dict: Results with file count and other metadata
          """
          import zipfile

          # Results to return
          plate_result = {
               'plate_number': plate_number,
               'description': description,
               'total_files': 0,
               'ab1_count': 0,
               'fsa_count': 0,
               'txt_count': 0,
               'file_types': {},
               'file_names': [],
               'zip_name': Path(zip_path).name
          }

          try:
               # Get all files from zip
               with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_contents = zip_ref.namelist()

               plate_result['total_files'] = len(zip_contents)
               plate_result['file_names'] = zip_contents.copy()

               # Count files by extension for detailed breakdown
               for file_name in zip_contents:
                    ext = Path(file_name).suffix.lower()
                    if ext:
                         plate_result['file_types'][ext] = plate_result['file_types'].get(ext, 0) + 1
                         # Count .ab1 files specifically
                         if ext == '.ab1':
                              plate_result['ab1_count'] += 1
                         # Count .fsa files specifically
                         elif ext == '.fsa':
                              plate_result['fsa_count'] += 1
                         # Count text files
                         elif ext == '.txt':
                              plate_result['txt_count'] += 1
                    else:
                         # Files without extension
                         plate_result['file_types']['(no extension)'] = plate_result['file_types'].get('(no extension)', 0) + 1

               self.log(f"Plate zip processed: {plate_result['zip_name']} - {plate_result['total_files']} files total, {plate_result['ab1_count']} .ab1 files, {plate_result['fsa_count']} .fsa files")
               self.log(f"File types: {plate_result['file_types']}")

               return plate_result

          except Exception as e:
               self.log(f"Error processing plate zip file {zip_path}: {e}")
               return None

     def _try_delete_if_empty(self, folder_path, depth=0):
          """
          Recursively check if a folder is empty or contains only empty folders,
          and delete it if so.

          Args:
               folder_path (pathlib.Path): Path to the folder to check
               depth (int): Current recursion depth (for logging)

          Returns:
               bool: True if the folder was empty or successfully deleted, False otherwise
          """
          if not folder_path.exists():
               return True

          # Check if this folder should be preserved
          folder_name_lower = folder_path.name.lower()
          if folder_name_lower.startswith('ind') or folder_name_lower.startswith('fb-pcr'):
               self.log(f"Preserving protected folder: {folder_path.name}")
               return False

          # Get list of items in the folder
          try:
               items = list(folder_path.iterdir())
          except Exception as e:
               self.log(f"Error listing directory {folder_path}: {e}")
               return False

          # If folder is already empty, try to delete it
          if not items:
               try:
                    folder_path.rmdir()
                    self.log(f"Deleted empty folder in final cleanup: {folder_path}")
                    return True
               except Exception as e:
                    self.log(f"Failed to delete empty folder {folder_path}: {e}")
                    return False

          # Check if all items are directories that can be deleted
          all_removable = True
          for item in items:
               if item.is_dir():
                    # Recursively check if this subdirectory can be deleted
                    if not self._try_delete_if_empty(item, depth + 1):
                         all_removable = False
               else:
                    # It's a file, so this directory cannot be removed
                    all_removable = False

          # If all subdirectories were removed, try to delete this directory too
          if all_removable:
               # Check again if the folder is now empty
               if not list(folder_path.iterdir()):
                    try:
                         folder_path.rmdir()
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
               root_folder (pathlib.Path): Root folder to start cleanup from
          """
          self.log(f"Running final cleanup on: {root_folder}")

          # Get all folders in the root directory
          folders_to_check = []
          for item in root_folder.iterdir():
               if item.is_dir():
                    folders_to_check.append(item)

          # Process each folder to see if it's empty or contains only empty folders
          for folder in folders_to_check:
               self._try_delete_if_empty(folder)

          self.log("Final cleanup complete")




if __name__ == "__main__":
    # Simple test if run directly
    from mseqauto.config import MseqConfig # type: ignore
    from file_system_dao import FileSystemDAO
    from datetime import datetime

    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    processor = FolderProcessor(file_dao, None, config)

    # Test folder processing
    test_path = str(Path.cwd())
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
