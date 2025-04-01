# Description: FileProcessor class for processing files with business logic
import re
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
from datetime import datetime, timedelta
#print(sys.path)

from mseqauto.config import MseqConfig

config = MseqConfig()


class FileProcessor:
     def __init__(self, config_obj):
          self.config = config_obj
          self.directory_cache = {}

          # Precompiled regex patterns
          self.regex_patterns = {
                    'inumber': re.compile(r'bioi-(\d+)', re.IGNORECASE),
                    'pcr_number': re.compile(r'{pcr(\d+).+}', re.IGNORECASE),
                    'brace_content': re.compile(r'{.*?}'),
                    'bioi_folder': re.compile(r'.+bioi-\d+.+', re.IGNORECASE),
                    'ind_blank_file': re.compile(r'{\d+[A-H]}.ab1$', re.IGNORECASE),  # Individual blanks pattern
                    'plate_blank_file': re.compile(r'^\d{2}[A-H]__.ab1$', re.IGNORECASE)  # Plate blanks pattern
                    # Add other patterns as needed
               }
    
     def is_control_file(self, file_name, control_list):
          """Check if file is a control sample"""
          clean_name = self.clean_braces_format(file_name)
          clean_name = self.remove_extension(clean_name)
          return clean_name.lower() in [control.lower() for control in control_list]

     def is_blank_file(self, file_name):
          """
          Check if file is a blank sample
          Handles two formats:
          1. Individual sequencing blanks: {07H}.ab1, {11F}.ab1
          2. Plate sequencing blanks: 01A__.ab1, 02B__.ab1
          """
          # Check for individual sequencing blanks pattern {[digits][letter]}.ab1
          if self.regex_patterns['ind_blank_file'].match(file_name):
               return True

          # Check for plate sequencing blanks pattern [digits][letter]__.ab1
          if self.regex_patterns['plate_blank_file'].match(file_name):
               return True

          return False

     def remove_extension(self, file_name, extension=None):
          """Remove file extension"""
          if extension and file_name.endswith(extension):
               return file_name[:-len(extension)]
          return os.path.splitext(file_name)[0]

     '''def get_most_recent_inumber(self, path):
          """Find the most recent I number based on folder modification times"""
          try:
               # Current timestamp and cutoff (7 days ago)
               current_timestamp = datetime.now().timestamp()
               cutoff_timestamp = current_timestamp - (7 * 24 * 3600)

               recent_dirs = []  # Changed from 'folders' to 'recent_dirs'

               # Get folders modified in the last 7 days
               with os.scandir(path) as entries:
                    for entry in entries:
                    if entry.is_dir():
                         last_modified_timestamp = entry.stat().st_mtime
                         if last_modified_timestamp >= cutoff_timestamp:
                              recent_dirs.append(entry.name)

               # Sort folders by modification time (newest first)
               sorted_dirs = sorted(recent_dirs, key=lambda f: os.path.getmtime(os.path.join(path, f)), reverse=True)

               # Extract I number from the most recent folder
               if sorted_dirs:
                    inum = self.get_inumber_from_name(sorted_dirs[0])
                    return inum

               return None
          except Exception as e:
               print(f"Error getting most recent I number: {e}")
               return None'''
          
     '''def get_inumber_from_name(self, name):
          #I think we should keep
          """Extract I number from a name using precompiled regex"""
          match = self.regex_patterns['inumber'].search(str(name).lower())
          if match:
               return match.group(1)  # Return just the number
          return None'''    
           
     '''def get_recent_files(self, paths, days=None, hours=None):
          """Get list of files modified within specified time period"""
          # Set cutoff date based on days or hours
          current_date = datetime.now()
          if days:
               cutoff_date = current_date - timedelta(days=days)
          elif hours:
               cutoff_date = current_date - timedelta(hours=hours)
          else:
               cutoff_date = current_date - timedelta(days=1)  # Default to 1 day

          # cutoff_timestamp = cutoff_date.timestamp()

          # Collect recent files from all specified paths
          file_info_list = []
          for directory in paths:
               try:
                    with os.scandir(directory) as entries:
                    for entry in entries:
                         if entry.is_file():
                              last_modified_timestamp = entry.stat().st_mtime
                              last_modified_date = datetime.fromtimestamp(last_modified_timestamp)

                              if last_modified_date >= cutoff_date and entry.name.endswith('.txt'):
                                   file_info_list.append((entry.name, last_modified_timestamp))
               except Exception as e:
                    print(f"Error scanning directory {directory}: {e}")

          # Sort by modification time (newest first)
          sorted_files = sorted(file_info_list, key=lambda x: x[1], reverse=True)

          # Return just the file names
          return [file_info[0] for file_info in sorted_files]'''
     
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