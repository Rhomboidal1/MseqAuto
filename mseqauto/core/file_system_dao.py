# file_system_dao.py
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import re
from datetime import datetime, timedelta
from shutil import move, copyfile
from zipfile import ZipFile, ZIP_DEFLATED
from mseqauto.config import MseqConfig

config = MseqConfig()

class FileSystemDAO:
    def __init__(self, config, logger=None):
        self.config = config
        self.directory_cache = {}
        
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
    # Directory Operations
    def get_directory_contents(self, path, refresh=False): #KEEP
        """Get directory contents with caching"""
        if path not in self.directory_cache or refresh:
            if not os.path.exists(path):
                self.directory_cache[path] = []
                return []

            try:
                contents = os.listdir(path)
                self.directory_cache[path] = contents
            except Exception as e:
                print(f"Error reading directory {path}: {e}")
                self.directory_cache[path] = []

        return self.directory_cache[path]

    def get_folders(self, path, pattern=None): #KEEP
        """
        Get folders matching an optional regex pattern
        
        Args:
            path (str): Path to search for folders
            pattern (str or re.Pattern): A regex pattern string or a compiled regex object
            
        Returns:
            list: List of folder paths matching the pattern
        """
        folder_list = []
        self.log(f"Searching for folders in: {path}")

    
        # If pattern is a compiled regex, print its pattern attribute
        if hasattr(pattern, 'pattern'):
            print(f"Regex pattern string: {pattern.pattern}")
        self.log(f"Using pattern: {pattern}")

        contents = self.get_directory_contents(path)
        self.log(f"Found {len(contents)} items in directory")
        
        for item in contents:
            full_path = os.path.join(path, item)
            
            if os.path.isdir(full_path):
                if pattern is None:
                    self.log(f"  No pattern provided, adding {item}")  # Changed from print to self.log
                    folder_list.append(full_path)
                else:
                    # Check if pattern is a compiled regex object or a string
                    if hasattr(pattern, 'search'):  # Compiled regex
                        match = pattern.search(item.lower())
                        match_method = "using compiled regex search"
                    else:  # String pattern
                        match = re.search(pattern, item.lower())
                        match_method = "using re.search"
                    
                    if match:
                        folder_list.append(full_path)
                    else:
                        self.log(f"  No match for {item}")  # Changed from print to self.log
        
        self.log(f"Returning {len(folder_list)} matching folders: {[os.path.basename(f) for f in folder_list]}")  # Changed from print to self.log
        return folder_list

    def get_files_by_extension(self, folder, extension): #KEEP
        """Get all files with specified extension in a folder"""
        files = []
        for item in self.get_directory_contents(folder):
            if item.endswith(extension):
                files.append(os.path.join(folder, item))
        return files

    def contains_file_type(self, folder, extension): #KEEP
        """Check if folder contains files with specified extension"""
        for item in self.get_directory_contents(folder):
            if item.endswith(extension):
                return True
        return False

    def create_folder_if_not_exists(self, path): #KEEP
        """Create folder if it doesn't exist"""
        if not os.path.exists(path):
            os.mkdir(path)
        return path

    def move_folder(self, source, destination, max_retries=3, delay=1.0): #KEEP
        """
        Move folder with proper error handling and retries
        
        Args:
            source: Source folder path
            destination: Destination folder path
            max_retries: Maximum number of retries (default 3)
            delay: Delay between retries in seconds (default 1.0)
            
        Returns:
            bool: True if successful, False otherwise
        """
        from shutil import move
        import time
        import gc
        
        # Make sure destination parent folder exists
        dest_parent = os.path.dirname(destination)
        if not os.path.exists(dest_parent):
            try:
                os.makedirs(dest_parent)
            except Exception as e:
                self.warning(f"Error creating destination parent folder: {e}")
                return False
        
        # Try multiple times with delay between attempts
        for attempt in range(max_retries):
            try:
                # Force garbage collection to release file handles
                gc.collect()
                
                # Attempt to move the folder
                move(source, destination)
                self.log(f"Successfully moved {os.path.basename(source)} to {os.path.basename(destination)}")
                #print(f"Successfully moved {os.path.basename(source)} to {os.path.basename(destination)}")
                return True
            except Exception as e:
                self.warning(f"Error moving folder on attempt {attempt+1}/{max_retries}: {e}")
                #print(f"Error moving folder on attempt {attempt+1}/{max_retries}: {e}")
                # Wait before next attempt
                if attempt < max_retries - 1:
                    time.sleep(delay)
        
        self.error(f"Failed to move folder after {max_retries} attempts: {os.path.basename(source)}")
        return False

    def get_folder_name(self, path): #Rename to get_basename and move to path_utilities.py
        """Get the folder name from a path"""
        return os.path.basename(path)

    def get_parent_folder(self, path): 
        #Rename to get_dirname and move to path_utilities.py
        """Get the parent folder path"""
        return os.path.dirname(path)

    def join_paths(self, base_path, *args): 
        #Move to path_utilities.py
        """Join path components"""
        return os.path.join(base_path, *args)
    
    #################### Data Loading ####################

    def load_order_key(self, key_file_path): 
        #Keep
        """Load the order key file"""
        try:
            import numpy as np
            return np.loadtxt(key_file_path, dtype=str, delimiter='\t')
        except Exception as e:
            print(f"Error loading order key file: {e}")
            return None
        
    #################### File Existence Checks ####################
    def file_exists(self, path): 
        #KEEP
        """Check if a file exists"""
        return os.path.isfile(path)

    def folder_exists(self, path): 
        #KEEP
        """Check if a folder exists"""
        return os.path.isdir(path)
    
    #################### File Statistics ####################
    def count_files_by_extensions(self, folder, extensions):
        """Count files with specific extensions in a folder"""
        counts = {ext: 0 for ext in extensions}
        for item in self.get_directory_contents(folder):
            file_path = os.path.join(folder, item)
            if os.path.isfile(file_path):
                for ext in extensions:
                    if item.endswith(ext):
                        counts[ext] += 1
        return counts

    def get_folder_creation_time(self, folder):
        """Get the creation time of a folder"""
        return os.path.getctime(folder)

    def get_folder_modification_time(self, folder):
        """Get the last modification time of a folder"""
        return os.path.getmtime(folder)

    def clean_braces_format(self, file_name): 
        #Move to path_utilities.py
        """Remove anything contained in {} from filename"""
        return re.sub(r'{.*?}', '', self.neutralize_suffixes(file_name))

    def adjust_abi_chars(self, file_name): 
        #move to path_utilities.py
        """Adjust characters in file name to match ABI naming conventions"""
        # Create translation table
        translation_table = str.maketrans({
            ' ': '',
            '+': '&',
            '*': '-',
            '|': '-',
            '/': '-',
            '\\': '-',
            ':': '-',
            '"': '',
            "'": '',
            '<': '-',
            '>': '-',
            '?': '',
            ',': ''
        })

        # Apply translation
        return file_name.translate(translation_table)

    def normalize_filename(self, file_name, remove_extension=True, logger=None):
        #Move to path_utilities.py
        """Normalize filename with optional logging"""
        # Step 1: Adjust characters
        adjusted_name = self.adjust_abi_chars(file_name)

        # Step 2: Remove extension if needed (specific handling for .ab1)
        if remove_extension:
            if adjusted_name.endswith(config.ABI_EXTENSION):
                # Special handling for .ab1 files
                adjusted_name = adjusted_name[:-4]
            elif '.' in adjusted_name:
                # Legacy behavior for other extensions
                name_without_ext = adjusted_name[:adjusted_name.rfind('.')]
                adjusted_name = name_without_ext

        # Step 3: Remove suffixes
        neutralized_name = self.neutralize_suffixes(adjusted_name)

        # Step 4: Remove brace content
        cleaned_name = re.sub(r'{.*?}', '', neutralized_name)

        # Only log if a logger is provided
        if logger:
            logger(f"Normalized '{file_name}' to '{cleaned_name}'")

        return cleaned_name

    def neutralize_suffixes(self, file_name):
        #Move to path_utilities.py
        """Remove suffixes like _Premixed and _RTI"""
        new_file_name = file_name
        new_file_name = new_file_name.replace('_Premixed', '')
        new_file_name = new_file_name.replace('_RTI', '')
        return new_file_name

    def remove_extension(self, file_name, extension=None):
        #Move to FileProcessor
        """Remove file extension"""
        if extension and file_name.endswith(extension):
            return file_name[:-len(extension)]
        return os.path.splitext(file_name)[0]

    #################### Zip operations ####################
    def check_for_zip(self, folder_path):
        #Keep
        """Check if folder contains any zip files"""
        for item in self.get_directory_contents(folder_path):
            file_path = os.path.join(folder_path, item)
            if os.path.isfile(file_path) and file_path.endswith(self.config.ZIP_EXTENSION):
                return True
        return False

    def zip_files(self, source_folder: str, zip_path: str, file_extensions=None, exclude_extensions=None):
        #Keep
        """Create a zip file from files in source_folder matching extensions"""
        with ZipFile(zip_path, 'w') as zip_file:
            for item in self.get_directory_contents(source_folder):
                # Convert item to string if needed
                item_str = str(item)
                file_path = os.path.join(source_folder, item_str)
                if not os.path.isfile(file_path):
                    continue
                if file_extensions and not any(item_str.endswith(ext) for ext in file_extensions):
                    continue

                if exclude_extensions and any(item_str.endswith(ext) for ext in exclude_extensions):
                    continue

                zip_file.write(file_path, arcname=item_str, compress_type=ZIP_DEFLATED)

        return True

    def get_zip_contents(self, zip_path):
        #Keep
        """Get list of files in a zip archive"""
        try:
            with ZipFile(zip_path, 'r') as zip_ref:
                return zip_ref.namelist()
        except Exception as e:
            print(f"Error reading zip file {zip_path}: {e}")
            return []

    def copy_zip_to_dump(self, zip_path, dump_folder):
        #Keep
        """Copy zip file to dump folder"""
        if not os.path.exists(dump_folder):
            os.makedirs(dump_folder)

        dest_path = os.path.join(dump_folder, os.path.basename(zip_path))
        copyfile(zip_path, dest_path)
        return dest_path

    
    def get_pcr_number(self, filename):
        #Move to path_utilities.py-done
        """Extract PCR number from file name"""
        pcr_num = ''
        # Look for PCR pattern in brackets - only this matters for folder sorting
        if re.search('{pcr\\d+.+}', filename.lower()):
            pcr_bracket = re.search("{pcr\\d+.+}", filename.lower()).group()
            pcr_num = re.search("pcr(\\d+)", pcr_bracket).group(1)
        return pcr_num

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
    #################### Advanced Directory Operations ####################
    def get_inumber_from_name(self, name):
        #I think we should keep
        """Extract I number from a name using precompiled regex"""
        match = self.regex_patterns['inumber'].search(str(name).lower())
        if match:
            return match.group(1)  # Return just the number
        return None
    
    def get_recent_files(self, paths, days=None, hours=None):
        #Keep
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
        return [file_info[0] for file_info in sorted_files]
    
    def collect_active_inumbers(self, paths, days=5, hours=12, min_inum=None, 
                                return_most_recent=False, return_files=False):
        """
        Comprehensive method to collect active I-numbers with various filtering options
        
        Args:
            paths (list): List of paths to search for files
            days (int, optional): Include files modified in the last N days (default: 5)
            hours (int, optional): Include files modified in the last N hours (default: 12)
            min_inum (str, optional): Only include I-numbers greater than this value
            return_most_recent (bool): If True, return only the most recent I-number
            return_files (bool): If True, return the files instead of just I-numbers
        
        Returns:
            list: Deduplicated list of I-numbers or files meeting criteria
        """
        collected_items = {}  # Use dict to track inums and their source files
        
        # Get files from the past N days
        if days:
            day_files = self.get_recent_files(paths, days=days)
            for file in day_files:
                inum = self.get_inumber_from_name(file)
                if inum and (min_inum is None or int(inum) > int(min_inum)):
                    if inum not in collected_items:
                        collected_items[inum] = []
                    collected_items[inum].append(file)
        
        # Get files from the past N hours
        if hours:
            hour_files = self.get_recent_files(paths, hours=hours)
            for file in hour_files:
                inum = self.get_inumber_from_name(file)
                if inum:
                    if inum not in collected_items:
                        collected_items[inum] = []
                    collected_items[inum].append(file)
        
        # Handle return options
        if return_most_recent and collected_items:
            # Find most recent I-number (highest number assuming sequential assignment)
            most_recent = max(collected_items.keys(), key=lambda x: int(x))
            if return_files:
                return collected_items[most_recent]
            return [most_recent]
        
        if return_files:
            # Flatten the file list
            all_files = []
            for files in collected_items.values():
                all_files.extend(files)
            return all_files
        
        # Return just the I-numbers
        return list(collected_items.keys())

    def get_most_recent_inumber(self, path):
        """Find the most recent I number based on folder modification times"""
        try:
            folders = self.get_directory_contents(path)
            if not folders:
                return None
                
            # Sort folders by modification time (newest first)
            sorted_folders = sorted(
                [f for f in folders if os.path.isdir(os.path.join(path, f))],
                key=lambda f: os.path.getmtime(os.path.join(path, f)), 
                reverse=True
            )
            
            # Extract I number from the most recent folder
            if sorted_folders:
                return self.get_inumber_from_name(sorted_folders[0])
            return None
        except Exception as e:
            print(f"Error getting most recent I number: {e}")
            return None

    def get_folders_with_inumbers(self, path, folder_pattern='bioi_folder', exclude_patterns=None):
        """
        Get folders matching a specified pattern and extract their I-numbers
        
        Args:
            path (str): Directory to search
            folder_pattern (str): Regex pattern key to match folders (default: 'bioi_folder')
            exclude_patterns (list): List of patterns to exclude (default: ['reinject'])
        
        Returns:
            tuple: (list of I-numbers, list of folder paths)
        """
        if exclude_patterns is None:
            exclude_patterns = ['reinject']
            
        matching_folders = []
        
        # Scan the directory for matching folders
        for item in self.get_directory_contents(path):
            item_path = os.path.join(path, item)
            if not os.path.isdir(item_path):
                continue
                
            # Check if folder matches pattern
            if self.regex_patterns[folder_pattern].search(item):
                # Check exclusion patterns
                if not any(exclude in item.lower() for exclude in exclude_patterns):
                    matching_folders.append(item_path)
        
        # Extract unique I-numbers
        i_numbers = []
        for folder in matching_folders:
            i_num = self.get_inumber_from_name(os.path.basename(folder))
            if i_num and i_num not in i_numbers:
                i_numbers.append(i_num)
        
        return i_numbers, matching_folders

    #################### File Operations ####################
    def move_file(self, source, destination):
        #Keep
        """Move a file with error handling"""
        try:
            move(source, destination)
            return True
        except Exception as e:
            print(f"Error moving file {source}: {e}")
            return False

    def rename_file_without_braces(self, file_path):
        #Keep
        """Rename a file to remove anything in braces"""
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

    def standardize_filename_for_matching(self, file_name, remove_extension=True):
        #Move to path_utilities.py
        """
        Standardized filename cleaning method for consistent matching across all code
        Used for PCR files, reinject lists, and any other string comparison operations
        """
        # Step 1: Remove file extension if needed
        if remove_extension and file_name.endswith(config.ABI_EXTENSION):
            clean_name = file_name[:-4]
        else:
            clean_name = file_name

        # Step 2: Remove content in brackets (PCR numbers, well locations)
        clean_name = re.sub(r'{.*?}', '', clean_name)

        # Step 3: Remove standard suffixes
        clean_name = clean_name.replace('_Premixed', '')
        clean_name = clean_name.replace('_RTI', '')

        # Step 4: Character standardization (only if needed for PCR matching)
        # Note: For PCR files we generally don't need this step as we want exact matches
        # clean_name = self.adjust_abi_chars(clean_name)

        return clean_name


if __name__ == "__main__":
    # Simple test if run directly
    from mseqauto.config import MseqConfig

    config = MseqConfig()
    dao = FileSystemDAO(config)

    # Test folder operations
    test_path = os.getcwd()
    folders = dao.get_folders(test_path)
    print(f"Found {len(folders)} folders in {test_path}")

    # Test file operations
    py_files = dao.get_files_by_extension(test_path, ".py")
    print(f"Found {len(py_files)} Python files in {test_path}")

    ab1_files = dao.get_files_by_extension(test_path, ".ab1")
    print(f"Found {len(ab1_files)} Python files in {test_path}")
