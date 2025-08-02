# file_system_dao.py
import os
import sys
from pathlib import Path
sys.path.append(str(Path(__file__).parents[2]))

import re
from datetime import datetime, timedelta
from shutil import move, copyfile
from zipfile import ZipFile, ZIP_DEFLATED
from mseqauto.config import MseqConfig  # type: ignore

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
            'plate_blank_file': re.compile(r'^\d{2}[A-H]__.ab1$', re.IGNORECASE),  # Plate blanks pattern
            'plate_well_prefix': re.compile(r'^\d{2}[a-z]_(.+)$', re.IGNORECASE),  # NEW: Plate well position pattern
            'fb_pcr_zip': re.compile(r'^FB-PCR\d+_\d+(?:_\d+)?\.zip$', re.IGNORECASE),  # FB-PCR zip pattern
            'plate_folder': re.compile(r'^P\d{5}_.*')  # Plate folder pattern P12345_Anythinghere
            # Add other patterns as needed
        }

    # Directory Operations
    def get_directory_contents(self, path, refresh=False):
        """Get directory contents with caching"""
        path = Path(path)  # Convert to Path object

        if path not in self.directory_cache or refresh:
            if not path.exists():  # Remove duplicate parameter
                self.directory_cache[path] = []
                return []
            try:
                contents = list(path.iterdir())
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
            print(f"Regex pattern string: {pattern.pattern}") #type: ignore
        self.log(f"Using pattern: {pattern}")

        contents = self.get_directory_contents(path)
        self.log(f"Found {len(contents)} items in directory")

        for item in contents:
            full_path = Path(path) / item

            if full_path.is_dir():
                if pattern is None:
                    self.log(f"  No pattern provided, adding {item}")  # Changed from print to self.log
                    folder_list.append(full_path)
                else:
                    # Check if pattern is a compiled regex object or a string
                    if hasattr(pattern, 'search'):  # Compiled regex
                        match = pattern.search(item.lower())
                        match_method = "using compiled regex search"
                    else:  # String pattern
                        match = re.search(pattern, item.name.lower())
                        match_method = "using re.search"

                    if match:
                        folder_list.append(full_path)
                    else:
                        self.log(f"  No match for {item}")  # Changed from print to self.log

        self.log(f"Returning {len(folder_list)} matching folders: {[f.name for f in folder_list]}")  # Changed from print to self.log
        return folder_list

    def get_files_by_extension(self, folder_path_str, extension, recursive=False):
        """
        Get all files with a specific extension in a folder

        Args:
            folder_path_str (str): Path to the folder
            extension (str): File extension to look for (e.g., '.ab1')
            recursive (bool): Whether to search recursively in subfolders

        Returns:
            list: List of file paths with the specified extension
        """
        folder_path = Path(folder_path_str)
        if recursive:
            # Use pathlib.rglob to get all files recursively
            files = []
            for item in folder_path.rglob(f'*{extension}'):
                if item.is_file():
                    files.append(str(item))
            return files
        else:
            # Just search in the current folder
            try:
                return [str(f) for f in folder_path.iterdir()
                    if f.is_file() and
                    f.name.lower().endswith(extension.lower())]
            except Exception as e:
                self.log(f"Error getting files by extension {extension} in {folder_path_str}: {e}")
                return []

    def contains_file_type(self, folder, extension): #KEEP
        """Check if folder contains files with specified extension"""
        for item in self.get_directory_contents(folder):
            if item.name.endswith(extension):
                return True
        return False

    def create_folder_if_not_exists(self, path_str): #KEEP
        """Create folder if it doesn't exist"""
        path = Path(path_str)
        if not path.exists():
            path.mkdir()
        return str(path)

    def move_folder(self, source, destination, max_retries=3, delay=0.1):
        """
        Move folder with proper error handling and retries

        Args:
            source: Source folder path
            destination: Destination folder path
            max_retries: Maximum number of retries (default 3)
            delay: Delay between retries in seconds (default 0.1)

        Returns:
            bool: True if successful, False otherwise
        """
        import time
        import shutil
        from pathlib import Path

        source_path = Path(source)
        destination_path = Path(destination)

        # Ensure the parent directory exists
        destination_path.parent.mkdir(parents=True, exist_ok=True)

        # Simple move operation with retries
        for retry in range(max_retries):
            try:
                # Use shutil.move for the actual move operation
                shutil.move(str(source_path), str(destination_path))
                self.log(f"Successfully moved {source_path.name} to {destination_path}")
                return True

            except Exception as e:
                self.warning(f"Retry {retry+1}/{max_retries}: Error moving folder: {e}")
                if retry < max_retries - 1:
                    time.sleep(delay)
                else:
                    self.error(f"Failed to move folder after {max_retries} attempts: {e}")
                    return False

        return False

    def get_folder_name(self, path_str): #Rename to get_basename and move to path_utilities.py
        """Get the folder name from a path"""
        return Path(path_str).name

    def get_parent_folder(self, path_str):
        #Rename to get_dirname and move to path_utilities.py
        """Get the parent folder path"""
        return str(Path(path_str).parent)

    def join_paths(self, base_path_str, *args):
        #Move to path_utilities.py
        """Join path components"""
        return str(Path(base_path_str, *args))

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
    def file_exists(self, path_str):
        #KEEP
        """Check if a file exists"""
        return Path(path_str).is_file()

    def folder_exists(self, path_str):
        #KEEP
        """Check if a folder exists"""
        return Path(path_str).is_dir()

    #################### File Statistics ####################
    def count_files_by_extensions(self, folder, extensions):
        """Count files with specific extensions in a folder"""
        counts = {ext: 0 for ext in extensions}
        folder_path = Path(folder)
        for item in self.get_directory_contents(folder):
            file_path = folder_path / item
            if file_path.is_file():
                for ext in extensions:
                    if item.name.endswith(ext):
                        counts[ext] += 1
        return counts

    def get_folder_creation_time(self, folder):
        """Get the creation time of a folder"""
        return Path(folder).stat().st_ctime

    def get_folder_modification_time(self, folder):
        """Get the last modification time of a folder"""
        return Path(folder).stat().st_mtime

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

    def extract_order_number_from_filename(self, filename):
        """Extract order number from filename braces if present"""
        # Look for exactly 6-digit numbers in braces, which represent order numbers
        matches = re.findall(r'{(\d{6})}', filename)
        if matches:
            # Return the first matching 6-digit number
            return matches[0]
        return None

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
        # Convert to string if it's a Path object
        if hasattr(file_name, 'name'):
            file_name_str = file_name.name
        else:
            file_name_str = str(file_name)

        if extension and file_name_str.endswith(extension):
            return file_name_str[:-len(extension)]
        return Path(file_name_str).stem

    #################### Zip operations ####################
    def check_for_zip(self, folder_path):
        #Keep
        """Check if folder contains any zip files"""
        folder_path_obj = Path(folder_path)
        for item in self.get_directory_contents(folder_path):
            file_path = folder_path_obj / item
            if file_path.is_file() and item.name.endswith(self.config.ZIP_EXTENSION):
                return True
        return False

    def zip_files(self, source_folder: str, zip_path: str, file_extensions=None, exclude_extensions=None):
        #Keep
        """Create a zip file from files in source_folder matching extensions"""
        source_folder_path = Path(source_folder)
        with ZipFile(zip_path, 'w') as zip_file:
            for item in self.get_directory_contents(source_folder):
                file_path = source_folder_path / item
                if not file_path.is_file():
                    continue
                if file_extensions and not any(item.name.endswith(ext) for ext in file_extensions):
                    continue

                if exclude_extensions and any(item.name.endswith(ext) for ext in exclude_extensions):
                    continue

                zip_file.write(file_path, arcname=item.name, compress_type=ZIP_DEFLATED)

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
        dump_folder_path = Path(dump_folder)
        if not dump_folder_path.exists():
            dump_folder_path.mkdir(parents=True)

        dest_path = dump_folder_path / Path(zip_path).name
        copyfile(zip_path, dest_path)
        return str(dest_path)

    def find_recent_zips(self, folder_path, max_age_minutes=15):
        """
        Finds recently created or modified zip files within a specified folder.

        Args:
            folder_path (str): The path to the folder to search for zip files.
            max_age_minutes (int): The maximum age in minutes for a zip file to be considered recent.

        Returns:
            list: A list of Path objects representing the paths to recently created zip files.
        """
        from datetime import datetime, timedelta

        folder_path_obj = Path(folder_path)
        self.log(f"Checking for recently created zip files in {folder_path_obj.name}")

        # Calculate the cutoff time for determining recent files
        now = datetime.now()
        cutoff_time = now - timedelta(minutes=max_age_minutes)
        recent_zips = []

        # Ensure the folder exists before proceeding
        if not folder_path_obj.exists():
            return recent_zips

        # Iterate through items in the specified folder
        for item in self.get_directory_contents(folder_path):
            # Check if the item is a zip file based on the configured extension
            if item.name.endswith(self.config.ZIP_EXTENSION):
                zip_path_obj = folder_path_obj / item

                # Skip if the item is not a file
                if not zip_path_obj.is_file():
                    continue

                # Get the modification time of the zip file
                modified_time = datetime.fromtimestamp(zip_path_obj.stat().st_mtime)

                # Check if the file's modification time is within the specified age limit
                if modified_time >= cutoff_time:
                    recent_zips.append(zip_path_obj)
                    self.debug(f"Found recent zip: {item}")

        return recent_zips


    def copy_recent_zips_to_dump(self, folder_list, zip_dump_folder, max_age_minutes=15):
        """
        Copies recently created zip files found in specified folders to a designated dump folder.

        Args:
            folder_list (list): A list of folder paths to check for zip files.
            zip_dump_folder (str): The path to the target folder where recent zip files will be copied.
            max_age_minutes (int): The maximum age in minutes for a zip file to be considered recent.

        Returns:
            int: The number of zip files successfully copied to the dump folder.
        """
        self.log(f"Checking for recently created zip files (within {max_age_minutes} minutes)")

        zip_dump_folder_obj = Path(zip_dump_folder)

        # Create the zip dump folder if it does not already exist
        if not zip_dump_folder_obj.exists():
            zip_dump_folder_obj.mkdir()

        copied_count = 0

        # Process each folder in the provided list
        for parent_folder in folder_list:
            parent_folder_obj = Path(parent_folder)

            # Special handling for folders starting with "bioi-" (assuming they contain order subfolders)
            if parent_folder_obj.name.lower().startswith("bioi-"):
                # Iterate through subfolders within the BioI folder
                for item in self.get_directory_contents(parent_folder):
                    order_path_obj = parent_folder_obj / item
                    if order_path_obj.is_dir():
                        # Find recent zip files within this order subfolder
                        recent_zips = self.find_recent_zips(order_path_obj, max_age_minutes)

                        # Copy each found recent zip file to the dump folder if it doesn't already exist
                        for zip_path in recent_zips:
                            dump_path_obj = zip_dump_folder_obj / zip_path.name
                            if not dump_path_obj.exists():
                                try:
                                    self.copy_zip_to_dump(zip_path, zip_dump_folder)
                                    copied_count += 1
                                except Exception as e:
                                    self.warning(f"Failed to copy zip to dump: {e}")
            else:
                # For other folders (e.g., PCR), find recent zip files directly
                recent_zips = self.find_recent_zips(parent_folder_obj, max_age_minutes)

                # Copy each found recent zip file to the dump folder if it doesn't already exist
                for zip_path in recent_zips:
                    dump_path_obj = zip_dump_folder_obj / zip_path.name
                    if not dump_path_obj.exists():
                        try:
                            self.copy_zip_to_dump(zip_path, zip_dump_folder)
                            copied_count += 1
                        except Exception as e:
                            self.warning(f"Failed to copy zip to dump: {e}")

        return copied_count

    def find_fb_pcr_zips(self, folder_path):
        """
        Find FB-PCR zip files matching the pattern FB-PCR####_###### or FB-PCR####_######_#
        Searches in multiple locations:
        1. Top-level folder (level 0)
        2. Immediate subfolders (level 1) - where FB-PCR zips typically are
        3. Zip dump folder if it exists

        Args:
            folder_path (str): Path to the folder to search for FB-PCR zip files

        Returns:
            list: List of tuples (zip_path, pcr_number, order_number, version) for matching FB-PCR zip files
        """
        folder_path_obj = Path(folder_path)
        fb_pcr_zips = []

        if not folder_path_obj.exists():
            return fb_pcr_zips

        def _scan_single_folder(search_folder):
            """Helper: scan one folder for FB-PCR zips"""
            local_zips = []
            if not search_folder.exists():
                return local_zips

            for item in self.get_directory_contents(search_folder):
                if item.name.endswith('.zip'):
                    # Check if it matches the FB-PCR pattern
                    match = self.regex_patterns['fb_pcr_zip'].match(item.name)
                    if match:
                        full_path = str(search_folder / item.name)

                        # Extract PCR number and order number from filename
                        # Pattern: FB-PCR3048_149727.zip or FB-PCR3048_149727_2.zip
                        name_without_ext = item.name[:-4]  # Remove .zip
                        parts = name_without_ext.split('_')

                        if len(parts) >= 2:
                            pcr_part = parts[0]  # FB-PCR3048
                            order_number = parts[1]  # 149727
                            version = parts[2] if len(parts) > 2 else '1'  # 2 or default to 1

                            # Extract PCR number from FB-PCR3048
                            pcr_number = pcr_part[6:]  # Remove "FB-PCR" prefix

                            local_zips.append((full_path, pcr_number, order_number, version))
                            self.log(f"Found FB-PCR zip: {item.name} in {search_folder} (PCR: {pcr_number}, Order: {order_number}, Version: {version})")
            return local_zips

        # 1. Check top-level folder (level 0)
        fb_pcr_zips.extend(_scan_single_folder(folder_path_obj))

        # 2. Check immediate subfolders (level 1) - where FB-PCR zips typically are
        for item in self.get_directory_contents(folder_path):
            if item.is_dir():
                subfolder_path = folder_path_obj / item.name
                fb_pcr_zips.extend(_scan_single_folder(subfolder_path))

        # 3. Check zip dump folder specifically if it exists
        zip_dump_folder = folder_path_obj / self.config.ZIP_DUMP_FOLDER
        if zip_dump_folder.exists():
            fb_pcr_zips.extend(_scan_single_folder(zip_dump_folder))

        # Remove duplicates (in case same file found in multiple locations)
        seen_files = set()
        unique_fb_pcr_zips = []
        for zip_info in fb_pcr_zips:
            zip_path = zip_info[0]
            file_name = Path(zip_path).name
            if file_name not in seen_files:
                seen_files.add(file_name)
                unique_fb_pcr_zips.append(zip_info)
            else:
                self.log(f"Skipping duplicate FB-PCR zip: {file_name}")

        return unique_fb_pcr_zips

    def find_plate_folder_zips(self, folder_path):
        """
        Find plate folder zip files matching the pattern P#####_*.zip
        Searches in multiple locations:
        1. Top-level folder (level 0)
        2. Immediate subfolders (level 1) - where plate zips typically are
        3. Zip dump folder if it exists

        Args:
            folder_path (str): Path to the folder to search for plate folder zip files

        Returns:
            list: List of tuples (zip_path, plate_number, description) for matching plate folder zip files
        """
        folder_path_obj = Path(folder_path)
        plate_zips = []

        if not folder_path_obj.exists():
            return plate_zips

        def _scan_single_folder(search_folder):
            """Helper: scan one folder for plate folder zips"""
            local_zips = []
            if not search_folder.exists():
                return local_zips

            for item in self.get_directory_contents(search_folder):
                if item.name.endswith('.zip'):
                    # Check if it matches the plate folder pattern (but for zip files)
                    # We need to remove .zip and check if the remaining name matches P#####_*
                    name_without_ext = item.name[:-4]  # Remove .zip
                    if self.regex_patterns['plate_folder'].match(name_without_ext):
                        full_path = str(search_folder / item.name)

                        # Extract plate number and description from filename
                        # Pattern: P12345_description.zip
                        parts = name_without_ext.split('_', 1)  # Split on first underscore only

                        if len(parts) >= 2:
                            plate_number = parts[0][1:]  # Remove "P" prefix (P12345 -> 12345)
                            description = parts[1]  # Everything after first underscore

                            local_zips.append((full_path, plate_number, description))
                            self.log(f"Found plate folder zip: {item.name} in {search_folder} (Plate: {plate_number}, Description: {description})")
                        else:
                            # Handle case where there's no description after underscore
                            plate_number = parts[0][1:]  # Remove "P" prefix
                            description = ""
                            local_zips.append((full_path, plate_number, description))
                            self.log(f"Found plate folder zip: {item.name} in {search_folder} (Plate: {plate_number}, No description)")
            return local_zips

        # 1. Check top-level folder (level 0)
        plate_zips.extend(_scan_single_folder(folder_path_obj))

        # 2. Check immediate subfolders (level 1) - where plate zips typically are
        for item in self.get_directory_contents(folder_path):
            if item.is_dir():
                subfolder_path = folder_path_obj / item.name
                plate_zips.extend(_scan_single_folder(subfolder_path))

        # 3. Check zip dump folder specifically if it exists
        zip_dump_folder = folder_path_obj / self.config.ZIP_DUMP_FOLDER
        if zip_dump_folder.exists():
            plate_zips.extend(_scan_single_folder(zip_dump_folder))

        # Remove duplicates (in case same file found in multiple locations)
        seen_files = set()
        unique_plate_zips = []
        for zip_info in plate_zips:
            zip_path = zip_info[0]
            file_name = Path(zip_path).name
            if file_name not in seen_files:
                seen_files.add(file_name)
                unique_plate_zips.append(zip_info)
            else:
                self.log(f"Skipping duplicate plate folder zip: {file_name}")

        return unique_plate_zips


    def get_pcr_number(self, filename):
        #Move to path_utilities.py-done
        """Extract PCR number from file name"""
        pcr_num = ''
        # Look for PCR pattern in brackets - only this matters for folder sorting
        if re.search('{pcr\\d+.+}', filename.lower()):
            pcr_bracket = re.search("{pcr\\d+.+}", filename.lower()).group() #type: ignore
            pcr_num = re.search("pcr(\\d+)", pcr_bracket).group(1) #type: ignore
        return pcr_num


    def is_control_file(self, file_name, control_list):
        """Check if file is a control sample"""
        clean_name = self.clean_braces_format(file_name)
        clean_name = self.remove_extension(clean_name)
        clean_name_lower = clean_name.lower()

        # Check each control pattern
        for control in control_list:
            control_lower = control.lower()

            # Direct match (for individual sequencing controls)
            if clean_name_lower == control_lower:
                return True

            # For plate controls: check if filename matches pattern [2digits][letter]_[control_name]
            # Example: "12a_pgem_m13f-20" should match "pgem_m13f-20"
            well_match = self.regex_patterns['plate_well_prefix'].match(clean_name_lower)
            if well_match:
                # Extract the part after the well position
                name_after_well = well_match.group(1)
                if name_after_well == control_lower:
                    return True

        return False

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
            path_obj = Path(path)
            sorted_folders = sorted(
                [f for f in folders if (path_obj / f).is_dir()],
                key=lambda f: (path_obj / f).stat().st_mtime,
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

        path = Path(path)  # Ensure path is a Path object
        matching_folders = []

        # Scan the directory for matching folders
        for item in self.get_directory_contents(path):
            item_path = path / item.name
            if not item_path.is_dir():
                continue

            # Check if folder matches pattern
            if self.regex_patterns[folder_pattern].search(item.name):
                # Check exclusion patterns
                if not any(exclude in item.name.lower() for exclude in exclude_patterns):
                    matching_folders.append(item_path)

        # Extract unique I-numbers
        i_numbers = []
        for folder in matching_folders:
            i_num = self.get_inumber_from_name(folder.name)
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

        path_obj = Path(file_path)
        dir_name = path_obj.parent
        base_name = path_obj.name
        new_name = re.sub(r'{.*?}', '', base_name)

        new_path = dir_name / new_name

        try:
            if path_obj.exists():
                path_obj.rename(new_path)
                return str(new_path)
        except Exception as e:
            print(f"Error renaming file {file_path}: {e}")

        return file_path


    def standardize_filename_for_matching(self, file_name, remove_extension=True, preserve_order_number=True):
        """
        Standardized filename cleaning method for consistent matching across all code
        Used for PCR files, reinject lists, validation, and any other string comparison operations

        Args:
            file_name (str): The filename to standardize
            remove_extension (bool): Whether to remove file extension
            preserve_order_number (bool): Whether to preserve order number in result
        """
        # Step 1: Extract order number if present before any cleaning
        order_number = None
        if preserve_order_number:
            order_number = self.extract_order_number_from_filename(file_name)

        # Step 2: Remove file extension if needed
        if remove_extension and file_name.endswith(self.config.ABI_EXTENSION):
            clean_name = file_name[:-4]
        else:
            clean_name = file_name

        # Step 3: Check if the name is just a well location (e.g., {01G})
        well_pattern = re.compile(r'^{\d+[A-Z]}$')  # Added $ to match ONLY well locations
        if well_pattern.match(clean_name):
            # For well locations, keep them as is to avoid empty string normalization
            return clean_name

        # Step 4: Apply ABI character adjustments - ENSURE SPACES ARE REMOVED
        clean_name = self.adjust_abi_chars(clean_name)

        # Step 4.5: EXPLICITLY remove spaces if adjust_abi_chars didn't do it
        clean_name = clean_name.replace(' ', '')

        # Step 5: Remove content in brackets
        clean_name = re.sub(r'{.*?}', '', clean_name)

        # Step 6: Remove standard suffixes
        clean_name = self.neutralize_suffixes(clean_name)

        # Step 7: If we have an order number and want to preserve it, add it back
        if preserve_order_number and order_number:
            clean_name = f"{clean_name}#{order_number}"  # Use # as a delimiter that won't be in normal filenames

        # Step 8: If after all cleaning we have an empty string, return the original
        # to avoid normalization conflicts
        if not clean_name.strip():
            return file_name

        return clean_name

    def standardize_for_customer_files(self, file_name, remove_extension=True):
        """
        Specialized method for customer file matching in order key
        Removes well locations and normalizes for order key lookup

        Args:
            file_name (str): The filename to standardize
            remove_extension (bool): Whether to remove file extension

        Returns:
            str: Cleaned filename for order key matching
        """
        # Step 1: Remove file extension if needed
        if remove_extension and file_name.endswith(self.config.ABI_EXTENSION):
            clean_name = file_name[:-4]
        else:
            clean_name = file_name

        # Step 2: Apply ABI character adjustments and remove spaces
        clean_name = self.adjust_abi_chars(clean_name)
        clean_name = clean_name.replace(' ', '')

        # Step 3: Remove ALL content in brackets (including well locations)
        clean_name = re.sub(r'{.*?}', '', clean_name)

        # Step 4: Remove standard suffixes
        clean_name = self.neutralize_suffixes(clean_name)

        # Step 5: If after all cleaning we have an empty string, return the original
        if not clean_name.strip():
            return file_name

        return clean_name

    # Remove this method - we'll use standardize_for_customer_files directly

    def standardize_for_reinject_matching(self, file_name, remove_extension=True):
        """
        Specialized method for reinject list matching
        Preserves well locations for complex well-based matching

        Args:
            file_name (str): The filename to standardize
            remove_extension (bool): Whether to remove file extension

        Returns:
            str: Cleaned filename for reinject matching
        """
        # Step 1: Remove file extension if needed
        if remove_extension and file_name.endswith(self.config.ABI_EXTENSION):
            clean_name = file_name[:-4]
        else:
            clean_name = file_name

        # Step 2: Check if the name is just a well location (e.g., {01G})
        well_pattern = re.compile(r'^{\d+[A-Z]}$')
        if well_pattern.match(clean_name):
            # For well locations, keep them as is
            return clean_name

        # Step 3: Apply ABI character adjustments and remove spaces
        clean_name = self.adjust_abi_chars(clean_name)
        clean_name = clean_name.replace(' ', '')

        # Step 4: For reinject matching, we may want to preserve some bracket content
        # Remove only outer brackets but preserve structure for well matching
        # This is where reinject-specific logic would go

        # Step 5: Remove standard suffixes
        clean_name = self.neutralize_suffixes(clean_name)

        # Step 6: If after all cleaning we have an empty string, return the original
        if not clean_name.strip():
            return file_name

        return clean_name


if __name__ == "__main__":
    # Simple test if run directly
    from mseqauto.config import MseqConfig  # type: ignore

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
