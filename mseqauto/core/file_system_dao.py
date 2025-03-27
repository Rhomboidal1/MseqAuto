# file_system_dao.py
import os
import re
from datetime import datetime, timedelta
from shutil import move, copyfile
from zipfile import ZipFile, ZIP_DEFLATED
from mseqauto.config import MseqConfig

config = MseqConfig()

class FileSystemDAO:
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

    def get_directory_contents(self, path, refresh=False):
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

    def get_folders(self, path, pattern=None):
        """Get folders matching an optional regex pattern"""
        folder_list = []
        for item in self.get_directory_contents(path):
            full_path = os.path.join(path, item)
            if os.path.isdir(full_path):
                if pattern is None or re.search(pattern, item.lower()):
                    folder_list.append(full_path)
        return folder_list

    def get_files_by_extension(self, folder, extension):
        """Get all files with specified extension in a folder"""
        files = []
        for item in self.get_directory_contents(folder):
            if item.endswith(extension):
                files.append(os.path.join(folder, item))
        return files

    def contains_file_type(self, folder, extension):
        """Check if folder contains files with specified extension"""
        for item in self.get_directory_contents(folder):
            if item.endswith(extension):
                return True
        return False

    def create_folder_if_not_exists(self, path):
        """Create folder if it doesn't exist"""
        if not os.path.exists(path):
            os.mkdir(path)
        return path

    def move_folder(self, source, destination):
        """Move folder with proper error handling"""
        try:
            move(source, destination)
            return True
        except Exception as e:
            # Proper logging would be implemented here
            print(f"Error moving folder {source}: {e}")
            return False

    def get_folder_name(self, path):
        """Get the folder name from a path"""
        return os.path.basename(path)

    def get_parent_folder(self, path):
        """Get the parent folder path"""
        return os.path.dirname(path)

    def join_paths(self, base_path, *args):
        """Join path components"""
        return os.path.join(base_path, *args)

    def load_order_key(self, key_file_path):
        """Load the order key file"""
        try:
            import numpy as np
            return np.loadtxt(key_file_path, dtype=str, delimiter='\t')
        except Exception as e:
            print(f"Error loading order key file: {e}")
            return None

    def file_exists(self, path):
        """Check if a file exists"""
        return os.path.isfile(path)

    def folder_exists(self, path):
        """Check if a folder exists"""
        return os.path.isdir(path)

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
        """Remove anything contained in {} from filename"""
        return re.sub(r'{.*?}', '', self.neutralize_suffixes(file_name))

    def adjust_abi_chars(self, file_name):
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
        """Remove suffixes like _Premixed and _RTI"""
        new_file_name = file_name
        new_file_name = new_file_name.replace('_Premixed', '')
        new_file_name = new_file_name.replace('_RTI', '')
        return new_file_name

    def remove_extension(self, file_name, extension=None):
        """Remove file extension"""
        if extension and file_name.endswith(extension):
            return file_name[:-len(extension)]
        return os.path.splitext(file_name)[0]

    # Zip operations
    def check_for_zip(self, folder_path):
        """Check if folder contains any zip files"""
        for item in self.get_directory_contents(folder_path):
            file_path = os.path.join(folder_path, item)
            if os.path.isfile(file_path) and file_path.endswith(self.config.ZIP_EXTENSION):
                return True
        return False

    def zip_files(self, source_folder: str, zip_path: str, file_extensions=None, exclude_extensions=None):
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
        """Get list of files in a zip archive"""
        try:
            with ZipFile(zip_path, 'r') as zip_ref:
                return zip_ref.namelist()
        except Exception as e:
            print(f"Error reading zip file {zip_path}: {e}")
            return []

    def copy_zip_to_dump(self, zip_path, dump_folder):
        """Copy zip file to dump folder"""
        if not os.path.exists(dump_folder):
            os.makedirs(dump_folder)

        dest_path = os.path.join(dump_folder, os.path.basename(zip_path))
        copyfile(zip_path, dest_path)
        return dest_path

    # PCR folder operations
    def get_pcr_number(self, filename):
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

    def get_most_recent_inumber(self, path):
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
            return None

    def get_recent_files(self, paths, days=None, hours=None):
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

    def get_inumber_from_name(self, name):
        """Extract I number from a name using precompiled regex"""
        match = self.regex_patterns['inumber'].search(str(name).lower())
        if match:
            return match.group(1)  # Return just the number
        return None

    def get_inumbers_greater_than(self, files, lower_inum):
        """Get files with I number greater than specified value"""
        if not lower_inum:
            return files

        try:
            lower_inum_int = int(lower_inum)
            result = []

            for file_name in files:
                inum = self.get_inumber_from_name(file_name)
                if inum and int(inum) > lower_inum_int:
                    result.append(file_name)

            return result
        except ValueError:
            # If conversion to int fails, return the original list
            return files

    def move_file(self, source, destination):
        """Move a file with error handling"""
        try:
            move(source, destination)
            return True
        except Exception as e:
            print(f"Error moving file {source}: {e}")
            return False

    def rename_file_without_braces(self, file_path):
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
    ab1_files = dao.get_files_by_extension(test_path, ".py")
    print(f"Found {len(ab1_files)} Python files in {test_path}")
