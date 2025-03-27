# path_utilities.py
import os
import re

def clean_braces_format(file_name):
    """Remove anything contained in {} from filename"""
    return re.sub(r'{.*?}', '', neutralize_suffixes(file_name))
    
def remove_braces_from_string(text):
    """
    Remove braces and their contents from a string.
    
    Args:
        text (str): Text to process
    
    Returns:
        str: Text with braces and their contents removed
    """
    return re.sub(r'\{[^}]*\}', '', text)

def adjust_abi_chars(file_name):
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

def neutralize_suffixes(file_name):
    """Remove suffixes like _Premixed and _RTI"""
    new_file_name = file_name
    new_file_name = new_file_name.replace('_Premixed', '')
    new_file_name = new_file_name.replace('_RTI', '')
    return new_file_name

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

def normalize_filename(file_name, remove_extension=True, logger=None):
    """Normalize filename with optional logging"""
    # Step 1: Adjust characters
    adjusted_name = adjust_abi_chars(file_name)

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
    neutralized_name = neutralize_suffixes(adjusted_name)

    # Step 4: Remove brace content
    cleaned_name = re.sub(r'{.*?}', '', neutralized_name)

    # Only log if a logger is provided
    if logger:
        logger(f"Normalized '{file_name}' to '{cleaned_name}'")

    return cleaned_name

def get_inumber_from_name(name):
    """Extract I number from a name (e.g., 'BioI-12345' -> '12345')"""
    return extract_number_from_pattern(name, r'bioi-(\d+)', 1)

def get_pcr_number(filename):
    """Extract PCR number from a filename (e.g., '{PCR1234}' -> '1234')"""
    if re.search('{pcr\\d+.+}', filename.lower()):
        pcr_bracket = re.search("{pcr\\d+.+}", filename.lower()).group()
        return re.search("pcr(\\d+)", pcr_bracket).group(1)
    return None

def get_order_number_from_folder_name(folder_name):
    """Extract order number from a folder name (e.g., 'BioI-12345_Name_67890' -> '67890')"""
    match = re.search(r'_(\d+)$', folder_name)
    if match:
        return match.group(1)
    return None
######### Path Utility Functions #########
def get_folder_name(self, path):
    """Get the folder name from a path"""
    return os.path.basename(path)

def get_parent_folder(self, path):
    """Get the parent folder path"""
    return os.path.dirname(path)

def join_paths(self, base_path, *args):
    """Join path components"""
    return os.path.join(base_path, *args)


def extract_number_from_pattern(text, pattern, group_index=0):
    """
    Extract a number from text using a regex pattern.
    
    Args:
        text (str): Text to search in
        pattern (str): Regex pattern to match
        group_index (int): Group index to extract (defaults to 0 - whole match)
    
    Returns:
        str: Extracted number or None if not found
    """
    match = re.search(pattern, str(text).lower())
    if match:
        return match.group(group_index)
    return None
######### Pattern Matching Utilities #########
def is_bioi_folder(folder_name):
    """Check if folder name matches BioI pattern"""
    return bool(re.search(r'bioi-\d+', folder_name.lower()))

def is_pcr_folder(folder_name):
    """Check if folder name matches PCR pattern"""
    return bool(re.search(r'fb-pcr\d+_\d+', folder_name.lower()))

def is_plate_folder(folder_name):
    """Check if folder name matches plate pattern"""
    return bool(re.search(r'^p\d+', folder_name.lower()))

def is_order_folder(folder_name):
    """Check if folder name matches order pattern"""
    return bool(re.search(r'bioi-\d+_.+_\d+', folder_name.lower()))

def standardize_path(path):
    """
    Standardize paths to use the correct directory separator for the OS.
    
    Args:
        path (str): Path to standardize
    
    Returns:
        str: Standardized path with correct separators
    """
    if path is None:
        return None
        
    # Convert all paths to use os.path.sep
    standardized_path = path.replace('\\', os.path.sep).replace('/', os.path.sep)
    
    # If path is a network path and starts with a single backslash, add another
    # This handles UNC paths on Windows
    if os.name == 'nt' and standardized_path.startswith(os.path.sep) and not standardized_path.startswith(os.path.sep + os.path.sep):
        standardized_path = os.path.sep + standardized_path
        
    return standardized_path
