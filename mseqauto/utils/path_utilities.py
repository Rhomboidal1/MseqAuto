import re
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

print(sys.path)

from mseqauto.config import MseqConfig
from mseqauto.utils import setup_logger
from mseqauto.config import MseqConfig
from mseqauto.core import OSCompatibilityManager, FileSystemDAO, MseqAutomation, FolderProcessor
config = MseqConfig()
file_dao = FileSystemDAO(config)
processor = FolderProcessor(file_dao, None, config)

#remove braces from string
#adjust abi chars
#neutralize suffixes
# standardize filename for matching
# #normalize filename
# get inumber from name
# get pcr number
# rename getfoldername to get_basename
# rename getparentfoledr to get_dirname
# join_paths
# is_bioi_folder
# is_pcr_folder
# is_plate_folder
# is_reinject_folder
# standardize path from FolderProcessor


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

def remove_braces_from_string(file_name): #clean_braces_format
     """Remove anything contained in {} from filename"""
     return re.sub(r'{.*?}', '', neutralize_suffixes(file_name))

def get_pcr_number(filename):
     """Extract PCR number from file name"""
     pcr_num = ''
     # Look for PCR pattern in brackets - only this matters for folder sorting
     if re.search('{pcr\\d+.+}', filename.lower()):
          pcr_bracket = re.search("{pcr\\d+.+}", filename.lower()).group()
          pcr_num = re.search("pcr(\\d+)", pcr_bracket).group(1)
     return pcr_num

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
     #Move to path_utilities.py
     """Extract I number from a name using precompiled regex"""

     match = file_dao.regex_patterns['inumber'].search(str(name).lower())
     if match:
          return match.group(1)  # Return just the number
     return None