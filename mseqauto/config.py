# config.py
import getpass
import os
import platform
import re


class MseqConfig:
    # Get current username for H drive mapping
    USERNAME = getpass.getuser()

    # Paths
    # Check for Python 32-bit, fallback to current Python interpreter
    if os.path.exists(r"C:\Program Files (x86)\Python312-32\python.exe"):
        PYTHON32_PATH = r"C:\Program Files (x86)\Python312-32\python.exe"
    else:
        # If no 32-bit Python is found, use the current interpreter
        import sys
        PYTHON32_PATH = sys.executable
    VENV32_PATH = os.path.join(os.path.dirname(PYTHON32_PATH), "venv32")
    MSEQ_PATH = r"C:\DNA\Mseq4\bin"
    MSEQ_EXECUTABLE = r"j.exe -jprofile mseq.ijl"

    # File Extensions
    ABI_EXTENSION = '.ab1'
    ZIP_EXTENSION = '.zip'
    FSA_EXTENSION = '.fsa'

    # Network drives - these are mapped differently in file dialogs
    NETWORK_DRIVES = {
        "P:": r"ABISync (P:)",
        "H:": f"{USERNAME} (\\\\w2k16\\users) (H:)"  # Dynamic user mapping
    }

    # Precompiled regex patterns
    REGEX_PATTERNS = {
        'inumber': re.compile(r'bioi-(\d+)', re.IGNORECASE),
        'pcr_number': re.compile(r'{pcr(\d+).+}', re.IGNORECASE),
        'brace_content': re.compile(r'{.*?}'),
        'bioi_folder': re.compile(r'^bioi-\d+$', re.IGNORECASE),
        'ind_blank_file': re.compile(r'{\d+[A-H]}.ab1$', re.IGNORECASE),  # Individual blanks pattern
        'plate_blank_file': re.compile(r'^\d{2}[A-H]__.ab1$', re.IGNORECASE),  # Plate blanks pattern
        'order_folder': re.compile(r'bioi-\d+_.+_\d+', re.IGNORECASE),  # For BioI order folders
        'pcr_folder': re.compile(r'fb-pcr\d+(_\d+)?', re.IGNORECASE),     # For PCR folders
        'reinject': re.compile(r'reinject', re.IGNORECASE),             # For filtering out reinject folders
        'double_well': re.compile(r'^{\d+[A-Z]}{(\d+[A-Z])}')
    }

    # Timeouts for UI operations - extended for Windows 11
    TIMEOUTS = {
        "browse_dialog": 10,  # Reduced from 10 3
        "preferences": 8,  # Reduced from 8 2
        "copy_files": 8,  # Reduced from 8 2
        "error_window": 20,  # Reduced from 20 3
        "call_bases": 15,  # Reduced from 15 3
        "process_completion": 15,  # Reduced from 60 15
        "read_info": 8  # Reduced from 8 2
    }

    CONTROLS = [
        '_pGEM_T7Promoter',
        '_pGEM_SP6Promoter',
        '_pGEM_M13F-20',
        '_pGEM_M13R-27',
        '_Water_T7Promoter',
        '_Water_SP6Promoter',
        '_Water_M13F-20',
        '_Water_M13R-27'
    ]

    PLATE_CONTROLS = [
        'pgem_m13f-20',
        'water_m13f-20'
    ]

    # Special folder names
    CONTROLS_FOLDER = "Controls"
    BLANK_FOLDER = "Blank"
    ALT_INJECTIONS_FOLDER = "Alternate Injections"
    ZIP_DUMP_FOLDER = "zip dump"
    IND_NOT_READY_FOLDER = "IND Not Ready"

    # Mseq artifacts
    MSEQ_ARTIFACTS = {'chromat_dir', 'edit_dir', 'phd_dir', 'mseq4.ini'}

    # File types
    TEXT_FILES = [
        '.raw.qual.txt',
        '.raw.seq.txt',
        '.seq.info.txt',
        '.seq.qual.txt',
        '.seq.txt'
    ]

    # Excel validation styling
    EXCEL_STYLES = {
        "success": "00CC00",  # Green
        "attention": "FF4747",  # Red
        "resolved": "FFA500",  # Orange
        "break": "DDD9C4"  # Light gray
    }

    # Special case handling
    ANDREEV_NAME = "andreev"  # Used to identify special handling for Andreev orders

    # Paths for data operations
    #KEY_FILE_PATH = r"P:\order_key.txt"
    KEY_FILE_PATH = r"P:\order_key.txt"
    BATCH_FILE_PATH = r"P:\generate-data-sorting-key-file.bat"
    REINJECT_FOLDER = r"P:\Data\Reinjects"

    # Windows version detection
    @staticmethod
    def is_windows_11():
        """Detect if running on Windows 11"""
        if platform.system() != 'Windows':
            return False

        win_version = int(platform.version().split('.')[0])
        win_build = int(platform.version().split('.')[2]) if len(platform.version().split('.')) > 2 else 0
        return win_version >= 10 and win_build >= 22000
