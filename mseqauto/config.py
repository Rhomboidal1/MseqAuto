import getpass
import platform
import re
import sys
from pathlib import Path, WindowsPath


class MseqConfig:
    # Get current username for H drive mapping
    USERNAME = getpass.getuser()

    # Paths using pathlib
    # Check for Python 32-bit, fallback to current Python interpreter
    _python32_candidate = Path("C:/Program Files (x86)/Python312-32/python.exe")
    if _python32_candidate.exists():
        PYTHON32_PATH = _python32_candidate
    else:
        # If no 32-bit Python is found, use the current interpreter
        PYTHON32_PATH = Path(sys.executable)
    
    VENV32_PATH = PYTHON32_PATH.parent / "venv32"
    MSEQ_PATH = Path("C:/DNA/Mseq4/bin")
    MSEQ_EXECUTABLE = "j.exe -jprofile mseq.ijl"


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
        'order_number_extract': re.compile(r'_(\d+)$'),  # Extract order number from end of folder name
        'raw_data_zip': re.compile(r'rd\d+\.zip$', re.IGNORECASE),  # Raw data zip files
        'plate_folder': re.compile(r'^P\d+_.+$', re.IGNORECASE),  # For plate folders
        'ind_blank_file': re.compile(r'{\d+[A-H]}.ab1$', re.IGNORECASE),  # Individual blanks pattern
        'plate_blank_file': re.compile(r'^\d{2}[A-H]__.ab1$', re.IGNORECASE),  # Plate blanks pattern
        'order_folder': re.compile(r'bioi-\d+_.+_\d+', re.IGNORECASE),  # For BioI order folders
        'pcr_folder': re.compile(r'fb-pcr\d+(_\d+)?', re.IGNORECASE),     # For PCR folders
        'reinject': re.compile(r'reinject', re.IGNORECASE),             # For filtering out reinject folders
        'double_well': re.compile(r'^{\d+[A-Z]}{(\d+[A-Z])}')
    }

    # Timeouts for UI operations - extended for Windows 11
    TIMEOUTS = {
        "browse_dialog": 10,
        "preferences": 8,
        "copy_files": 8,
        "error_window": 20,
        "call_bases": 15,
        "process_completion": 15,
        "read_info": 8,
        "post_read_dialog": 15  # seconds to wait for files after read dialog appears
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

    # Paths for data operations using pathlib
    KEY_FILE_PATH = Path("P:/order_key.txt")
    BATCH_FILE_PATH = Path("P:/generate-data-sorting-key-file.bat")
    REINJECT_FOLDER = Path("P:/Data/Reinjects")

    # Convenience methods for path operations
    @classmethod
    def get_python32_path_str(cls) -> str:
        """Get the Python 32-bit path as a string for subprocess calls"""
        return str(cls.PYTHON32_PATH)
    
    @classmethod
    def get_mseq_path_str(cls) -> str:
        """Get the mSeq path as a string for subprocess calls"""
        return str(cls.MSEQ_PATH)
    
    @classmethod
    def get_key_file_path_str(cls) -> str:
        """Get the order key file path as a string"""
        return str(cls.KEY_FILE_PATH)
    
    @classmethod
    def get_batch_file_path_str(cls) -> str:
        """Get the batch file path as a string"""
        return str(cls.BATCH_FILE_PATH)
    
    @classmethod
    def get_reinject_folder_str(cls) -> str:
        """Get the reinject folder path as a string"""
        return str(cls.REINJECT_FOLDER)

    # Windows version detection
    @staticmethod
    def is_windows_11():
        """Detect if running on Windows 11"""
        if platform.system() != 'Windows':
            return False

        win_version = int(platform.version().split('.')[0])
        win_build = int(platform.version().split('.')[2]) if len(platform.version().split('.')) > 2 else 0
        return win_version >= 10 and win_build >= 22000