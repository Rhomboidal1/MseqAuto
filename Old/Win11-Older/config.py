# config.py
import getpass

class MseqConfig:
    # Get current username for H drive mapping
    USERNAME = getpass.getuser()
    
    # Paths
    PYTHON32_PATH = r"C:\Python312-32\python.exe"
    MSEQ_PATH = r"C:\DNA\Mseq4\bin"
    MSEQ_EXECUTABLE = r"j.exe -jprofile mseq.ijl"
    
    # Network drives - these are mapped differently in file dialogs
    NETWORK_DRIVES = {
        "P:": r"ABISync (P:)",
        "H:": f"{USERNAME} (\\\\w2k16\\users) (H:)"  # Dynamic user mapping
    }
    
    # Timeouts for UI operations
    TIMEOUTS = {
        "browse_dialog": 5,
        "preferences": 5, 
        "copy_files": 5,
        "error_window": 20,
        "call_bases": 10,
        "process_completion": 45,
        "read_info": 5
    }
    
    # Special folders
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
    
    # Paths for data operations
    KEY_FILE_PATH = r"P:\order_key.txt"
    BATCH_FILE_PATH = r"P:\generate-data-sorting-key-file.bat"
    REINJECT_FOLDER = r"P:\Data\Reinjects"