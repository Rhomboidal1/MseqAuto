# plate_auto_mseq.py
import os
import sys
import tkinter.filedialog as filedialog
import subprocess
from config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import re

# Check for 32-bit Python requirement
if sys.maxsize > 2**32:
    # Path to 32-bit Python
    py32_path = MseqConfig.PYTHON32_PATH
    
    if os.path.exists(py32_path):
        # Get the full path of the current script
        script_path = os.path.abspath(__file__)
        
        # Re-run this script with 32-bit Python and exit current process
        subprocess.run([py32_path, script_path])
        sys.exit(0)
    else:
        print("32-bit Python not found at", py32_path)
        print("Continuing with 64-bit Python (may cause issues)")

def main():
    config = MseqConfig()
    file_dao = FileSystemDAO(config)
    ui_automation = MseqAutomation(config)
    processor = FolderProcessor(file_dao, ui_automation, config)
    
    # Get folder selection from user
    data_folder = filedialog.askdirectory(title="Select today's data folder to mseq plates")
    data_folder = re.sub(r'/', '\\\\', data_folder)
    
    # Get all plate folders (starting with 'p')
    plate_folders = file_dao.get_folders(data_folder, r'p\d+.+')
    
    # Process each plate folder
    for folder in plate_folders:
        print(folder)
        processor.process_plate_folder(folder)
    
    # Close mSeq application
    ui_automation.close()

if __name__ == "__main__":
    main()