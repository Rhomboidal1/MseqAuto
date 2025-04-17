# ab1_processor.py
import os
import subprocess
import shutil
import logging
import re
import tempfile
from pathlib import Path
import datetime
from zipfile import ZipFile, ZIP_DEFLATED

class AB1Processor:
    """Processes AB1 files to generate outputs compatible with the existing workflow"""
    
    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        
        # Set up phred/phrap paths
        self.phred_path = self._find_executable('phred', [
            r"C:\DNA\Mseq4\bin\phred.exe",
            r"phred.exe"
        ])
        
        self.cross_match_path = self._find_executable('cross_match', [
            r"C:\DNA\Mseq4\bin\cross_match.exe",
            r"cross_match.exe"
        ])
        
        # Define standard quality threshold
        self.quality_threshold = 20
        
    def _find_executable(self, name, paths):
        """Find an executable in the given paths"""
        for path in paths:
            if os.path.exists(path):
                self.logger.info(f"Found {name} at {path}")
                return path
        
        self.logger.warning(f"Could not find {name} in standard paths")
        return None
    
    def process_folder(self, folder_path):
        """Process all AB1 files in a folder"""
        self.logger.info(f"Processing folder: {folder_path}")
        
        # Extract project name
        project_name = os.path.basename(os.path.normpath(folder_path))
        
        # Create directories similar to mSeq
        chromat_dir = os.path.join(folder_path, "chromat_dir")
        phd_dir = os.path.join(folder_path, "phd_dir")
        edit_dir = os.path.join(folder_path, "edit_dir")
        
        for directory in [chromat_dir, phd_dir, edit_dir]:
            os.makedirs(directory, exist_ok=True)
        
        # Find AB1 files
        ab1_files = []
        for item in os.listdir(folder_path):
            if item.endswith(self.config.ABI_EXTENSION):
                ab1_files.append(os.path.join(folder_path, item))
        
        if not ab1_files:
            self.logger.warning(f"No AB1 files found in {folder_path}")
            return False
        
        # Process the files
        return self._process_ab1_files(folder_path, ab1_files, project_name)
    
    def _process_ab1_files(self, folder_path, ab1_files, project_name):
        """Core processing logic for AB1 files"""
        # Implementation based on process_ab1_files from phred_mseq_replacement.py
        # [Implementation details here]
        
        # Return success/failure
        return True
        
    # [Additional methods from phred_mseq_replacement.py]
    
    def check_process_status(self, folder_path):
        """
        Check if a folder has been processed
        
        Returns:
            tuple: (was_processed, has_braces, has_ab1_files)
        """
        # Check for required output files instead of mSeq directories
        was_processed = True
        has_braces = False
        has_ab1_files = False
        
        # Check for the 5 text files that indicate processing is complete
        txt_file_count = 0
        project_name = os.path.basename(folder_path)
        
        for txt_ext in self.config.TEXT_FILES:
            if os.path.exists(os.path.join(folder_path, f"{project_name}{txt_ext}")):
                txt_file_count += 1
        
        # If we don't have all 5 text files, it wasn't processed
        if txt_file_count < 5:
            was_processed = False
        
        # Check for AB1 files and braces
        for item in os.listdir(folder_path):
            if item.endswith(self.config.ABI_EXTENSION):
                has_ab1_files = True
                if '{' in item or '}' in item:
                    has_braces = True
        
        return was_processed, has_braces, has_ab1_files
    
    def zip_processed_files(self, folder_path):
        """
        Create a zip file of the processed files
        
        Returns:
            str: Path to the created zip file
        """
        # Get base name for zip file
        folder_name = os.path.basename(folder_path)
        zip_path = os.path.join(folder_path, f"{folder_name}.zip")
        
        # Determine which files to include
        files_to_zip = []
        project_name = folder_name
        
        # Add the 5 text files
        for txt_ext in self.config.TEXT_FILES:
            txt_file = os.path.join(folder_path, f"{project_name}{txt_ext}")
            if os.path.exists(txt_file):
                files_to_zip.append(txt_file)
        
        # Add AB1 files
        for item in os.listdir(folder_path):
            if item.endswith(self.config.ABI_EXTENSION):
                files_to_zip.append(os.path.join(folder_path, item))
        
        # Create zip file
        with ZipFile(zip_path, 'w', ZIP_DEFLATED) as zipf:
            for file_path in files_to_zip:
                zipf.write(file_path, arcname=os.path.basename(file_path))
        
        self.logger.info(f"Created zip file: {zip_path}")
        return zip_path