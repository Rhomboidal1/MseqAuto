# file_system_dao.py
import os
import re
from shutil import move

class FileSystemDAO:
    def __init__(self, config):
        self.config = config
    
    def get_folders(self, path, pattern=None):
        """Get folders matching an optional regex pattern"""
        folders = []
        for item in os.listdir(path):
            full_path = os.path.join(path, item)
            if os.path.isdir(full_path):
                if pattern is None or re.search(pattern, item.lower()):
                    folders.append(full_path)
        return folders
    
    def get_files_by_extension(self, folder, extension):
        """Get all files with specified extension in a folder"""
        files = []
        for item in os.listdir(folder):
            if item.endswith(extension):
                files.append(os.path.join(folder, item))
        return files
    
    def contains_file_type(self, folder, extension):
        """Check if folder contains files with specified extension"""
        for item in os.listdir(folder):
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
    
    def join_paths(self, *args):
        """Join path components"""
        return os.path.join(*args)
    
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
        for item in os.listdir(folder):
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