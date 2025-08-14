#!/usr/bin/env python3
"""
Simple Reinject List Inspector

This script extracts and displays the contents of both raw_reinject_list and reinject_list
to help debug what entries are being added.

Usage:
    # With command line I-numbers:
    python inspect_reinject_lists.py 21456 21457
    
    # Or run without arguments to browse for a data folder:
    python inspect_reinject_lists.py
    
    The script will automatically detect I-numbers from folder names in the selected directory.
"""

import re
import sys
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import numpy as np

# Add parent directories to PYTHONPATH for imports
sys.path.append(str(Path(__file__).parents[2]))

# DON'T import mseqauto modules here - do it after getting the folder!


def simple_print_log(message):
    """Simple logging function"""
    print(message)


def get_folder_from_user():
    """
    Get folder path from user via file dialog.
    
    Returns:
        Path: Selected folder path, or None if no folder selected
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    root.attributes('-topmost', True)  # Bring to front
    root.focus_force()  # Force focus
    
    try:
        folder_path = filedialog.askdirectory(
            title="Select today's data folder to inspect for reinjects",
            mustexist=True,
            parent=root  # Make sure dialog is attached to root
        )
        if folder_path:
            print(f"Selected folder: {folder_path}")
            return Path(folder_path)
        else:
            print("No folder selected.")
            return None
    except Exception as e:
        print(f"Error opening folder dialog: {e}")
        return None
    finally:
        root.destroy()


def get_recent_inumbers_from_folder(folder_path, file_dao):
    """
    Get I-numbers from folder names in the selected directory
    
    Args:
        folder_path: Path to search for I-numbers
        file_dao: FileSystemDAO instance1
        
    Returns:
        list: List of I-numbers found
    """
    i_numbers = []
    try:
        for item in folder_path.iterdir():
            if item.is_dir():
                # Extract I-number from folder name
                i_num = file_dao.get_inumber_from_name(item.name)
                if i_num and i_num not in i_numbers:
                    i_numbers.append(i_num)
    except Exception as e:
        print(f"Error scanning folder {folder_path}: {e}")
    
    return sorted(i_numbers, key=int) if i_numbers else []


class SimpleReinjectInspector:
    """Simple helper to get reinject lists"""
    
    def __init__(self, MseqConfig, FileSystemDAO):
        self.config = MseqConfig()
        self.file_dao = FileSystemDAO(self.config, simple_print_log)
        
    def is_valid_entry(self, raw_name):
        """Check if entry is just a well location - matches the original function"""
        if not raw_name or not isinstance(raw_name, str):
            return False
        # Skip entries that are just well locations in braces
        if re.match(r'^\{\d+[A-H]\}$', raw_name):
            return False
        # Skip empty entries
        if not raw_name.strip():
            return False
        return True
    
    def get_reinject_lists(self, i_numbers):
        """
        Extract reinject lists - simplified version
        
        Args:
            i_numbers: List of I-numbers to search for
            
        Returns:
            tuple: (reinject_list, raw_reinject_list)
        """
        print(f"\n=== REINJECT LIST EXTRACTION ===")
        print(f"Searching for I-numbers: {i_numbers}")
        
        reinject_list = []
        raw_reinject_list = []
        processed_files = set()

        # Process text files 
        spreadsheets_path = Path(r'G:\Lab\Spreadsheets')
        abi_path = spreadsheets_path / 'Individual Uploaded to ABI'
        reinject_files = []

        print(f"\nSearching for reinject files in:")
        print(f"  - {spreadsheets_path}")
        print(f"  - {abi_path}")

        # Build list of reinject files
        if spreadsheets_path.exists():
            for file_path in self.file_dao.get_directory_contents(spreadsheets_path):
                if 'reinject' in file_path.name.lower() and file_path.suffix == '.txt':
                    if any(i_num in file_path.name for i_num in i_numbers):
                        reinject_files.append(str(file_path))
                        print(f"  Found reinject file: {file_path.name}")

        if abi_path.exists():
            for file_path in self.file_dao.get_directory_contents(abi_path):
                if 'reinject' in file_path.name.lower() and file_path.suffix == '.txt':
                    if any(i_num in file_path.name for i_num in i_numbers):
                        reinject_files.append(str(file_path))
                        print(f"  Found reinject file: {file_path.name}")

        if not reinject_files:
            print("  No reinject .txt files found matching the I-numbers")

        # Process each found reinject file
        for file_path in reinject_files:
            try:
                if file_path in processed_files:
                    continue

                processed_files.add(file_path)
                print(f"\n--- Processing reinject file: {Path(file_path).name} ---")
                
                data = np.loadtxt(file_path, dtype=str, delimiter='\t', skiprows=5)
                print(f"File has {data.shape[0]} rows and {data.shape[1]} columns")

                # Parse rows
                for j in range(0, min(96, data.shape[0])):
                    if j < data.shape[0] and data.shape[1] > 1:
                        raw_name = data[j, 1]
                        
                        if self.is_valid_entry(raw_name):
                            cleaned_name = self.file_dao.standardize_filename_for_matching(raw_name)
                            raw_reinject_list.append(raw_name)
                            reinject_list.append(cleaned_name)
                            
                            print(f"  Row {j+1}: '{raw_name}' -> '{cleaned_name}'")
                        else:
                            print(f"  Row {j+1}: '{raw_name}' -> SKIPPED (invalid entry)")
                
            except Exception as e:
                print(f"ERROR processing reinject file {file_path}: {e}")

        return reinject_list, raw_reinject_list


def main():
    print("Reinject List Inspector")
    print("======================")
    
    # Check for command line arguments first
    if len(sys.argv) > 1:
        # Use command line I-numbers if provided
        i_numbers = sys.argv[1:]
        print(f"Using I-numbers from command line: {i_numbers}")
        
        # Import modules only after we know we need them
        try:
            from mseqauto.config import MseqConfig
            from mseqauto.core.file_system_dao import FileSystemDAO
        except ImportError as e:
            print(f"Import error: {e}")
            print("Make sure you're running this from the correct directory and all dependencies are installed.")
            return
        
        # Create inspector
        inspector = SimpleReinjectInspector(MseqConfig, FileSystemDAO)
        
    else:
        # No command line arguments - use folder browser FIRST, then import modules
        print("No I-numbers provided. Please select a data folder to auto-detect I-numbers.")
        print("Opening folder selection dialog...")
        
        folder_path = get_folder_from_user()
        if not folder_path:
            print("No folder selected. Exiting.")
            return
        
        print(f"Selected folder: {folder_path}")
        
        # ONLY NOW import package modules after folder selection
        try:
            from mseqauto.config import MseqConfig
            from mseqauto.core.file_system_dao import FileSystemDAO
        except ImportError as e:
            print(f"Import error: {e}")
            print("Make sure you're running this from the correct directory and all dependencies are installed.")
            return
        
        # Initialize inspector to get file_dao for I-number extraction
        inspector = SimpleReinjectInspector(MseqConfig, FileSystemDAO)
        
        # Get I-numbers from the selected folder
        i_numbers = get_recent_inumbers_from_folder(folder_path, inspector.file_dao)
        
        if not i_numbers:
            print(f"No I-numbers found in folder: {folder_path}")
            return
        
        print(f"Found I-numbers in folder: {i_numbers}")
        
        # Ask user which I-numbers to use
        print("\nWhich I-numbers would you like to inspect?")
        print("1. All found I-numbers")
        print("2. Select specific I-numbers")
        
        choice = input("Enter choice (1 or 2): ").strip()
        
        if choice == "2":
            print(f"Available I-numbers: {', '.join(i_numbers)}")
            selected = input("Enter I-numbers separated by spaces: ").strip().split()
            # Validate selection
            valid_selected = [i for i in selected if i in i_numbers]
            if valid_selected:
                i_numbers = valid_selected
                print(f"Using selected I-numbers: {i_numbers}")
            else:
                print("No valid I-numbers selected. Using all found I-numbers.")
    
    try:
        reinject_list, raw_reinject_list = inspector.get_reinject_lists(i_numbers)
        
        # Print the results
        print(f"\n" + "="*60)
        print(f"RESULTS")
        print(f"="*60)
        
        print(f"\nTotal entries found: {len(reinject_list)}")
        
        print(f"\n" + "-"*40)
        print(f"RAW REINJECT LIST ({len(raw_reinject_list)} entries):")
        print(f"-"*40)
        for i, entry in enumerate(raw_reinject_list, 1):
            print(f"{i:3d}. {entry}")
        
        print(f"\n" + "-"*40)
        print(f"CLEANED REINJECT LIST ({len(reinject_list)} entries):")
        print(f"-"*40)
        for i, entry in enumerate(reinject_list, 1):
            print(f"{i:3d}. {entry}")
        
        # Show mapping
        print(f"\n" + "-"*40)
        print(f"RAW -> CLEANED MAPPING:")
        print(f"-"*40)
        for i, (raw, cleaned) in enumerate(zip(raw_reinject_list, reinject_list), 1):
            if raw != cleaned:
                print(f"{i:3d}. '{raw}' -> '{cleaned}'")
            else:
                print(f"{i:3d}. '{raw}' (no change)")
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
