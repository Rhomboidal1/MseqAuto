# simple_ind_auto_mseq.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
from logger import setup_logger

# Simplified function
def get_folder_from_user():
    """Get folder selection from user"""
    print("Opening folder selection dialog...")
    
    root = tk.Tk()
    root.withdraw()
    
    folder_path = filedialog.askdirectory(
        title="Select today's data folder to mseq orders",
        mustexist=True
    )
    
    root.destroy()
    
    if folder_path:
        print(f"Selected folder: {folder_path}")
        return folder_path
    else:
        print("No folder selected")
        return None

def main():
    # Setup logger
    logger = setup_logger("simple_ind_auto_mseq")
    logger.info("Starting simple IND auto mSeq...")
    
    # Get folder first
    logger.info("About to select folder...")
    data_folder = get_folder_from_user()
    
    if not data_folder:
        logger.error("No folder selected, exiting")
        print("No folder selected, exiting")
        return
    
    logger.info(f"Selected folder: {data_folder}")
    print("Script would now continue processing...")

if __name__ == "__main__":
    main()