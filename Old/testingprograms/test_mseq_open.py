# test_mseq_open.py
import os
import sys
from config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import time

def main():
    print("Starting mSeq open test...")
    
    # Initialize components
    config = MseqConfig()
    print("Config loaded")
    file_dao = FileSystemDAO(config)
    print("FileSystemDAO initialized")
    ui_automation = MseqAutomation(config)
    print("UI Automation initialized")
    
    # Hard-coded test folder (replace with a path that exists on your system)
    test_folder = r"C:\Users\rhomb\Documents"
    print(f"Using test folder: {test_folder}")
    
    # Directly try to open mSeq and navigate to the folder
    print("Attempting to open mSeq and navigate...")
    
    # Connect to mSeq or start it
    print("Connecting to or starting mSeq...")
    app, main_window = ui_automation.connect_or_start_mseq()
    print(f"mSeq window: {main_window.window_text() if main_window else 'None'}")
    
    # Set focus to main window
    if main_window:
        main_window.set_focus()
        print("Set focus to main window")
    else:
        print("ERROR: No main window found")
        return
    
    # Try to open new project
    print("Pressing Ctrl+N...")
    from pywinauto.keyboard import send_keys
    send_keys('^n')
    
    # Wait for Browse dialog
    print("Waiting for Browse For Folder dialog...")
    try:
        from pywinauto import timings
        timings.wait_until(timeout=10, retry_interval=0.5, 
                         func=lambda: app.window(title='Browse For Folder').exists(), 
                         value=True)
        
        browse_dialog = app.window(title='Browse For Folder')
        print("Browse For Folder dialog found")
        
        # Navigate to test folder
        print(f"Attempting to navigate to: {test_folder}")
        success = ui_automation.navigate_folder_tree(browse_dialog, test_folder)
        
        if success:
            print("Navigation successful, clicking OK")
            ok_button = app.BrowseForFolder.child_window(title="OK", class_name="Button")
            ok_button.click_input()
        else:
            print("Navigation failed")
    
    except timings.TimeoutError:
        print("ERROR: Browse For Folder dialog did not appear")
    
    # Keep script running to see what happens
    print("Test completed. Waiting 5 seconds before closing...")
    time.sleep(5)
    
    # Close mSeq
    print("Closing mSeq...")
    ui_automation.close()
    print("Done")

if __name__ == "__main__":
    main()