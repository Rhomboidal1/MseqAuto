# test_mseq_open.py
import os
import sys
import getpass
from pathlib import Path
import platform
from config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import time

def get_documents_folder():
    """Get the Documents folder path using multiple methods"""
    current_user = getpass.getuser()
    print(f"Current user: {current_user}")
    
    # Standard paths to try
    paths_to_try = [
        # Standard Windows path
        os.path.join("C:", "Users", current_user, "Documents"),
        # Alternative spellings
        os.path.join("C:", "Users", current_user, "Document"),
        os.path.join("C:", "Users", current_user, "My Documents"),
        # OneDrive path
        os.path.join("C:", "Users", current_user, "OneDrive", "Documents"),
        # Windows localized names
        os.path.join("C:", "Users", current_user, "Dokumenter"),  # Danish
        os.path.join("C:", "Users", current_user, "Dokumente"),   # German
        os.path.join("C:", "Users", current_user, "Documentos"),  # Spanish
        # Windows folder shorthand - only works on Windows
        os.path.expandvars("%USERPROFILE%\\Documents"),
        # D: drive possibility
        os.path.join("D:", "Users", current_user, "Documents"),
    ]
    
    # Try to get Documents folder through Windows shell - most reliable method
    try:
        import ctypes
        from ctypes import windll, wintypes
        from uuid import UUID
        
        # GUID for the Documents folder
        FOLDERID_Documents = UUID('{FDD39AD0-238F-46AF-ADB4-6C85480369C7}')
        
        # Define the function to retrieve known folder path
        _SHGetKnownFolderPath = windll.shell32.SHGetKnownFolderPath
        _SHGetKnownFolderPath.argtypes = [
            ctypes.POINTER(UUID),
            wintypes.DWORD,
            wintypes.HANDLE,
            ctypes.POINTER(ctypes.c_wchar_p)
        ]
        
        path = ctypes.c_wchar_p()
        if _SHGetKnownFolderPath(ctypes.byref(FOLDERID_Documents), 0, 0, ctypes.byref(path)) == 0:
            documents_path = path.value
            print(f"Found Documents folder via Windows API: {documents_path}")
            if os.path.exists(documents_path):
                return documents_path
    except Exception as e:
        print(f"Error getting Documents via Windows API: {e}")
    
    # Try each path
    for path in paths_to_try:
        print(f"Checking path: {path}")
        if os.path.exists(path):
            print(f"Found Documents folder: {path}")
            return path
    
    # Last resort: use home directory
    home_dir = str(Path.home())
    print(f"No standard Documents folder found, using home directory: {home_dir}")
    return home_dir

def main():
    print("Starting mSeq open test...")
    print(f"Platform: {platform.platform()}")
    
    # Initialize components
    config = MseqConfig()
    print("Config loaded")
    file_dao = FileSystemDAO(config)
    print("FileSystemDAO initialized")
    ui_automation = MseqAutomation(config)
    print("UI Automation initialized")
    
    # Get Documents folder using our enhanced function
    documents_path = get_documents_folder()
    
    print(f"Using test folder: {documents_path}")
    
    # List folders in the documents path to verify access
    try:
        print(f"Folders in {documents_path}:")
        for item in os.listdir(documents_path):
            if os.path.isdir(os.path.join(documents_path, item)):
                print(f"  - {item}")
    except Exception as e:
        print(f"Error listing folders: {e}")
    
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
        # Use the improved dialog detection from our updated UI automation class
        browse_dialog = ui_automation._get_browse_dialog()
        
        if not browse_dialog or not browse_dialog.exists():
            print("Falling back to generic dialog detection...")
            timings.wait_until(timeout=10, retry_interval=0.5, 
                             func=lambda: (app.window(title='Browse For Folder').exists() or 
                                          app.window(title_re='Browse.*Folder').exists()), 
                             value=True)
            
            # Try different possible titles
            for title in ['Browse For Folder', 'Browse for Folder']:
                try:
                    browse_dialog = app.window(title=title)
                    if browse_dialog and browse_dialog.exists():
                        break
                except:
                    pass
                    
            # Last resort: Try with regex
            if not browse_dialog or not browse_dialog.exists():
                browse_dialog = app.window(title_re='Browse.*Folder')
        
        if not browse_dialog or not browse_dialog.exists():
            print("ERROR: Could not find Browse dialog despite waiting")
            return
            
        print("Browse For Folder dialog found")
        
        # Print information about the dialog
        print(f"Dialog title: {browse_dialog.window_text()}")
        print(f"Dialog rectangle: {browse_dialog.rectangle()}")
        
        # Navigate to test folder
        print(f"Attempting to navigate to: {documents_path}")
        success = ui_automation.navigate_folder_tree(browse_dialog, documents_path)
        
        if success:
            print("Navigation successful, clicking OK")
            # Try to find OK button with better error handling
            try:
                ok_button = browse_dialog.child_window(title="OK", class_name="Button")
                if not ok_button or not ok_button.exists():
                    # Try alternative approaches
                    for btn_title in ["OK", "Ok", "&OK", "O&K"]:
                        try:
                            ok_button = browse_dialog.child_window(title=btn_title, class_name="Button")
                            if ok_button and ok_button.exists():
                                break
                        except:
                            pass
                
                if ok_button and ok_button.exists():
                    print(f"Found OK button, clicking it")
                    ok_button.click_input()
                else:
                    print("OK button not found, trying to continue with Enter key...")
                    # Try to press Enter key instead
                    browse_dialog.set_focus()
                    send_keys('{ENTER}')
            except Exception as e:
                print(f"Error clicking OK button: {e}")
                # Try to press Enter key
                browse_dialog.set_focus()
                send_keys('{ENTER}')
        else:
            print("Navigation failed")
    
    except timings.TimeoutError:
        print("ERROR: Browse For Folder dialog did not appear")
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
    
    # Keep script running to see what happens
    print("Test completed. Waiting 5 seconds before closing...")
    time.sleep(5)
    
    # Close mSeq
    print("Closing mSeq...")
    ui_automation.close()
    print("Done")

if __name__ == "__main__":
    main()