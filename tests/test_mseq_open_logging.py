# test_mseq_open.py
import os
import sys
import getpass
import logging
from pathlib import Path
import platform
from MseqAuto.mseqauto.config import MseqConfig
from file_system_dao import FileSystemDAO
from folder_processor import FolderProcessor
from ui_automation import MseqAutomation
import time

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("mseq_test.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def get_documents_folder():
    """Get the Documents folder path using multiple methods"""
    current_user = getpass.getuser()
    logger.info(f"Current user: {current_user}")
    
    # Method 1: Windows Shell API
    try:
        logger.info("Trying Windows Shell API method")
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
            logger.info(f"Found Documents folder via Windows API: {documents_path}")
            if os.path.exists(documents_path):
                return documents_path, "Windows Shell API"
    except Exception as e:
        logger.warning(f"Error getting Documents via Windows API: {e}")
    
    # Method 2: Environment variable
    logger.info("Trying environment variable method")
    try:
        env_path = os.path.expandvars("%USERPROFILE%\\Documents")
        logger.info(f"Checking environment path: {env_path}")
        if os.path.exists(env_path):
            logger.info(f"Found Documents via environment variable: {env_path}")
            return env_path, "Environment Variable"
    except Exception as e:
        logger.warning(f"Error checking environment path: {e}")
    
    # Method 3: Standard paths
    standard_paths = [
        # Standard Windows path
        os.path.join("C:", "Users", current_user, "Documents"),
        # Alternative spellings
        os.path.join("C:", "Users", current_user, "Document"),
        os.path.join("C:", "Users", current_user, "My Documents"),
        # OneDrive path
        os.path.join("C:", "Users", current_user, "OneDrive", "Documents"),
        # D: drive possibility
        os.path.join("D:", "Users", current_user, "Documents"),
    ]
    
    logger.info("Trying standard paths method")
    for path in standard_paths:
        logger.info(f"Checking path: {path}")
        if os.path.exists(path):
            logger.info(f"Found Documents via standard path: {path}")
            return path, "Standard Path"
    
    # Method 4: Last resort - home directory
    home_dir = str(Path.home())
    logger.info(f"No standard Documents folder found, using home directory: {home_dir}")
    return home_dir, "Home Directory Fallback"

def main():
    logger.info("=" * 50)
    logger.info("Starting mSeq open test...")
    logger.info(f"Platform: {platform.platform()}")
    
    # Initialize components
    config = MseqConfig()
    logger.info("Config loaded")
    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")
    ui_automation = MseqAutomation(config)
    logger.info("UI Automation initialized")
    
    # Get Documents folder using our enhanced function
    start_time = time.time()
    documents_path, method_used = get_documents_folder()
    
    # OVERRIDE WITH CUSTOM FOLDER PATH HERE
    custom_folder_path = r"P:\Data\Testing\Win11Test"  # Change this to your desired test folder
    test_with_custom_folder = True  # Set to False to use the detected Documents folder
    
    if test_with_custom_folder:
        logger.info(f"Overriding automatically detected path with custom folder: {custom_folder_path}")
        documents_path = custom_folder_path
        method_used = "Manual Override"
    else:
        logger.info(f"Found Documents folder using method: {method_used}")
    
    logger.info(f"Path discovery took {time.time() - start_time:.2f} seconds")
    logger.info(f"Using test folder: {documents_path}")
    
    # List folders in the documents path to verify access
    try:
        folder_count = 0
        logger.info(f"Folders in {documents_path}:")
        for item in os.listdir(documents_path):
            if os.path.isdir(os.path.join(documents_path, item)):
                logger.info(f"  - {item}")
                folder_count += 1
                if folder_count >= 5:  # Just show first 5 folders
                    logger.info(f"  ... and {len([f for f in os.listdir(documents_path) if os.path.isdir(os.path.join(documents_path, f))]) - 5} more")
                    break
    except Exception as e:
        logger.error(f"Error listing folders: {e}")
    
    # Directly try to open mSeq and navigate to the folder
    logger.info("Attempting to open mSeq and navigate...")
    
    # Connect to mSeq or start it
    start_time = time.time()
    logger.info("Connecting to or starting mSeq...")
    app, main_window = ui_automation.connect_or_start_mseq()
    logger.info(f"Connection took {time.time() - start_time:.2f} seconds")
    logger.info(f"mSeq window: {main_window.window_text() if main_window else 'None'}")
    
    # Set focus to main window
    if main_window:
        main_window.set_focus()
        logger.info("Set focus to main window")
    else:
        logger.error("ERROR: No main window found")
        return
    
    # Try to open new project
    logger.info("Pressing Ctrl+N...")
    from pywinauto.keyboard import send_keys
    send_keys('^n')
    
    # Wait for Browse dialog
    logger.info("Waiting for Browse For Folder dialog...")
    try:
        from pywinauto import timings
        start_time = time.time()
        
        # Use the improved dialog detection from our updated UI automation class
        browse_dialog = ui_automation._get_browse_dialog()
        
        if not browse_dialog or not browse_dialog.exists():
            logger.info("Falling back to generic dialog detection...")
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
            logger.error("ERROR: Could not find Browse dialog despite waiting")
            return
            
        logger.info(f"Dialog detection took {time.time() - start_time:.2f} seconds")
        logger.info("Browse For Folder dialog found")
        
        # Print information about the dialog
        logger.info(f"Dialog title: {browse_dialog.window_text()}")
        logger.info(f"Dialog rectangle: {browse_dialog.rectangle()}")
        
        # Navigate to test folder
        logger.info(f"Attempting to navigate to: {documents_path}")
        start_time = time.time()
        success = ui_automation.navigate_folder_tree(browse_dialog, documents_path)
        logger.info(f"Navigation took {time.time() - start_time:.2f} seconds")
        
        if success:
            logger.info("Navigation successful, clicking OK")
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
                    logger.info(f"Found OK button, clicking it")
                    ok_button.click_input()
                else:
                    logger.info("OK button not found, trying to continue with Enter key...")
                    # Try to press Enter key instead
                    browse_dialog.set_focus()
                    send_keys('{ENTER}')
            except Exception as e:
                logger.error(f"Error clicking OK button: {e}")
                # Try to press Enter key
                browse_dialog.set_focus()
                send_keys('{ENTER}')
        else:
            logger.error("Navigation failed")
    
    except timings.TimeoutError:
        logger.error("ERROR: Browse For Folder dialog did not appear")
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        import traceback
        logger.error(traceback.format_exc())
    
    # Keep script running to see what happens
    logger.info("Test completed. Waiting 5 seconds before closing...")
    time.sleep(5)
    
    # Close mSeq
    logger.info("Closing mSeq...")
    ui_automation.close()
    logger.info("Done")
    logger.info("=" * 50)

if __name__ == "__main__":
    main()