# ui_folder_test.py - Improved folder navigation test
import os
import time
import sys
import getpass
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

# Define the drive mappings
DRIVE_MAPPINGS = {
    "P:": "ABISync (P:)",
    "H:": f"{getpass.getuser()} (\\\\w2k16\\users) (H:)"
}

def test_folder_navigation(test_folder=None):
    """Test navigating to a folder in the Browse For Folder dialog with detailed logging"""
    # If no test folder provided, use a folder we know exists
    if not test_folder:
        # Use a folder that's guaranteed to exist
        test_folder = r"P:\Data"
    
    print(f"Starting folder navigation test to: {test_folder}")
    
    try:
        # Try to connect to existing mSeq or start a new instance
        try:
            app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
            print("Connected to existing mSeq application")
        except Exception as e:
            print(f"Connection error: {e}")
            print("Starting new mSeq instance...")
            app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False)
            app.connect(title='mSeq', timeout=10)
        
        # Get the main window
        print("Finding main window...")
        if app.window(title_re='mSeq*').exists():
            main_window = app.window(title_re='mSeq*')
        else:
            main_window = app.window(title_re='Mseq*')
        
        print(f"Main window found: {main_window.window_text()}")
        main_window.set_focus()
        
        # Open New Project dialog
        print("Pressing Ctrl+N...")
        send_keys('^n')
        
        # Wait for Browse For Folder dialog
        print("Waiting for Browse For Folder dialog...")
        try:
            timings.wait_until(timeout=5, retry_interval=0.1, 
                            func=lambda: app.window(title='Browse For Folder').exists(), 
                            value=True)
            
            browse_dialog = app.window(title='Browse For Folder')
            print("Browse For Folder dialog found")
            
            # Get TreeView control
            tree_view = browse_dialog.child_window(class_name="SysTreeView32")
            if not tree_view.exists():
                print("ERROR: TreeView control not found")
                return
            
            print("TreeView control found")
            
            # Get the virtual folder name
            import win32com.client
            shell = win32com.client.Dispatch("Shell.Application")
            namespace = shell.Namespace(0x11)  # CSIDL_DRIVES
            virtual_folder = namespace.Title
            desktop_path = f'\\Desktop\\{virtual_folder}'
            
            print(f"Starting navigation from: {desktop_path}")
            
            # Print all available root items
            try:
                # Try to expand the desktop first
                desktop_item = tree_view.get_item('\\Desktop')
                print("Found Desktop")
                
                # Log all top-level items in the tree
                print("\nExploring TreeView structure:")
                print("Root items:")
                for root_item in tree_view.roots():
                    print(f" - {root_item.text()}")
                    
                    # Try to expand and list children
                    try:
                        root_item.expand()
                        time.sleep(0.5)
                        print(f"   Children of {root_item.text()}:")
                        for child in root_item.children():
                            print(f"    - {child.text()}")
                    except Exception as e:
                        print(f"   Error expanding {root_item.text()}: {e}")
                
                # Start from Desktop
                desktop_item = None
                for root_item in tree_view.roots():
                    if "Desktop" in root_item.text():
                        desktop_item = root_item
                        break
                
                if not desktop_item:
                    print("ERROR: Could not find Desktop in tree view")
                    return
                
                print(f"\nFound Desktop: {desktop_item.text()}")
                desktop_item.expand()
                time.sleep(0.5)
                
                # Look for This PC or Computer
                this_pc_item = None
                for child in desktop_item.children():
                    if "PC" in child.text() or "Computer" in child.text() or virtual_folder in child.text():
                        this_pc_item = child
                        break
                
                if not this_pc_item:
                    print("ERROR: Could not find This PC or Computer in tree view")
                    return
                
                print(f"Found This PC: {this_pc_item.text()}")
                this_pc_item.expand()
                time.sleep(0.5)
                
                # List all drives
                print("\nAvailable drives:")
                for drive_item in this_pc_item.children():
                    print(f" - {drive_item.text()}")
                
                # Parse the path
                if ":" in test_folder:
                    # Path has a drive letter
                    parts = test_folder.split("\\")
                    drive = parts[0]  # e.g., "P:"
                    folders = parts[1:] if len(parts) > 1 else []
                else:
                    # Network path
                    parts = test_folder.split("\\")
                    drive = "\\" + "\\".join(parts[:3])  # e.g., \\server\share
                    folders = parts[3:] if len(parts) > 3 else []
                
                print(f"\nPath analysis: Drive={drive}, Folders={folders}")
                
                # First navigate to the drive
                drive_found = False
                mapped_name = DRIVE_MAPPINGS.get(drive, None)
                
                for drive_item in this_pc_item.children():
                    drive_text = drive_item.text()
                    print(f"Checking drive: {drive_text}")
                    
                    # Check for both exact match and mapped name
                    if (drive == drive_text or 
                        drive in drive_text or 
                        (mapped_name and mapped_name in drive_text)):
                            
                        print(f"Found matching drive: {drive_text}")
                        browse_dialog.set_focus()
                        drive_item.click_input()
                        drive_found = True
                        current_item = drive_item
                        time.sleep(0.5)
                        break
                
                if not drive_found:
                    print(f"ERROR: Could not find drive '{drive}' or '{mapped_name}' in tree view")
                    # Cancel the dialog
                    cancel_button = browse_dialog.child_window(title="Cancel", class_name="Button")
                    if cancel_button.exists():
                        cancel_button.click_input()
                    return
                
                # Navigate through subfolders
                for folder in folders:
                    print(f"\nNavigating to subfolder: {folder}")
                    current_item.expand()
                    time.sleep(0.5)
                    
                    # Show current folder children
                    print("Available subfolders:")
                    children = list(current_item.children())
                    for child in children:
                        print(f" - {child.text()}")
                    
                    # Look for exact match first
                    folder_found = False
                    for child in children:
                        if child.text() == folder:
                            print(f"Found folder: {folder}")
                            browse_dialog.set_focus()
                            child.click_input()
                            folder_found = True
                            current_item = child
                            time.sleep(0.5)
                            break
                    
                    if not folder_found:
                        print(f"WARNING: Could not find exact match for folder '{folder}'")
                        
                        # Try partial match
                        for child in children:
                            if folder.lower() in child.text().lower():
                                print(f"Found partial match: {child.text()}")
                                browse_dialog.set_focus()
                                child.click_input()
                                folder_found = True
                                current_item = child
                                time.sleep(0.5)
                                break
                        
                        if not folder_found:
                            print(f"ERROR: Could not find folder '{folder}' or similar")
                            # Continue with the rest of the path - don't cancel yet
                
                # Was navigation successful?
                if folder_found or (len(folders) == 0 and drive_found):
                    print("\nNavigation completed successfully!")
                    
                    # Optional: Click OK to see if the folder is valid for mSeq
                    # Uncomment the lines below to test this
                    # ok_button = browse_dialog.child_window(title="OK", class_name="Button")
                    # if ok_button.exists():
                    #     ok_button.click_input()
                    #     print("Clicked OK button")
                else:
                    print("\nNavigation was incomplete")
                
                # Cancel the dialog
                cancel_button = browse_dialog.child_window(title="Cancel", class_name="Button")
                if cancel_button.exists():
                    cancel_button.click_input()
                    print("Clicked Cancel button")
                
            except Exception as e:
                print(f"Error during TreeView exploration: {type(e).__name__}: {e}")
                import traceback
                traceback.print_exc()
        
        except timings.TimeoutError:
            print("ERROR: Browse For Folder dialog did not appear")
        
    except Exception as e:
        print(f"General error: {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    # If a folder path is provided as an argument, use it
    if len(sys.argv) > 1:
        test_folder = sys.argv[1]
    else:
        # Default test path
        test_folder = r"P:\Data"
    
    test_folder_navigation(test_folder)