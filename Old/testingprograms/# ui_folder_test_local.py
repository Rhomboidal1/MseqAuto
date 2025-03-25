# ui_folder_test_local.py - Test local folder navigation
import os
import time
import getpass
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

def test_local_folder_navigation():
    """Test navigating to a local folder in the Browse For Folder dialog"""
    # Use a folder that definitely exists on your machine
    test_folder = r"C:\Users\{}\Documents".format(getpass.getuser())
    print(f"Testing navigation to: {test_folder}")
    
    try:
        # Start mSeq application
        app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False)
        app.connect(title='mSeq', timeout=10)
        
        # Get the main window
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
        timings.wait_until(timeout=5, retry_interval=0.1, 
                         func=lambda: app.window(title='Browse For Folder').exists(), 
                         value=True)
        
        browse_dialog = app.window(title='Browse For Folder')
        print("Browse For Folder dialog found")
        
        # Get TreeView control
        tree_view = browse_dialog.child_window(class_name="SysTreeView32")
        
        # Parse path
        parts = test_folder.split("\\")
        drive = parts[0]  # e.g., "C:"
        folders = parts[1:] if len(parts) > 1 else []
        
        print(f"Path parts: Drive={drive}, Folders={folders}")
        
        # Start from Desktop
        desktop_item = None
        for root_item in tree_view.roots():
            if "Desktop" in root_item.text():
                desktop_item = root_item
                break
        
        if not desktop_item:
            print("ERROR: Could not find Desktop in tree view")
            tree_view.print_items()  # Print all items to diagnose
            return
        
        print(f"Found Desktop: {desktop_item.text()}")
        desktop_item.expand()
        time.sleep(2.0)
        
        # Look for This PC or Computer
        this_pc_item = None
        for child in desktop_item.children():
            print(f"Desktop child: {child.text()}")
            if "PC" in child.text() or "Computer" in child.text():
                this_pc_item = child
                break
        
        if not this_pc_item:
            print("ERROR: Could not find This PC or Computer in tree view")
            return
        
        print(f"Found This PC: {this_pc_item.text()}")
        this_pc_item.expand()
        time.sleep(2.0)
        
        # Navigate to the drive
        drive_found = False
        for drive_item in this_pc_item.children():
            drive_text = drive_item.text()
            print(f"Drive: {drive_text}")
            if drive in drive_text:
                print(f"Found drive: {drive_text}")
                browse_dialog.set_focus()
                drive_item.click_input()
                drive_found = True
                current_item = drive_item
                time.sleep(2.0)
                break
        
        if not drive_found:
            print(f"ERROR: Could not find drive '{drive}' in tree view")
            return
        
        # Navigate through subfolders
        for folder in folders:
            print(f"Looking for folder: {folder}")
            current_item.expand()
            time.sleep(2.0)
            
            children = list(current_item.children())
            print(f"Available folders:")
            for child in children:
                print(f" - {child.text()}")
            
            folder_found = False
            for child in children:
                if folder in child.text():
                    print(f"Found folder: {child.text()}")
                    browse_dialog.set_focus()
                    child.click_input()
                    folder_found = True
                    current_item = child
                    time.sleep(2.0)
                    break
            
            if not folder_found:
                print(f"ERROR: Could not find folder '{folder}'")
                break
        
        # Successfully navigated to the folder
        print("Navigation completed successfully")
        
        # Cancel dialog
        cancel_button = browse_dialog.child_window(title="Cancel", class_name="Button")
        cancel_button.click_input()
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_local_folder_navigation()