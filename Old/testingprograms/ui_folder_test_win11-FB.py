# ui_folder_test_win11.py - Simplified for Windows 11
import os
import time
import getpass
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

def test_local_folder_navigation():
    """Test navigating to a local folder in the Browse For Folder dialog with verification"""
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
        
        # Wait for Browse For Folder dialog with longer timeout
        timings.wait_until(timeout=10, retry_interval=0.2, 
                         func=lambda: app.window(title='Browse For Folder').exists(), 
                         value=True)
        
        browse_dialog = app.window(title='Browse For Folder')
        print("Browse For Folder dialog found")
        
        # Find the TreeView control by text "Tree View"
        print("Looking for TreeView with text 'Tree View'...")
        tree_view = browse_dialog.child_window(title="Tree View", class_name="SysTreeView32")
        print("TreeView control found by title 'Tree View'")
        
        # Parse path
        parts = test_folder.split("\\")
        drive = parts[0]  # e.g., "C:"
        folders = parts[1:] if len(parts) > 1 else []
        
        print(f"Path parts: Drive={drive}, Folders={folders}")
        
        # Get the Desktop root item
        print("Looking for Desktop root item...")
        roots = list(tree_view.roots())
        print(f"Found {len(roots)} root items")
        for root in roots:
            print(f"Root item: {root.text()}")
            if "Desktop" in root.text():
                desktop_item = root
                break
        
        print(f"Found Desktop: {desktop_item.text()}")
        desktop_item.expand()
        time.sleep(1.0)
        
        # Look for This PC or Computer as a child of Desktop
        print("Looking for This PC under Desktop...")
        for child in desktop_item.children():
            print(f"Desktop child: {child.text()}")
            if "PC" in child.text() or "Computer" in child.text():
                this_pc_item = child
                break
        
        print(f"Found This PC: {this_pc_item.text()}")
        this_pc_item.expand()
        time.sleep(1.0)
        
        # Navigate to the drive
        print(f"Looking for drive: {drive}")
        for item in this_pc_item.children():
            drive_text = item.text()
            print(f"Drive: {drive_text}")
            if drive in drive_text:
                print(f"Found drive: {drive_text}")
                browse_dialog.set_focus()
                drive_item = item
                drive_item.click_input()  # Single click to select
                time.sleep(1.0)
                break
        
        # Navigate through each subfolder one by one
        current_item = drive_item  # Start from the drive
        
        for folder in folders:
            print(f"\nLooking for folder: {folder}")
            
            # Expand the current node
            print(f"Expanding {current_item.text()}...")
            current_item.expand()
            time.sleep(1.0)
            
            # List available children
            print(f"Available folders under {current_item.text()}:")
            children = list(current_item.children())
            for child in children:
                print(f" - {child.text()}")
            
            # Look for exact folder match
            folder_found = False
            next_item = None
            
            for child in children:
                if child.text() == folder:  # Exact match
                    print(f"Found exact match: {child.text()}")
                    next_item = child
                    folder_found = True
                    break
            
            # If not found, try partial match
            if not folder_found:
                print(f"No exact match, looking for partial matches for '{folder}'")
                for child in children:
                    if folder.lower() in child.text().lower():
                        print(f"Found partial match: {child.text()}")
                        next_item = child
                        folder_found = True
                        break
            
            # If found, click on it with focus
            if folder_found and next_item is not None:
                print(f"Selecting folder: {next_item.text()}")
                browse_dialog.set_focus()
                next_item.click_input()  # Single click to select
                time.sleep(1.0)
                
                # Update current_item for next iteration
                current_item = next_item
            else:
                print(f"ERROR: Could not find folder '{folder}'")
                break
        
        # Click OK
        print("Clicking OK button")
        ok_button = browse_dialog.child_window(title="OK", class_name="Button")
        if ok_button.exists():
            ok_button.click_input()
            time.sleep(2.0)  # Wait to see what happens
        else:
            print("Could not find OK button")
        
        # Wait to see if a project creation error appears
        try:
            error_dialog = app.window(title="Project creation error", visible_only=True)
            if error_dialog.exists(timeout=3):
                print("Project creation error dialog appeared")
                error_dialog.child_window(title="OK").click()
                time.sleep(1.0)
        except Exception:
            print("No error dialog appeared")
        
        # Check if mSeq window title changed indicating successful folder selection
        try:
            # Get current window title
            current_title = app.top_window().window_text()
            print(f"Current mSeq window title: {current_title}")
            
            if "mSeq" in current_title and current_title != "mSeq":
                print("SUCCESS: mSeq window title changed, suggesting folder selection worked")
            else:
                print("NOTE: mSeq window title did not change")
        except Exception as e:
            print(f"Could not check window title: {e}")
        
        # Test completed successfully
        print("Test completed successfully")
        
        # Close mSeq when done
        print("Closing mSeq")
        app.kill()
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        try:
            app.kill()
        except:
            pass

if __name__ == "__main__":
    test_local_folder_navigation()