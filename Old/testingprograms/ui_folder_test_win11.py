# ui_folder_test_win11.py - Optimized for Windows 11
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
        
        # Windows 11 uses "Navigation Pane" as the title for the TreeView
        tree_view = browse_dialog.child_window(title="Navigation Pane", class_name="SysTreeView32")
        print("TreeView control found")
        
        # Parse path
        parts = test_folder.split("\\")
        drive = parts[0]  # e.g., "C:"
        folders = parts[1:] if len(parts) > 1 else []
        
        print(f"Path parts: Drive={drive}, Folders={folders}")
        
        # Start from Desktop which is a root item in Win11
        desktop_item = None
        for root_item in tree_view.roots():
            print(f"Root item: {root_item.text()}")
            if "Desktop" in root_item.text():
                desktop_item = root_item
                break
        
        if not desktop_item:
            print("ERROR: Could not find Desktop in tree view")
            return
        
        print(f"Found Desktop: {desktop_item.text()}")
        desktop_item.expand()
        time.sleep(2.0)  # Longer timeout for Windows 11
        
        # Look for This PC or Computer as a child of Desktop in Win11
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
        time.sleep(2.0)  # Longer timeout
        
        # Navigate to the drive
        drive_found = False
        drive_item = None
        for item in this_pc_item.children():
            drive_text = item.text()
            print(f"Drive: {drive_text}")
            if drive in drive_text:
                print(f"Found drive: {drive_text}")
                browse_dialog.set_focus()
                drive_item = item
                drive_item.click_input()  # Single click to select
                drive_found = True
                time.sleep(2.0)  # Longer timeout
                
                # Verify drive selection
                folder_edit = browse_dialog.child_window(class_name="Edit")
                if folder_edit.exists():
                    current_selection = folder_edit.window_text()
                    print(f"Current selection after drive click: {current_selection}")
                break
        
        if not drive_found or drive_item is None:
            print(f"ERROR: Could not find drive '{drive}' in tree view")
            return
        
        # Navigate through each subfolder one by one with verification
        current_item = drive_item  # Start from the drive
        
        for folder_index, folder in enumerate(folders):
            print(f"\nLooking for folder: {folder} ({folder_index+1}/{len(folders)})")
            
            # Explicitly expand the current node
            print(f"Expanding {current_item.text()}...")
            current_item.expand()
            time.sleep(1.0)  # Wait for expansion
            
            # List available children
            print(f"Available folders under {current_item.text()}:")
            children = []
            try:
                children = list(current_item.children())
                for child in children:
                    print(f" - {child.text()}")
            except Exception as e:
                print(f"Error getting children: {e}")
            
            # If no children found, try expanding again
            if not children:
                print("No children found, trying to expand again...")
                current_item.expand()
                time.sleep(2.0)
                try:
                    children = list(current_item.children())
                    for child in children:
                        print(f" - {child.text()}")
                except Exception as e:
                    print(f"Error getting children on second try: {e}")
            
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
            
            # If found, click on it with focus and verify
            if folder_found and next_item is not None:
                print(f"Selecting folder: {next_item.text()}")
                browse_dialog.set_focus()
                next_item.click_input()  # Single click to select
                time.sleep(1.0)
                
                # Verify selection using the edit box
                folder_edit = browse_dialog.child_window(class_name="Edit")
                if folder_edit.exists():
                    current_selection = folder_edit.window_text()
                    print(f"Current selection after click: {current_selection}")
                    
                    # Check if we have the expected path so far
                    expected_path = os.path.join(drive, *folders[:folder_index+1])
                    if expected_path.lower() not in current_selection.lower():
                        print(f"WARNING: Expected path to contain '{expected_path}' but got '{current_selection}'")
                        
                        # Try using double-click instead
                        print("Trying double-click instead...")
                        browse_dialog.set_focus()
                        next_item.click_input(double=True)
                        time.sleep(1.0)
                        
                        # Check again
                        current_selection = folder_edit.window_text()
                        print(f"Selection after double-click: {current_selection}")
                
                # Update current_item for next iteration
                current_item = next_item
            else:
                print(f"ERROR: Could not find folder '{folder}'")
                return
        
        # Final verification before clicking OK
        folder_edit = browse_dialog.child_window(class_name="Edit")
        if folder_edit.exists():
            final_selection = folder_edit.window_text()
            print(f"\nFINAL SELECTION: {final_selection}")
            
            # Check if we have the complete expected path
            if test_folder.lower() not in final_selection.lower() and "Documents".lower() not in final_selection.lower():
                print(f"ERROR: Final selection does not contain expected path!")
                print(f"Expected: {test_folder}")
                print(f"Got: {final_selection}")
                
                # Try direct path entry if navigation failed
                print("Attempting direct path entry...")
                folder_edit.set_edit_text(test_folder)
                time.sleep(1.0)
            else:
                print("Final selection looks correct")
        
        # Click OK
        print("Clicking OK button")
        ok_button = browse_dialog.child_window(title="OK", class_name="Button")
        if ok_button.exists():
            ok_button.click_input()
            time.sleep(2.0)  # Wait to see what happens
        
        # Close mSeq when done
        print("Test completed, closing mSeq")
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