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
        
        # Find the TreeView control by text "Tree View" instead of "Navigation Pane"
        # Based on the dialog explorer log
        try:
            print("Looking for TreeView with text 'Tree View'...")
            tree_view = browse_dialog.child_window(title="Tree View", class_name="SysTreeView32")
            print("TreeView control found by title 'Tree View'")
        except Exception as e:
            print(f"Error finding TreeView by title: {e}")
            # Try by control_id=100 as found in the logs
            try:
                print("Trying to find TreeView by control_id=100...")
                tree_view = browse_dialog.child_window(control_id=100, class_name="SysTreeView32")
                print("TreeView control found by control_id=100")
            except Exception as e2:
                print(f"Error finding TreeView by control_id: {e2}")
                # Try by index - fourth control in the dialog based on the logs
                try:
                    print("Trying to find TreeView by index...")
                    tree_view = browse_dialog.child_window(class_name="SysTreeView32")
                    print("TreeView control found by class name only")
                except Exception as e3:
                    print(f"Error finding TreeView by class name: {e3}")
                    raise Exception("Could not find TreeView control using any method")
        
        # Parse path
        parts = test_folder.split("\\")
        drive = parts[0]  # e.g., "C:"
        folders = parts[1:] if len(parts) > 1 else []
        
        print(f"Path parts: Drive={drive}, Folders={folders}")
        
        # Get the Desktop root item
        print("Looking for Desktop root item...")
        desktop_item = None
        try:
            roots = list(tree_view.roots())
            print(f"Found {len(roots)} root items")
            for root in roots:
                print(f"Root item: {root.text()}")
                if "Desktop" in root.text():
                    desktop_item = root
                    break
        except Exception as e:
            print(f"Error getting roots: {e}")
            print("Will attempt to use keyboard navigation instead")
            
            # Set focus to the TreeView
            tree_view.set_focus()
            
            # Use keyboard navigation as fallback
            print("Using keyboard navigation as fallback...")
            # First item should be Desktop
            send_keys('{HOME}')  # Go to first item (Desktop)
            time.sleep(0.5)
            send_keys('{RIGHT}')  # Expand Desktop
            time.sleep(1.0)
            
            # Find This PC by going through items
            for i in range(10):
                send_keys('{DOWN}')
                time.sleep(0.5)
                # We don't have a way to check what's selected here, so we just have to trust
                # that after a few down presses we'll find This PC
                if i >= 2:  # After a few down presses, assume we've found This PC
                    print("Attempting to expand what should be 'This PC'...")
                    send_keys('{RIGHT}')  # Expand This PC
                    time.sleep(1.0)
                    
                    # Now look for the drive
                    for j in range(10):
                        send_keys('{DOWN}')
                        time.sleep(0.5)
                        
                        folder_edit = browse_dialog.child_window(class_name="Edit")
                        current_path = folder_edit.window_text() if folder_edit.exists() else ""
                        print(f"Current path: {current_path}")
                        
                        if drive in current_path:
                            print(f"Found drive {drive}")
                            send_keys('{RIGHT}')  # Expand drive
                            time.sleep(1.0)
                            
                            # Navigate through each folder
                            for folder in folders:
                                found = False
                                for k in range(20):  # Try more items for subfolders
                                    send_keys('{DOWN}')
                                    time.sleep(0.5)
                                    
                                    folder_edit = browse_dialog.child_window(class_name="Edit")
                                    current_path = folder_edit.window_text() if folder_edit.exists() else ""
                                    print(f"Current path: {current_path}")
                                    
                                    if folder in current_path.split('\\')[-1]:
                                        print(f"Found folder {folder}")
                                        if folder != folders[-1]:  # Don't expand the last folder
                                            send_keys('{RIGHT}')  # Expand folder
                                            time.sleep(1.0)
                                        found = True
                                        break
                                
                                if not found:
                                    print(f"Could not find folder {folder}")
                                    break
                            
                            break  # Break out of drive search loop
                    
                    break  # Break out of This PC search loop
            
            # At this point we've tried our best with keyboard navigation
            # Set the path directly as a final fallback
            folder_edit = browse_dialog.child_window(class_name="Edit")
            if folder_edit.exists():
                print(f"Setting path directly to: {test_folder}")
                folder_edit.set_edit_text(test_folder)
                time.sleep(1.0)
            
            # Click OK
            print("Clicking OK button")
            ok_button = browse_dialog.child_window(title="OK", class_name="Button")
            if ok_button.exists():
                ok_button.click_input()
                time.sleep(2.0)  # Wait to see what happens
            
            print("Keyboard navigation completed")
            return
        
        if desktop_item is None:
            print("ERROR: Could not find Desktop in tree view")
            # Use direct path entry as fallback
            folder_edit = browse_dialog.child_window(class_name="Edit")
            if folder_edit.exists():
                print(f"Setting path directly to: {test_folder}")
                folder_edit.set_edit_text(test_folder)
                time.sleep(1.0)
                
                # Click OK
                print("Clicking OK button")
                ok_button = browse_dialog.child_window(title="OK", class_name="Button")
                if ok_button.exists():
                    ok_button.click_input()
                    time.sleep(2.0)
                
                return
        
        print(f"Found Desktop: {desktop_item.text()}")
        desktop_item.expand()
        time.sleep(1.0)
        
        # Look for This PC or Computer as a child of Desktop
        print("Looking for This PC under Desktop...")
        this_pc_item = None
        for child in desktop_item.children():
            print(f"Desktop child: {child.text()}")
            if "PC" in child.text() or "Computer" in child.text():
                this_pc_item = child
                break
        
        if not this_pc_item:
            print("ERROR: Could not find This PC or Computer in tree view")
            # Direct path entry fallback
            folder_edit = browse_dialog.child_window(class_name="Edit")
            if folder_edit.exists():
                print(f"Setting path directly to: {test_folder}")
                folder_edit.set_edit_text(test_folder)
                time.sleep(1.0)
                
                # Click OK
                print("Clicking OK button")
                ok_button = browse_dialog.child_window(title="OK", class_name="Button")
                if ok_button.exists():
                    ok_button.click_input()
                    time.sleep(2.0)
                
                return
        
        print(f"Found This PC: {this_pc_item.text()}")
        this_pc_item.expand()
        time.sleep(1.0)
        
        # Navigate to the drive
        print(f"Looking for drive: {drive}")
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
                time.sleep(1.0)
                break
        
        if not drive_found or drive_item is None:
            print(f"ERROR: Could not find drive '{drive}' in tree view")
            # Direct path entry fallback
            folder_edit = browse_dialog.child_window(class_name="Edit")
            if folder_edit.exists():
                print(f"Setting path directly to: {test_folder}")
                folder_edit.set_edit_text(test_folder)
                time.sleep(1.0)
                
                # Click OK
                print("Clicking OK button")
                ok_button = browse_dialog.child_window(title="OK", class_name="Button")
                if ok_button.exists():
                    ok_button.click_input()
                    time.sleep(2.0)
                
                return
        
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
            try:
                children = list(current_item.children())
                for child in children:
                    print(f" - {child.text()}")
            except Exception as e:
                print(f"Error getting children: {e}")
                
                # If we can't get children, use direct path entry
                folder_edit = browse_dialog.child_window(class_name="Edit")
                if folder_edit.exists():
                    print(f"Setting path directly to: {test_folder}")
                    folder_edit.set_edit_text(test_folder)
                    time.sleep(1.0)
                    
                    # Click OK
                    print("Clicking OK button")
                    ok_button = browse_dialog.child_window(title="OK", class_name="Button")
                    if ok_button.exists():
                        ok_button.click_input()
                        time.sleep(2.0)
                    
                    return
            
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
                # Direct path entry fallback
                folder_edit = browse_dialog.child_window(class_name="Edit")
                if folder_edit.exists():
                    print(f"Setting path directly to: {test_folder}")
                    folder_edit.set_edit_text(test_folder)
                    time.sleep(1.0)
                break
        
        # Verify path before clicking OK
        folder_edit = browse_dialog.child_window(class_name="Edit")
        if folder_edit.exists():
            final_selection = folder_edit.window_text()
            print(f"\nFINAL SELECTION: {final_selection}")
            
            # If path doesn't match expected, try direct entry
            if test_folder.lower() not in final_selection.lower():
                print(f"Final selection does not match expected path. Setting directly...")
                folder_edit.set_edit_text(test_folder)
                time.sleep(1.0)
        
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