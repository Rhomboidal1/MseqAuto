# ui_automation.py
import os
import time
import platform
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError

class MseqAutomation:
    def __init__(self, config):
        self.config = config
        self.app = None
        self.main_window = None
        self.first_time_browsing = True
        
        # Detect Windows version
        win_version = int(platform.version().split('.')[0])
        win_build = int(platform.version().split('.')[2]) if len(platform.version().split('.')) > 2 else 0
        self.is_win11 = win_version >= 10 and win_build >= 22000
    
    def connect_or_start_mseq(self):
        """Connect to existing mSeq or start a new instance"""
        try:
            self.app = Application(backend='win32').connect(title_re='Mseq*', timeout=1)
        except (ElementNotFoundError, timings.TimeoutError):
            try:
                self.app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
            except (ElementNotFoundError, timings.TimeoutError):
                start_cmd = f'cmd /c "cd /d {self.config.MSEQ_PATH} && {self.config.MSEQ_EXECUTABLE}"'
                self.app = Application(backend='win32').start(start_cmd, wait_for_idle=False) 
                self.app.connect(title='mSeq', timeout=10)
            except ElementAmbiguousError:
                self.app = Application(backend='win32').connect(title_re='mSeq*', found_index=0, timeout=1)
                app2 = Application(backend='win32').connect(title_re='mSeq*', found_index=1, timeout=1)
                app2.kill()
        except ElementAmbiguousError:
            self.app = Application(backend='win32').connect(title_re='Mseq*', found_index=0, timeout=1)
            app2 = Application(backend='win32').connect(title_re='Mseq*', found_index=1, timeout=1)
            app2.kill()
        
        # Get the main window
        if not self.app.window(title_re='mSeq*').exists():
            self.main_window = self.app.window(title_re='Mseq*')
        else:
            self.main_window = self.app.window(title_re='mSeq*')
            
        return self.app, self.main_window
    
    def wait_for_dialog(self, dialog_type):
        """Wait for a specific dialog to appear with Windows 11 compatibility"""
        timeout = self.config.TIMEOUTS.get(dialog_type, 5)
        # Add 2 seconds to timeout for Windows 11 to accommodate possible slower response
        if hasattr(self, 'is_win11') and self.is_win11:
            timeout += 2
        
        if dialog_type == "browse_dialog":
            try:
                return timings.wait_until(timeout=timeout, retry_interval=0.1, 
                                     func=lambda: (self.app.window(title='Browse For Folder').exists() or 
                                                  self.app.window(title_re='Browse.*Folder').exists()), 
                                     value=True)
            except timings.TimeoutError:
                # Additional fallback mechanism
                for i in range(int(timeout * 10)):
                    if (self.app.window(title='Browse For Folder').exists() or 
                        self.app.window(title_re='Browse.*Folder').exists()):
                        return True
                    time.sleep(0.1)
                return False
        elif dialog_type == "preferences":
            return timings.wait_until(timeout=timeout, retry_interval=0.1, 
                                     func=lambda: (self.app.window(title='Mseq Preferences').exists() or
                                                  self.app.window(title='mSeq Preferences').exists()), 
                                     value=True)
        elif dialog_type == "copy_files":
            return timings.wait_until(timeout=timeout, retry_interval=0.1,
                                     func=lambda: self.app.window(title_re='Copy.*sequence files').exists(),
                                     value=True)
        elif dialog_type == "error_window":
            return timings.wait_until(timeout=timeout, retry_interval=0.3,
                                     func=lambda: (self.app.window(title='File error').exists() or
                                                  self.app.window(title_re='.*[Ee]rror.*').exists()),
                                     value=True)
        elif dialog_type == "call_bases":
            return timings.wait_until(timeout=timeout, retry_interval=0.3,
                                     func=lambda: self.app.window(title_re='Call bases.*').exists(),
                                     value=True)
        elif dialog_type == "read_info":
            return timings.wait_until(timeout=timeout, retry_interval=0.1,
                                     func=lambda: self.app.window(title_re='Read information for.*').exists(),
                                     value=True)
    
    def _ensure_dialog_visible(self, dialog):
        """Make sure dialog is visible and positioned correctly"""
        try:
            # Check if dialog exists and is positioned properly
            if dialog.exists() and dialog.rectangle().width() > 0:
                # Dialog is visible, no action needed
                return True
            
            # Try to set focus and move if needed
            dialog.set_focus()
            
            # Get screen dimensions
            try:
                import win32api
                screen_width = win32api.GetSystemMetrics(0)
                screen_height = win32api.GetSystemMetrics(1)
                
                # Reposition dialog if it's offscreen
                rect = dialog.rectangle()
                if rect.left < 0 or rect.top < 0 or rect.right > screen_width or rect.bottom > screen_height:
                    dialog.move_window(100, 100, rect.width(), rect.height())
            except ImportError:
                # Fall back if win32api isn't available
                dialog.set_focus()
                
            return True
        except Exception as e:
            print(f"Error ensuring dialog visibility: {e}")
            return False
    
    def _get_tree_view(self, dialog):
        """Get tree view control regardless of Windows version"""
        import logging
        logger = logging.getLogger(__name__)
        
        logger.info("Attempting to find tree view control")
        
        # Try to identify the tree view using only class name first (most reliable)
        try:
            # Windows 11 more commonly uses SysTreeView32 without a specific title
            tree_control = dialog.child_window(class_name="SysTreeView32")
            if tree_control.exists():
                logger.info("Found tree view control by class name only")
                return tree_control
        except Exception as e:
            logger.warning(f"Could not find tree view by class name only: {e}")
        
        # Try different names used in Windows 10/11
        for name in ["Navigation Pane", "Tree View"]:
            try:
                logger.info(f"Trying to find tree view with title: {name}")
                tree_control = dialog.child_window(title=name, class_name="SysTreeView32")
                if tree_control.exists():
                    logger.info(f"Found tree view with title: {name}")
                    return tree_control
            except Exception as e:
                logger.warning(f"Could not find tree view with title {name}: {e}")
        
        # Special handling for the SHBrowseForFolder control which contains the tree view
        try:
            logger.info("Trying to find via SHBrowseForFolder control")
            shell_control = dialog.child_window(class_name="SHBrowseForFolder ShellNameSpace Control")
            if shell_control.exists():
                logger.info("Found SHBrowseForFolder control, looking for tree view inside")
                tree_control = shell_control.child_window(class_name="SysTreeView32")
                if tree_control.exists():
                    logger.info("Found tree view inside SHBrowseForFolder")
                    return tree_control
        except Exception as e:
            logger.warning(f"Could not find tree view via SHBrowseForFolder: {e}")
        
        # Last resort - try to find ANY SysTreeView32 control in the dialog
        try:
            logger.info("Last resort: trying to find ANY SysTreeView32 control")
            controls = dialog.children(class_name="SysTreeView32")
            if controls and len(controls) > 0:
                logger.info(f"Found {len(controls)} potential tree view controls, using first one")
                return controls[0]
        except Exception as e:
            logger.error(f"Failed to find ANY tree view control: {e}")
        
        logger.error("Could not find tree view control in the dialog")
        return None
    
    def _get_this_pc_item(self, desktop_item):
        """Get This PC node - handles Windows 10/11 differences"""
        for child in desktop_item.children():
            # Windows 11 might use 'This PC' directly
            # Windows 10 might use 'Computer' or include 'PC'
            if any(x in child.text() for x in ["PC", "Computer"]):
                return child
        
        # Last resort: Try to get the item by index (typically 3rd item)
        try:
            children = desktop_item.children()
            if len(children) >= 3:
                return children[2]  # Often This PC is the 3rd item
        except:
            pass
            
        return None
    
    def is_process_complete(self, folder_path):
        """Check if mSeq has finished processing the folder"""
        # Check both possible dialog titles
        if (self.app.window(title="Low quality files skipped").exists() or
            self.app.window(title_re=".*quality.*skipped").exists()):
            return True
        
        # Check if all 5 text files have been created
        count = 0
        for item in os.listdir(folder_path):
            if os.path.isfile(os.path.join(folder_path, item)):
                for extension in self.config.TEXT_FILES:
                    if item.endswith(extension):
                        count += 1
        return count == 5
    
    def _get_browse_dialog(self):
        """Get the browse dialog window with Win10/Win11 compatibility"""
        # Try different possible titles
        for title in ['Browse For Folder', 'Browse for Folder']:
            try:
                dialog = self.app.window(title=title)
                if dialog.exists():
                    return dialog
            except:
                pass
                
        # Last resort: Try with regex
        try:
            return self.app.window(title_re='Browse.*Folder')
        except:
            return None
    
    def navigate_folder_tree(self, dialog, path):
        """Navigate the folder tree in a file dialog with better Windows 11 support"""
        import logging
        logger = logging.getLogger(__name__)
        
        logger.info(f"Starting folder navigation to: {path}")
        dialog.set_focus()
        self._ensure_dialog_visible(dialog)
        
        # Get tree view using the robust method
        tree_view = self._get_tree_view(dialog)
        if not tree_view:
            logger.error("Could not find TreeView control")
            return False
        
        logger.info("TreeView control found")
        
        # Handle different path formats
        if ":" in path:
            # Path has a drive letter
            parts = path.split("\\")
            drive = parts[0]  # e.g., "P:"
            folders = parts[1:] if len(parts) > 1 else []
            logger.info(f"Path parsed: Drive={drive}, Folders={folders}")
        else:
            # Network path
            parts = path.split("\\")
            drive = "\\" + "\\".join(parts[:3])  # e.g., \\server\share
            folders = parts[3:] if len(parts) > 3 else []
            logger.info(f"Network path parsed: Share={drive}, Folders={folders}")
        
        try:
            # Try to get all root items
            root_items = list(tree_view.roots())
            logger.info(f"Found {len(root_items)} root items in tree")
            
            # Log all root items to understand the hierarchy
            for i, item in enumerate(root_items):
                logger.info(f"Root item {i}: {item.text()}")
                
            # Windows 11 detection
            is_win11 = hasattr(self, 'is_win11') and self.is_win11
            if not is_win11:
                # Try to detect Windows 11 by looking at the children
                win11_indicators = ['Home', 'Gallery']
                root_texts = [item.text() for item in root_items]
                if any(indicator in root_texts for indicator in win11_indicators):
                    logger.info("Windows 11 detected from folder structure")
                    is_win11 = True
            
            # Look for Desktop item
            desktop_item = None
            for root_item in root_items:
                if "Desktop" in root_item.text():
                    desktop_item = root_item
                    logger.info(f"Found Desktop: {root_item.text()}")
                    break
    
            if not desktop_item:
                logger.error("Could not find Desktop in tree view")
                # Try the original approach as fallback
                try:
                    desktop_item = tree_view.get_item('\\Desktop')
                    logger.info("Found Desktop using fallback method")
                except Exception as e:
                    logger.warning(f"Could not find Desktop in tree view: {e}")
            
            # Windows 11: Try to directly go to This PC from root if possible
            this_pc_item = None
            current_item = None
            goto_drive = True  # Flag to determine if we need to navigate to drive
            
            if is_win11:
                logger.info("Using Windows 11 specific approach")
                
                # Try direct edit box input first
                try:
                    edit_box = dialog.child_window(class_name="Edit")
                    if edit_box.exists():
                        logger.info("Found edit box, attempting to set path directly")
                        edit_box.set_edit_text(path)
                        time.sleep(0.5)
                        # Press tab to confirm the path
                        from pywinauto.keyboard import send_keys
                        send_keys("{TAB}")
                        time.sleep(0.5)
                        # If the path is valid, this should work
                        logger.info("Set path via edit box, checking if valid")
                        return True
                except Exception as e:
                    logger.warning(f"Error using edit box: {e}")
                
                # If we found Desktop, expand it to look for all children
                if desktop_item:
                    logger.info("Clicking and expanding Desktop")
                    dialog.set_focus()
                    desktop_item.click_input()
                    desktop_item.expand()
                    time.sleep(1.0)
                    
                    # Get all Desktop children
                    try:
                        desktop_children = list(desktop_item.children())
                        logger.info(f"Desktop has {len(desktop_children)} children")
                        
                        # Log all Desktop children to understand the hierarchy
                        for i, child in enumerate(desktop_children):
                            logger.info(f"Desktop child {i}: {child.text()}")
                        
                        # Looking for This PC (known to be the 12th item in Win11)
                        if len(desktop_children) >= 12:
                            candidate = desktop_children[11]  # 0-based indexing
                            if "PC" in candidate.text():
                                this_pc_item = candidate
                                logger.info(f"Found This PC at position 11: {candidate.text()}")
                        
                        # If This PC not found at position 11, try looking by name
                        if not this_pc_item:
                            for child in desktop_children:
                                if "PC" in child.text():
                                    this_pc_item = child
                                    logger.info(f"Found This PC by name: {child.text()}")
                                    break
                        
                        # If Documents is the target and we can find it directly
                        if "Documents" in path and path.endswith("Documents"):
                            for child in desktop_children:
                                if "Document" in child.text():
                                    logger.info(f"Found Documents folder directly: {child.text()}")
                                    dialog.set_focus()
                                    child.click_input()
                                    time.sleep(0.5)
                                    logger.info("Target is Documents, navigation complete")
                                    return True
                    except Exception as e:
                        logger.warning(f"Error examining Desktop children: {e}")
            
            # Standard Windows 10 flow, or continue Windows 11 flow if we found This PC
            if desktop_item and (not is_win11 or this_pc_item):
                if this_pc_item:
                    logger.info("Using This PC found in Desktop")
                else:
                    # For Windows 10, find This PC in the normal way
                    logger.info("Looking for This PC in Desktop children (Win10 flow)")
                    try:
                        desktop_children = list(desktop_item.children())
                        for child in desktop_children:
                            if any(pc_name in child.text() for pc_name in ["PC", "Computer"]):
                                this_pc_item = child
                                logger.info(f"Found This PC: {child.text()}")
                                break
                    except Exception as e:
                        logger.warning(f"Error finding This PC: {e}")
                
                if not this_pc_item:
                    logger.error("Could not find This PC in Desktop children")
                    return False
                
                logger.info("Clicking and expanding This PC")
                dialog.set_focus()
                this_pc_item.click_input()
                this_pc_item.expand()
                time.sleep(1.0)
                
                # Check what drives are available
                try:
                    drive_children = list(this_pc_item.children())
                    logger.info(f"This PC has {len(drive_children)} children (drives)")
                    for i, drive_item in enumerate(drive_children):
                        logger.info(f"Drive {i}: {drive_item.text()}")
                except Exception as e:
                    logger.warning(f"Error enumerating drives: {e}")
                
                # Set current item to This PC
                current_item = this_pc_item
            
            # If we don't have a current_item yet, we're in trouble
            if not current_item:
                logger.error("Failed to establish a starting point for navigation")
                return False
            
            # Navigate to the drive
            drive_found = False
            mapped_name = self.config.NETWORK_DRIVES.get(drive, None)
            logger.info(f"Looking for drive '{drive}' or mapped name '{mapped_name}'")
            
            try:
                # Get current item's children
                current_item_children = list(current_item.children())
                
                # First pass: look for exact match
                for item in current_item_children:
                    drive_text = item.text()
                    if drive == drive_text or (mapped_name and mapped_name == drive_text):
                        dialog.set_focus()
                        item.click_input()
                        drive_found = True
                        current_item = item
                        time.sleep(1.0)
                        logger.info(f"Found exact drive match: {drive_text}")
                        break
                
                # Second pass: look for partial match
                if not drive_found:
                    for item in current_item_children:
                        drive_text = item.text()
                        if drive in drive_text or (mapped_name and mapped_name in drive_text):
                            dialog.set_focus()
                            item.click_input()
                            drive_found = True
                            current_item = item
                            time.sleep(1.0)
                            logger.info(f"Found partial drive match: {drive_text}")
                            break
            except Exception as e:
                logger.error(f"Error finding drive: {e}")
            
            if not drive_found:
                # Try scrolling and looking again
                try:
                    self._scroll_if_needed(current_item)
                    logger.info("Scrolled to find more drives")
                    
                    # Second attempt after scrolling
                    current_item_children = list(current_item.children())
                    for item in current_item_children:
                        drive_text = item.text()
                        if (drive == drive_text or 
                            drive in drive_text or 
                            (mapped_name and mapped_name in drive_text)):
                            
                            dialog.set_focus()
                            item.click_input()
                            drive_found = True
                            current_item = item
                            time.sleep(1.0)
                            logger.info(f"Found drive after scrolling: {drive_text}")
                            break
                except Exception as e:
                    logger.warning(f"Error while scrolling for drives: {e}")
            
            if not drive_found:
                # Final attempt: try direct path input
                try:
                    edit_box = dialog.child_window(class_name="Edit")
                    if edit_box.exists():
                        logger.info("Using edit box to set path directly")
                        edit_box.set_edit_text(path)
                        time.sleep(0.5)
                        from pywinauto.keyboard import send_keys
                        send_keys("{TAB}")
                        return True
                except Exception as e:
                    logger.warning(f"Error using edit box fallback: {e}")
                
                logger.error(f"Could not find drive '{drive}' or '{mapped_name}' in tree view")
                return False
            
            # Navigate through subfolders
            logger.info(f"Starting subfolder navigation with {len(folders)} folders")
            for i, folder in enumerate(folders):
                logger.info(f"Navigating to folder {i+1}/{len(folders)}: {folder}")
                
                current_item.expand()
                time.sleep(1.0)
                
                # Look for exact match first
                folder_found = False
                
                try:
                    folder_children = list(current_item.children())
                    logger.info(f"Current folder has {len(folder_children)} children")
                    
                    # Log a few child items
                    for j in range(min(5, len(folder_children))):
                        logger.info(f"Child {j}: {folder_children[j].text()}")
                    
                    # Look for exact match
                    for child in folder_children:
                        if child.text() == folder:
                            dialog.set_focus()
                            child.click_input()
                            folder_found = True
                            current_item = child
                            time.sleep(1.0)
                            logger.info(f"Found exact match for {folder}")
                            break
                except Exception as e:
                    logger.warning(f"Error enumerating children: {e}")
                
                if not folder_found:
                    # Try scrolling
                    try:
                        self._scroll_if_needed(current_item)
                        logger.info("Scrolled to find more folders")
                        
                        # Try partial match after scrolling
                        folder_children = list(current_item.children())
                        for child in folder_children:
                            if folder.lower() in child.text().lower():
                                dialog.set_focus()
                                child.click_input()
                                folder_found = True
                                current_item = child
                                time.sleep(1.0)
                                logger.info(f"Found partial match for {folder}: {child.text()}")
                                break
                    except Exception as e:
                        logger.warning(f"Error while scrolling: {e}")
                
                # Check if we found the folder
                if folder_found:
                    # Small delay to ensure the UI updates
                    time.sleep(0.5)
                else:
                    # Windows 11 fallback: try edit box
                    if is_win11:
                        try:
                            logger.info("Folder not found in tree, trying edit box")
                            edit_box = dialog.child_window(class_name="Edit")
                            if edit_box.exists():
                                logger.info("Using edit box to set path directly")
                                edit_box.set_edit_text(path)
                                time.sleep(0.5)
                                from pywinauto.keyboard import send_keys
                                send_keys("{TAB}")
                                return True
                        except Exception as e:
                            logger.warning(f"Error using edit box fallback: {e}")
                    
                    logger.error(f"Could not find folder '{folder}' or similar")
                    return False
            
            # Final folder should now be selected
            logger.info("Navigation completed successfully")
            return True
            
        except Exception as e:
            # Log the error but don't raise it - we want to continue even if navigation fails
            logger.error(f"Error during folder navigation: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False
    
    def process_folder(self, folder_path):
        """Process a folder with mSeq"""
        try:
            if not os.path.exists(folder_path):
                print(f"Warning: Folder does not exist: {folder_path}")
                return False
                
            # Check if there are any .ab1 files to process
            ab1_files = [f for f in os.listdir(folder_path) if f.endswith('.ab1')]
            if not ab1_files:
                print(f"No .ab1 files found in {folder_path}, skipping processing")
                return False
        except Exception as e:
            print(f"Error checking folder {folder_path}: {e}")
            return False
        
        self.app, self.main_window = self.connect_or_start_mseq()
        self.main_window.set_focus()
        send_keys('^n')  # Ctrl+N for new project
        
        # Wait for and handle Browse For Folder dialog
        self.wait_for_dialog("browse_dialog")
        dialog_window = self._get_browse_dialog()
        
        if not dialog_window:
            print("Failed to find Browse For Folder dialog")
            return False
            
        # Add a delay for the first browsing operation
        if self.first_time_browsing:
            self.first_time_browsing = False
            time.sleep(0.3)  # Increased time for Windows 11 compatibility
        else:
            time.sleep(1.5)  # Increased time for Windows 11 compatibility
        
        # Navigate to the target folder
        navigate_success = self.navigate_folder_tree(dialog_window, folder_path)
        if not navigate_success:
            print(f"Navigation failed for {folder_path}")
            # Try to continue anyway
        
        # Find and click OK button with better error handling
        try:
            ok_button = dialog_window.child_window(title="OK", class_name="Button")
            if not ok_button.exists():
                # Try alternative approaches
                for btn_title in ["OK", "Ok", "&OK", "O&K"]:
                    ok_button = dialog_window.child_window(title=btn_title, class_name="Button")
                    if ok_button.exists():
                        break
            
            if ok_button.exists():
                ok_button.click_input()
            else:
                print("OK button not found, trying to continue...")
                # Try to press Enter key instead
                dialog_window.set_focus()
                send_keys('{ENTER}')
        except Exception as e:
            print(f"Error clicking OK button: {e}")
            # Try to press Enter key
            dialog_window.set_focus()
            send_keys('{ENTER}')
        
        # Handle mSeq Preferences dialog
        try:
            self.wait_for_dialog("preferences")
            pref_window = None
            
            for title in ['Mseq Preferences', 'mSeq Preferences']:
                try:
                    pref_window = self.app.window(title=title)
                    if pref_window.exists():
                        break
                except:
                    pass
            
            if pref_window and pref_window.exists():
                # Find and click OK button
                ok_button = None
                for btn_title in ["&OK", "OK", "Ok"]:
                    try:
                        ok_button = pref_window.child_window(title=btn_title, class_name="Button")
                        if ok_button.exists():
                            break
                    except:
                        pass
                
                if ok_button and ok_button.exists():
                    ok_button.click_input()
                else:
                    # Fallback to Enter key
                    pref_window.set_focus()
                    send_keys('{ENTER}')
            else:
                print("Preferences dialog not found, trying to continue...")
        except Exception as e:
            print(f"Error with preferences dialog: {e}")
        
        # Handle Copy sequence files dialog with improved handling
        try:
            self.wait_for_dialog("copy_files")
            copy_files_window = self.app.window(title_re='Copy.*sequence files')
            
            # Different ways to access list view depending on Windows version
            try:
                # Windows 10 approach
                shell_view = copy_files_window.child_window(title="ShellView", class_name="SHELLDLL_DefView")
                list_view = shell_view.child_window(class_name="DirectUIHWND")
                list_view.click_input()
            except:
                try:
                    # Windows 11 approach
                    list_view = copy_files_window.child_window(class_name="DirectUIHWND")
                    list_view.click_input()
                except:
                    # Last resort - try clicking in the middle of the dialog
                    rect = copy_files_window.rectangle()
                    copy_files_window.click_input(coords=((rect.right - rect.left)//2, 
                                                        (rect.bottom - rect.top)//2))
            
            # Select all files
            send_keys('^a')  # Select all files
            
            # Click Open button with better error handling
            open_button = None
            for btn_title in ["&Open", "Open"]:
                try:
                    open_button = copy_files_window.child_window(title=btn_title, class_name="Button")
                    if open_button.exists():
                        break
                except:
                    pass
            
            if open_button and open_button.exists():
                open_button.click_input()
            else:
                # Fallback to Enter key
                copy_files_window.set_focus()
                send_keys('{ENTER}')
        except Exception as e:
            print(f"Error with copy files dialog: {e}")
        
        # Handle File error dialog with improved handling
        try:
            self.wait_for_dialog("error_window")
            error_window = None
            
            for title in ['File error', 'Error']:
                try:
                    error_window = self.app.window(title=title)
                    if error_window.exists():
                        break
                except:
                    pass
        
            if not error_window or not error_window.exists():
                error_window = self.app.window(title_re='.*[Ee]rror.*')
        
            if error_window and error_window.exists():
                # Try to find OK button
                ok_button = None
                for btn_title in ["OK", "&OK", "Ok"]:
                    try:
                        ok_button = error_window.child_window(title=btn_title, class_name="Button")
                        if ok_button.exists():
                            break
                    except:
                        pass
                
                # If no specific button found, try any button
                if not ok_button or not ok_button.exists():
                    ok_button = error_window.child_window(class_name="Button")
                
                if ok_button and ok_button.exists():
                    ok_button.click_input()
                else:
                    # Fallback to Enter key
                    error_window.set_focus()
                    send_keys('{ENTER}')
        except Exception as e:
                print(f"Error with file error dialog: {e}")

        # Handle Call bases dialog with improved handling
        try:
            self.wait_for_dialog("call_bases")
            call_bases_window = self.app.window(title_re='Call bases.*')
            
            if call_bases_window and call_bases_window.exists():
                # Try to find Yes button
                yes_button = None
                for btn_title in ["&Yes", "Yes"]:
                    try:
                        yes_button = call_bases_window.child_window(title=btn_title, class_name="Button")
                        if yes_button.exists():
                            break
                    except:
                        pass
                
                if yes_button and yes_button.exists():
                    yes_button.click_input()
                else:
                    # Fallback to Enter key
                    call_bases_window.set_focus()
                    send_keys('{ENTER}')
        except Exception as e:
            print(f"Error with call bases dialog: {e}")
        
        # Wait for processing to complete with graceful timeout handling
        try:
            timings.wait_until(
                timeout=self.config.TIMEOUTS["process_completion"],
                retry_interval=0.2,
                func=lambda: self.is_process_complete(folder_path),
                value=True
            )
        except timings.TimeoutError:
            print(f"Warning: Timeout waiting for processing to complete for {folder_path}")
            print("This may be normal if the folder has already been processed or has special files")
        
        # Handle Low quality files skipped dialog if it appears
        try:
            for title in ["Low quality files skipped", "Quality files skipped"]:
                if self.app.window(title=title).exists():
                    low_quality_window = self.app.window(title=title)
                    ok_button = low_quality_window.child_window(class_name="Button")
                    ok_button.click_input()
                    break
        except Exception as e:
            print(f"Error handling quality files dialog: {e}")
        
        # Handle Read information dialog
        try:
            self.wait_for_dialog("read_info")
            if self.app.window(title_re='Read information for*').exists():
                read_window = self.app.window(title_re='Read information for*')
                read_window.close()
        except (timings.TimeoutError, Exception) as e:
            if isinstance(e, Exception):
                print(f"Error with read information dialog: {e}")
            else:
                print(f"Read information dialog did not appear for {folder_path}")
            # Continue processing
        
        return True
    
    def close(self):
        """Close the mSeq application"""
        if self.app:
            try:
                self.app.kill()
            except Exception as e:
                print(f"Error closing mSeq: {e}")
                # Try alternative approach
                if self.main_window and self.main_window.exists():
                    try:
                        self.main_window.close()
                    except:
                        pass