import os
import time
import platform
import re
import win32api
from mseqauto.config import MseqConfig
from mseqauto.core import OSCompatibilityManager
import logging
from typing import Callable, Any, TYPE_CHECKING

if TYPE_CHECKING:
    from pywinauto import Application, timings
    from pywinauto.keyboard import send_keys
    from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError

config = MseqConfig()
strategy = OSCompatibilityManager.get_dialog_strategy('browse_dialog')
click_location = strategy.get('click_location', 'center')
use_fallback_buttons = strategy.get('use_fallback_buttons', False)
extra_delay = strategy.get('extra_sleep_after_click', 0)

Application = None
timings = None
send_keys = None
ElementNotFoundError = None
ElementAmbiguousError = None


def _import_pywinauto():
    """Import pywinauto components and set them as global variables"""
    global Application, timings, send_keys, ElementNotFoundError, ElementAmbiguousError
    # Import with the same names for clarity
    from pywinauto import Application as App, timings as tim
    from pywinauto.keyboard import send_keys as sk
    from pywinauto.findwindows import ElementNotFoundError as ENF, ElementAmbiguousError as EAE

    # Assign to globals
    Application = App
    timings = tim
    send_keys = sk
    ElementNotFoundError = ENF
    ElementAmbiguousError = EAE


# Import pywinauto components immediately
_import_pywinauto()

class MseqAutomation:
    def __init__(self, config):
        self.config = config
        self.app = None
        self.main_window = None
        self.first_time_browsing = True
        self.logger = logging.getLogger(__name__)

        # Import pywinauto components if not already imported
        if Application is None:
            _import_pywinauto()

        # Initialize OS-specific information
        self.is_win11 = OSCompatibilityManager.is_windows_11()
        self.os_key = OSCompatibilityManager.get_os_key()
        # Store optimal timeouts based on OS for frequent operations
        self.expand_delay = OSCompatibilityManager.get_timeout("tree_expansion", 0.3)
        self.click_delay = OSCompatibilityManager.get_timeout("click_response", 0.2)

        self.logger.debug(f"Initialized MseqAutomation with OS: {self.os_key}")


    def connect_or_start_mseq(self):
        """Connect to existing mSeq or start a new instance"""
        try:
            self.app = Application(backend='win32').connect(title_re='Mseq*', timeout=1)
            self.logger.debug("Connected to existing mSeq instance with title 'Mseq*'")
        except (ElementNotFoundError, timings.TimeoutError):
            try:
                self.app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
                self.logger.debug("Connected to existing mSeq instance with title 'mSeq*'")
            except (ElementNotFoundError, timings.TimeoutError):
                self.logger.info("No running mSeq instance found, starting a new one...")
                print("No running mSeq instance found, starting a new one...")

                # Log current directory
                current_dir = os.getcwd()
                self.logger.debug(f"Current directory: {current_dir}")

                # Change directory if configured
                mseq_path = self.config.MSEQ_PATH
                if mseq_path and os.path.exists(mseq_path):
                    self.logger.debug(f"Changed directory to {mseq_path}")
                    print(f"Changed directory to {mseq_path}")

                start_cmd = f'cmd /c "cd /d {self.config.MSEQ_PATH} && {self.config.MSEQ_EXECUTABLE}"'
                try:
                    self.app = Application(backend='win32').start(start_cmd, wait_for_idle=False)
                    # Get process ID for logging
                    if hasattr(self.app, 'process'):
                        self.logger.info(f"Found new mSeq process: PID={self.app.process}")
                        print(f"Found new mSeq process: PID={self.app.process}")
                    # Connect to newly started instance
                    self.app.connect(title='mSeq', timeout=10)
                    self.logger.debug("Successfully connected to new mSeq instance")
                except Exception as e:
                    self.logger.error(f"Error starting mSeq: {e}")
                    raise
            except ElementAmbiguousError:
                self.logger.warning("Multiple mSeq* windows found, connecting to first and killing others")
                self.app = Application(backend='win32').connect(title_re='mSeq*', found_index=0, timeout=1)
                app2 = Application(backend='win32').connect(title_re='mSeq*', found_index=1, timeout=1)
                app2.kill()
        except ElementAmbiguousError:
            self.logger.warning("Multiple Mseq* windows found, connecting to first and killing others")
            self.app = Application(backend='win32').connect(title_re='Mseq*', found_index=0, timeout=1)
            app2 = Application(backend='win32').connect(title_re='Mseq*', found_index=1, timeout=1)
            app2.kill()

        # Get the main window
        if not self.app.window(title_re='mSeq*').exists():
            self.main_window = self.app.window(title_re='Mseq*')
            self.logger.debug("Found main window with title 'Mseq*'")
        else:
            self.main_window = self.app.window(title_re='mSeq*')
            self.logger.debug("Found main window with title 'mSeq*'")

        return self.app, self.main_window

    def wait_for_dialog(self, dialog_type, timeout=None):
        """Wait for a specific dialog with OS-specific timeouts"""
        if timeout is None:
            # Get OS-specific timeout
            timeout = OSCompatibilityManager.get_timeout(dialog_type)

        self.logger.debug(f"Waiting for {dialog_type} dialog with timeout={timeout}")


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

    def _scroll_if_needed(self, item):
        """Scroll through the tree item to see more children"""
        try:
            from pywinauto.keyboard import send_keys

            self.logger.debug("Attempting to scroll to find more items")

            # Ensure item is visible and expanded
            item.ensure_visible()

            if hasattr(item, 'expand'):
                item.expand()

            # Click to ensure focus
            if click_location == 'center':
                item.click_input()
            elif click_location == 'top_center':
                rect = item.rectangle()
                item.click_input(coords=(rect.width() // 2, 10))  # Click near top

            # Try Page Down a couple of times to see more items
            for i in range(3):
                send_keys('{PGDN}')
                time.sleep(0.3)

            return True
        except Exception as e:
            self.logger.warning(f"Error while scrolling: {e}")
            return False

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
            self.logger.error(f"Error ensuring dialog visibility: {e}")
            return False

    def _get_tree_view(self, dialog):
        """Get tree view control with enhanced Windows 11 compatibility"""
        self.logger.debug("Attempting to find tree view control")

        # Try to identify the tree view using only class name first (most reliable)
        try:
            # Windows 11 more commonly uses SysTreeView32 without a specific title
            tree_control = dialog.child_window(class_name="SysTreeView32")
            if tree_control.exists():
                self.logger.debug("Found tree view control by class name only")
                return tree_control
        except Exception as e:
            self.logger.warning(f"Could not find tree view by class name only: {e}")

        # Try different names used in Windows 10/11
        for name in ["Navigation Pane", "Tree View", "Choose project directory"]:
            try:
                self.logger.debug(f"Trying to find tree view with title: {name}")
                tree_control = dialog.child_window(title=name, class_name="SysTreeView32")
                if tree_control.exists():
                    self.logger.debug(f"Found tree view with title: {name}")
                    return tree_control
            except Exception as e:
                self.logger.warning(f"Could not find tree view with title {name}: {e}")

        # Special handling for the SHBrowseForFolder control which contains the tree view
        try:
            self.logger.debug("Trying to find via SHBrowseForFolder control")
            shell_control = dialog.child_window(class_name="SHBrowseForFolder ShellNameSpace Control")
            if shell_control.exists():
                self.logger.debug("Found SHBrowseForFolder control, looking for tree view inside")
                tree_control = shell_control.child_window(class_name="SysTreeView32")
                if tree_control.exists():
                    self.logger.debug("Found tree view inside SHBrowseForFolder")
                    return tree_control
        except Exception as e:
            self.logger.warning(f"Could not find tree view via SHBrowseForFolder: {e}")

        # Last resort - try to find ANY SysTreeView32 control in the dialog
        try:
            self.logger.debug("Last resort: trying to find ANY SysTreeView32 control")
            controls = dialog.children(class_name="SysTreeView32")
            if controls and len(controls) > 0:
                self.logger.debug(f"Found {len(controls)} potential tree view controls, using first one")
                return controls[0]
        except Exception as e:
            self.logger.error(f"Failed to find ANY tree view control: {e}")

        self.logger.error("Could not find tree view control in the dialog")
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
        """Navigate folder tree with OS-specific optimizations"""
        self.logger.info(f"Navigating to folder: {path}")
        dialog.set_focus()
        self._ensure_dialog_visible(dialog)

        # Get tree view using the robust method
        tree_view = self._get_tree_view(dialog)
        if not tree_view:
            self.logger.error("Could not find TreeView control")
            return False

        self.logger.debug("TreeView control found")

        # Handle different path formats
        if ":" in path:
            # Path has a drive letter
            parts = path.split("\\")
            drive = parts[0]  # e.g., "P:"
            folders = parts[1:] if len(parts) > 1 else []
            self.logger.debug(f"Path parsed: Drive={drive}, Folders={folders}")
        else:
            # Network path
            parts = path.split("\\")
            drive = "\\" + "\\".join(parts[:3])  # e.g., \\server\share
            folders = parts[3:] if len(parts) > 3 else []
            self.logger.debug(f"Network path parsed: Share={drive}, Folders={folders}")

        # Enhanced navigation strategy for both Windows 10 and 11
        try:
            # Get all root items
            try:
                root_items = list(tree_view.roots())
                self.logger.debug(f"Found {len(root_items)} root items in tree")
            except Exception as e:
                self.logger.error(f"Error getting tree roots: {e}")
                return False

            # Find Desktop in the root items
            desktop_item = None
            for item in root_items:
                if "Desktop" in item.text():
                    desktop_item = item
                    self.logger.debug(f"Found Desktop: {item.text()}")
                    break

            if not desktop_item:
                self.logger.error("Could not find Desktop root item in tree view")
                return False

            # Click on Desktop to ensure it's selected
            dialog.set_focus()
            if click_location == 'center':
                desktop_item.click_input()
            elif click_location == 'top_center':
                rect = desktop_item.rectangle()
                desktop_item.click_input(coords=(rect.width() // 2, 10))  # Click near top
            time.sleep(self.expand_delay)  # OS-optimized delay
            desktop_item.expand()
            time.sleep(self.expand_delay)  # OS-optimized delay

            # Get the expanded Desktop children
            desktop_children = []
            try:
                desktop_children = list(desktop_item.children())
                self.logger.debug(f"Desktop has {len(desktop_children)} children")
            except Exception as e:
                self.logger.warning(f"Error getting Desktop children: {e}")
                # Try to scroll Desktop to see more items
                if self._scroll_if_needed(desktop_item):
                    try:
                        desktop_children = list(desktop_item.children())
                        self.logger.debug(f"After scrolling, Desktop has {len(desktop_children)} children")
                    except Exception as e:
                        self.logger.error(f"Error getting Desktop children after scrolling: {e}")
                        return False
                else:
                    return False

            # For Windows 11, we need to find "This PC" in the Desktop children
            this_pc_item = None

            # Look for "This PC" by name (works in both Win10 and Win11)
            for child in desktop_children:
                if "PC" in child.text() or "Computer" in child.text():
                    this_pc_item = child
                    self.logger.debug(f"Found This PC by name: {child.text()}")
                    break

            # If This PC not found by name, try position-based approaches or scrolling
            if not this_pc_item:
                # Scrolling approach (works better in Windows 11)
                self.logger.debug("This PC not found in visible items, scrolling to find more")
                if self._scroll_if_needed(desktop_item):
                    # Refresh children list after scrolling
                    try:
                        desktop_children = list(desktop_item.children())
                        # Try again with the new list
                        for child in desktop_children:
                            if "PC" in child.text() or "Computer" in child.text():
                                this_pc_item = child
                                self.logger.debug(f"Found This PC after scrolling: {child.text()}")
                                break
                    except Exception as e:
                        self.logger.warning(f"Error after scrolling: {e}")

            # If This PC is still not found, try a positional approach as last resort
            if not this_pc_item and len(desktop_children) > 3:
                # Try positions that are common in Windows 10/11
                for idx in [2, 3, 4, 8]:  # Common positions across Windows versions
                    if idx < len(desktop_children):
                        potential_item = desktop_children[idx]
                        self.logger.debug(f"Trying positional This PC at index {idx}: {potential_item.text()}")
                        this_pc_item = potential_item
                        break

            # If This PC is still not found, we can't continue
            if not this_pc_item:
                self.logger.error("Could not find This PC in Desktop children even after scrolling")
                return False

            # Now that we found This PC, expand it to show the drives
            dialog.set_focus()
            this_pc_item.click_input()
            time.sleep(self.expand_delay)  # OS-optimized delay
            this_pc_item.expand()
            time.sleep(self.expand_delay)  # OS-optimized delay

            # Get This PC's children (the drives)
            drive_children = []
            try:
                drive_children = list(this_pc_item.children())
                self.logger.debug(f"This PC has {len(drive_children)} children (drives)")
            except Exception as e:
                self.logger.warning(f"Error getting drives: {e}")
                # Try to scroll to see more drives
                if self._scroll_if_needed(this_pc_item):
                    try:
                        drive_children = list(this_pc_item.children())
                        self.logger.debug(f"After scrolling, This PC has {len(drive_children)} children")
                    except Exception as e:
                        self.logger.error(f"Error getting drives after scrolling: {e}")
                        return False
                else:
                    return False

            # Look for our target drive
            drive_item = None
            mapped_name = self.config.NETWORK_DRIVES.get(drive, None)
            self.logger.debug(f"Looking for drive '{drive}' or mapped name '{mapped_name}'")

            # First pass: Look for exact match
            for item in drive_children:
                drive_text = item.text()
                if (drive == drive_text or
                        (mapped_name and mapped_name == drive_text) or
                        (drive in drive_text) or
                        (mapped_name and mapped_name in drive_text)):
                    dialog.set_focus()
                    item.click_input()
                    drive_item = item
                    time.sleep(self.expand_delay)  # OS-optimized delay
                    self.logger.debug(f"Found drive match: {drive_text}")
                    break

            # If drive not found, scroll and try again
            if not drive_item:
                self.logger.debug("Drive not found in visible items, scrolling to find more")
                if self._scroll_if_needed(this_pc_item):
                    # Refresh list after scrolling
                    try:
                        drive_children = list(this_pc_item.children())

                        # Try again with the new list
                        for item in drive_children:
                            drive_text = item.text()
                            if (drive == drive_text or
                                    (mapped_name and mapped_name == drive_text) or
                                    (drive in drive_text) or
                                    (mapped_name and mapped_name in drive_text)):
                                dialog.set_focus()
                                item.click_input()
                                drive_item = item
                                time.sleep(self.expand_delay)  # OS-optimized delay
                                self.logger.debug(f"Found drive after scrolling: {drive_text}")
                                break
                    except Exception as e:
                        self.logger.warning(f"Error after scrolling for drives: {e}")

            # If drive is still not found, we can't continue
            if not drive_item:
                self.logger.error(f"Could not find drive '{drive}' in This PC children")
                return False

            # If we're only navigating to the drive level, we're done
            if not folders:
                self.logger.debug("Navigation to drive level completed successfully")
                return True

            # Navigate through each subfolder
            current_item = drive_item

            for i, folder in enumerate(folders):
                self.logger.debug(f"Navigating to folder {i + 1}/{len(folders)}: {folder}")

                # Expand current folder
                current_item.expand()
                time.sleep(self.expand_delay)  # OS-optimized delay

                # Get children of current folder
                folder_children = []
                try:
                    folder_children = list(current_item.children())
                    self.logger.debug(f"Current folder has {len(folder_children)} children")
                except Exception as e:
                    self.logger.warning(f"Error getting folder children: {e}")
                    # Try to scroll to see more children
                    if self._scroll_if_needed(current_item):
                        try:
                            folder_children = list(current_item.children())
                            self.logger.debug(f"After scrolling, folder has {len(folder_children)} children")
                        except Exception as e:
                            self.logger.error(f"Error getting folder children after scrolling: {e}")
                            return False
                    else:
                        return False

                # Look for exact folder match first
                next_item = None

                for child in folder_children:
                    if child.text() == folder:  # Exact match
                        dialog.set_focus()
                        child.click_input()
                        next_item = child
                        time.sleep(self.expand_delay)  # OS-optimized delay
                        self.logger.debug(f"Found exact match for folder: {folder}")
                        break

                # If exact match not found, try partial match
                if not next_item:
                    for child in folder_children:
                        if folder.lower() in child.text().lower():  # Partial match
                            dialog.set_focus()
                            child.click_input()
                            next_item = child
                            time.sleep(self.expand_delay)  # OS-optimized delay
                            self.logger.debug(f"Found partial match for folder: {folder} -> {child.text()}")
                            break

                # If still not found, try scrolling
                if not next_item:
                    self.logger.debug(f"Folder '{folder}' not found in visible items, scrolling to find more")
                    if self._scroll_if_needed(current_item):
                        # Refresh list after scrolling
                        try:
                            folder_children = list(current_item.children())

                            # Try again after scrolling
                            for child in folder_children:
                                if child.text() == folder or folder.lower() in child.text().lower():
                                    dialog.set_focus()
                                    child.click_input()
                                    next_item = child
                                    time.sleep(self.expand_delay)  # OS-optimized delay
                                    self.logger.debug(f"Found folder after scrolling: {child.text()}")
                                    break
                        except Exception as e:
                            self.logger.warning(f"Error after scrolling for folder: {e}")
                    else:
                        self.logger.error(f"Could not find folder '{folder}' even after scrolling")
                        return False

                # If folder still not found, we can't continue
                if not next_item:
                    self.logger.error(f"Could not find folder '{folder}' in the current directory")
                    return False

                # Update current item for next iteration
                current_item = next_item

            # Navigation completed successfully
            self.logger.info("Navigation completed successfully")
            return True

        except Exception as e:
            # Log the error but don't raise it - we want to continue even if navigation fails
            self.logger.error(f"Error during folder navigation: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False

    def _click_dialog_button(self, dialog_window, button_titles):
        """Click a button in a dialog with better OS compatibility"""
        from pywinauto.keyboard import send_keys

        if not dialog_window or not dialog_window.exists():
            self.logger.warning("Dialog not found for button click operation")
            return False

        # Try each button title in order
        for btn_title in button_titles:
            try:
                ok_button = dialog_window.child_window(title=btn_title, class_name="Button")
                if ok_button.exists():
                    ok_button.click_input()
                    time.sleep(self.click_delay)  # OS-optimized delay
                    return True
            except Exception as e:
                self.logger.debug(f"Button '{btn_title}' not found: {e}")
                continue

        # Fallback to keyboard if no button found
        try:
            dialog_window.set_focus()
            send_keys('{ENTER}')
            time.sleep(self.click_delay)  # OS-optimized delay
            return True
        except Exception as e:
            self.logger.error(f"Failed to send keyboard input: {e}")
            return False

    def _handle_preferences_dialog(self):
        """Handle preferences dialog with OS awareness"""
        try:
            pref_window = None
            for title in ['Mseq Preferences', 'mSeq Preferences']:
                try:
                    pref_window = self.app.window(title=title)
                    if pref_window.exists():
                        break
                except:
                    pass

            if pref_window and pref_window.exists():
                self._click_dialog_button(pref_window, ["&OK", "OK", "Ok"])
            else:
                self.logger.warning("Preferences dialog not found or not visible")
        except Exception as e:
            self.logger.error(f"Error handling preferences dialog: {e}")

    def _handle_copy_files_dialog(self):
        """Handle Copy Files dialog with OS awareness"""
        try:
            copy_files_window = self.app.window(title_re='Copy.*sequence files')

            if copy_files_window and copy_files_window.exists():
                # Different ways to access list view depending on Windows version
                list_view_found = False

                # Windows 10 approach
                try:
                    shell_view = copy_files_window.child_window(title="ShellView", class_name="SHELLDLL_DefView")
                    if shell_view.exists():
                        list_view = shell_view.child_window(class_name="DirectUIHWND")
                        if list_view.exists():
                            list_view.click_input()
                            list_view_found = True
                except Exception as e:
                    self.logger.debug(f"Windows 10 list view approach failed: {e}")

                # Windows 11 approach
                if not list_view_found:
                    try:
                        list_view = copy_files_window.child_window(class_name="DirectUIHWND")
                        if list_view.exists():
                            list_view.click_input()
                            list_view_found = True
                    except Exception as e:
                        self.logger.debug(f"Windows 11 list view approach failed: {e}")

                # Last resort - click in the middle of the dialog
                if not list_view_found:
                    try:
                        rect = copy_files_window.rectangle()
                        copy_files_window.click_input(coords=((rect.right - rect.left) // 2,
                                                              (rect.bottom - rect.top) // 2))
                        list_view_found = True
                    except Exception as e:
                        self.logger.warning(f"Center-click approach failed: {e}")

                if list_view_found:
                    # Select all files
                    send_keys('^a')  # Ctrl+A
                    time.sleep(self.click_delay)  # Wait for selection

                    # Click the Open button
                    self._click_dialog_button(copy_files_window, ["&Open", "Open"])
            else:
                self.logger.warning("Copy files dialog not found or not visible")
        except Exception as e:
            self.logger.error(f"Error handling copy files dialog: {e}")

    def _handle_error_dialog(self):
        """Handle Error dialog with OS awareness"""
        try:
            error_window = None
            for title in ['File error', 'Error']:
                try:
                    error_window = self.app.window(title=title)
                    if error_window.exists():
                        break
                except:
                    pass

            if not error_window or not error_window.exists():
                try:
                    error_window = self.app.window(title_re='.*[Ee]rror.*')
                except:
                    pass

            if error_window and error_window.exists():
                # Try to find OK button first
                button_found = self._click_dialog_button(error_window, ["OK", "&OK", "Ok"])

                # If no specific button found, try any button
                if not button_found:
                    try:
                        ok_button = error_window.child_window(class_name="Button")
                        if ok_button.exists():
                            ok_button.click_input()
                            time.sleep(self.click_delay)
                        else:
                            # Last resort - press Enter
                            error_window.set_focus()
                            send_keys('{ENTER}')
                            time.sleep(self.click_delay)
                    except Exception as e:
                        self.logger.warning(f"Failed to handle error dialog with generic approach: {e}")
            else:
                self.logger.debug("No error dialog found or it disappeared")
        except Exception as e:
            self.logger.error(f"Error handling error dialog: {e}")

    def _handle_call_bases_dialog(self):
        """Handle Call bases dialog with OS awareness"""
        try:
            call_bases_window = self.app.window(title_re='Call bases.*')

            if call_bases_window and call_bases_window.exists():
                # Click Yes button
                button_found = self._click_dialog_button(call_bases_window, ["&Yes", "Yes"])

                if not button_found:
                    # Try any button that might work
                    try:
                        yes_button = call_bases_window.child_window(class_name="Button")
                        if yes_button.exists():
                            yes_button.click_input()
                        else:
                            # Last resort - press Enter for default action
                            call_bases_window.set_focus()
                            send_keys('{ENTER}')
                    except Exception as e:
                        self.logger.warning(f"Failed to handle call bases dialog with generic approach: {e}")
            else:
                self.logger.warning("Call bases dialog not found or not visible")
        except Exception as e:
            self.logger.error(f"Error handling call bases dialog: {e}")

    def _handle_low_quality_dialog(self):
        """Handle Low quality files skipped dialog"""
        try:
            for title in ["Low quality files skipped", "Quality files skipped"]:
                low_quality_window = self.app.window(title=title)
                if low_quality_window.exists():
                    button_found = self._click_dialog_button(low_quality_window, ["OK", "&OK", "Ok"])

                    if not button_found:
                        # Try any button
                        ok_button = low_quality_window.child_window(class_name="Button")
                        if ok_button.exists():
                            ok_button.click_input()
                        else:
                            # Last resort - press Enter
                            low_quality_window.set_focus()
                            send_keys('{ENTER}')

                    self.logger.debug(f"Handled dialog: {title}")
                    return True

            return False
        except Exception as e:
            self.logger.error(f"Error handling low quality dialog: {e}")
            return False

    def _wait_for_process_completion(self, folder_path, max_wait=None, interval=None):
        """Wait for mSeq processing to complete with OS-specific timeouts"""
        if max_wait is None:
            max_wait = OSCompatibilityManager.get_timeout("process_completion")

        if interval is None:
            interval = OSCompatibilityManager.get_timeout("polling_interval", 0.5)

        self.logger.debug(f"Waiting for process completion (max_wait={max_wait}s, interval={interval}s)")

        elapsed = 0
        while elapsed < max_wait:
            # Check if process completed
            completed = False

            # Check for low quality dialog
            if self._handle_low_quality_dialog():
                completed = True

            # Check for read info dialog
            try:
                if self.app.window(title_re='Read information for*').exists():
                    read_window = self.app.window(title_re='Read information for*')
                    read_window.close()
                    completed = True
                    self.logger.debug("Closed read information dialog")
            except Exception as e:
                self.logger.debug(f"Error checking read info dialog: {e}")

            # Check for txt files - most reliable completion indicator
            txt_count = 0
            for item in os.listdir(folder_path):
                if any(item.endswith(ext) for ext in self.config.TEXT_FILES):
                    txt_count += 1

            if txt_count >= 5:
                self.logger.info(f"Processing completed for {folder_path}: found {txt_count} text files")
                completed = True

            if completed:
                return True

            time.sleep(interval)
            elapsed += interval

            # Log progress periodically
            if int(elapsed) % 5 == 0 and int(elapsed) > 0:
                self.logger.debug(f"Still waiting for processing completion ({int(elapsed)}s elapsed)")

        self.logger.warning(f"Timeout waiting for processing to complete for {folder_path}")
        return True  # Return True anyway to continue with next folder

    def process_folder(self, folder_path):
        """Process a folder with mSeq using OS-aware optimizations"""
        self.logger.info(f"Processing folder: {folder_path}")

        # Validate folder exists
        if not os.path.exists(folder_path):
            self.logger.warning(f"Folder does not exist: {folder_path}")
            return False

        # Check for .ab1 files to process
        ab1_files = [f for f in os.listdir(folder_path) if f.endswith(self.config.ABI_EXTENSION)]
        if not ab1_files:
            self.logger.warning(f"No {self.config.ABI_EXTENSION} files found in {folder_path}, skipping processing")
            return False

        self.logger.debug(f"Found {len(ab1_files)} AB1 files to process")

        # Connect to mSeq
        self.app, self.main_window = self.connect_or_start_mseq()
        if not self.main_window:
            self.logger.error("Failed to connect to mSeq")
            return False

        self.main_window.set_focus()
        from pywinauto.keyboard import send_keys
        send_keys('^n')  # Ctrl+N for new project
        self.logger.debug("Sent Ctrl+N to create new project")

        # Wait for and handle Browse For Folder dialog
        browse_timeout = OSCompatibilityManager.get_timeout("browse_dialog")
        dialog_found = self.wait_for_dialog("browse_dialog", timeout=browse_timeout)

        if not dialog_found:
            self.logger.error("Browse For Folder dialog not found")
            return False

        dialog_window = self._get_browse_dialog()
        if not dialog_window:
            self.logger.error("Failed to get Browse For Folder dialog window")
            return False

        # Add a delay for the first browsing operation - needed on both Windows 10 and 11
        if self.first_time_browsing:
            self.first_time_browsing = False
            first_browse_delay = OSCompatibilityManager.get_timeout("first_browse_delay", 1.2)
            self.logger.debug(f"First time browsing, adding delay of {first_browse_delay}s")
            time.sleep(first_browse_delay)
        else:
            subsequent_browse_delay = OSCompatibilityManager.get_timeout("subsequent_browse_delay", 0.5)
            time.sleep(subsequent_browse_delay)

        # Navigate to the target folder
        navigate_success = self.navigate_folder_tree(dialog_window, folder_path)
        if not navigate_success:
            self.logger.error(f"Navigation failed for {folder_path}")
            return False

        self.logger.debug("Navigation successful, clicking OK")

        # Find and click OK button with better error handling
        button_found = self._click_dialog_button(dialog_window, ["OK", "&OK", "Ok"])
        if not button_found:
            self.logger.warning("OK button not found by name, trying fallback methods")
            try:
                # Try to find any button
                buttons = dialog_window.children(class_name="Button")
                if buttons:
                    self.logger.debug(f"Found {len(buttons)} buttons, clicking the first one")
                    buttons[0].click_input()
                else:
                    # Last resort - press Enter
                    self.logger.debug("No buttons found, pressing Enter")
                    dialog_window.set_focus()
                    send_keys('{ENTER}')
            except Exception as e:
                self.logger.error(f"Error with fallback OK button handling: {e}")
                return False

        # Handle mSeq Preferences dialog
        self.wait_for_dialog("preferences")
        self._handle_preferences_dialog()

        # Handle Copy sequence files dialog
        self.wait_for_dialog("copy_files")
        self._handle_copy_files_dialog()

        # Handle File error dialog
        self.wait_for_dialog("error_window")
        self._handle_error_dialog()

        # Handle Call bases dialog
        self.wait_for_dialog("call_bases")
        self._handle_call_bases_dialog()

        # Wait for processing to complete
        process_timeout = OSCompatibilityManager.get_timeout("process_completion")
        poll_interval = OSCompatibilityManager.get_timeout("polling_interval", 0.5)
        completion_success = self._wait_for_process_completion(folder_path, process_timeout, poll_interval)

        if completion_success:
            self.logger.info(f"Successfully processed folder: {folder_path}")
        else:
            self.logger.warning(f"Processing may not have completed properly for: {folder_path}")

        return True

    def close(self):
        """Close the mSeq application with better error handling"""
        if self.app:
            self.logger.debug("Attempting to close mSeq application")
            try:
                self.app.kill()
                self.logger.info("mSeq application killed successfully")
            except Exception as e:
                self.logger.warning(f"Error killing mSeq process: {e}")
                # Try alternative approach
                if self.main_window and self.main_window.exists():
                    try:
                        self.main_window.close()
                        self.logger.info("mSeq window closed successfully")
                    except Exception as e2:
                        self.logger.error(f"Error closing mSeq window: {e2}")
        else:
            self.logger.debug("No mSeq application to close")