import os
import time
import platform
import re
import win32api
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
print(sys.path)
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
    def __init__(self, config, use_fast_navigation=True, log_level=None):
        self.config = config
        self.app = None
        self.main_window = None
        self.first_time_browsing = True
        self.logger = logging.getLogger(__name__)
        self.use_fast_navigation = use_fast_navigation
        if log_level is not None:
            self.logger.setLevel(log_level)

        # Import pywinauto components if not already imported
        if Application is None:
            _import_pywinauto()

        # Initialize OS-specific information
        self.is_win11 = OSCompatibilityManager.is_windows_11()
        self.os_key = OSCompatibilityManager.get_os_key()
        # Store optimal timeouts based on OS for frequent operations
        self.expand_delay = OSCompatibilityManager.get_timeout("tree_expansion", 0.3)
        self.click_delay = OSCompatibilityManager.get_timeout("click_response", 0.2)

        self.logger.info(f"Initialized MseqAutomation with OS: {self.os_key}")
        mseqinit = f"mSeqAutomation initialized with OS: {self.os_key}, Win11: {self.is_win11}"
        print(mseqinit)  # Print to console for immediate feedback
        

    def connect_or_start_mseq(self):
        """Connect to existing mSeq or start a new instance"""
        try:
            self.app = Application(backend='win32').connect(title_re='Mseq*', timeout=1)
            self.logger.info("Connected to existing mSeq instance with title 'Mseq*'")
            print("Connected to existing mSeq instance with title 'Mseq*'")
            # Check if the main window is available
        except (ElementNotFoundError, timings.TimeoutError):
            try:
                self.app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
                self.logger.info("Connected to existing mSeq instance with title 'mSeq*'") 
                print("Connected to existing mSeq instance with title 'mSeq*'")
            except (ElementNotFoundError, timings.TimeoutError):
                self.logger.info("No running mSeq instance found, starting a new one...")
                print("No running mSeq instance found, starting a new one...")

                # Log current directory
                current_dir = os.getcwd()
                self.logger.info(f"Current directory: {current_dir}")
                print(f"Current directory: {current_dir}")

                # Change directory if configured
                mseq_path = self.config.MSEQ_PATH
                if mseq_path and os.path.exists(mseq_path):
                    self.logger.info(f"Changed directory to {mseq_path}")
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
                    self.logger.info("Successfully connected to new mSeq instance")
                    print("Successfully connected to new mSeq instance")
                    # Allow some time for the application to initialize properly
                except Exception as e:
                    self.logger.error(f"Error starting mSeq: {e}")
                    print(f"Error starting mSeq: {e}")
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
            self.logger.info("Found main window with title capital 'Mseq*'")
            print("Found main window with title capital 'Mseq*'")
        else:
            self.main_window = self.app.window(title_re='mSeq*')
            self.logger.info("Found main window with title 'mSeq*'")
            print("Found main window with title 'mSeq*'")
        return self.app, self.main_window

    def wait_for_dialog(self, dialog_type, timeout=None):
        """Wait for a specific dialog with OS-specific timeouts"""
        if timeout is None:
            # Get OS-specific timeout
            timeout = OSCompatibilityManager.get_timeout(dialog_type)

        self.logger.info(f"Waiting for {dialog_type} dialog with timeout={timeout}")
        print(f"Waiting for {dialog_type} dialog with timeout={timeout}")


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
        """Enhanced scroll method for Windows 11 compatibility"""
        try:
            from pywinauto.keyboard import send_keys

            self.logger.info(f"Attempting to scroll to find more items in '{item.text()}'")
            print(f"Attempting to scroll to find more items in '{item.text()}'")

            # Ensure item is visible and expanded
            item.ensure_visible()
            
            # Make sure focus is properly set before expanding
            dialog = item.parent().parent()  # Get dialog window
            dialog.set_focus()

            # Click to ensure focus with robust method
            if click_location == 'center':
                item.click_input()
            elif click_location == 'top_center':
                rect = item.rectangle()
                item.click_input(coords=(rect.width() // 2, 10))  # Click near top
            print(f"Clicked on item with {click_location} {item.text()} to ensure focus")
            # Short delay after clicking
            time.sleep(self.click_delay)
            
            # Try expanding multiple times in different ways (helps with Windows 11)
            try:
                if hasattr(item, 'expand'):
                    item.expand()
                    time.sleep(self.expand_delay)
                    
                    # Sometimes Windows 11 needs additional expand commands
                    item.click_input()
                    send_keys('{RIGHT}')  # Right arrow to expand
                    print(f"Clicked and sent RIGHT key to expand '{item.text()}'")
            except Exception as expand_e:
                self.logger.warning(f"Error during expand operation: {expand_e}")
                print(f"Error during expand operation: {expand_e}")
        except Exception as e:
            self.logger.error(f"Error while scrolling: {e}")
            print(f"Error while scrolling: {e}")   
            time.sleep(self.expand_delay)
        except Exception as expand_e:
            self.logger.info(f"Expand operation failed: {expand_e}")
            print(f"Expand operation failed: {expand_e}")

            # More aggressive scrolling for Windows 11
            send_keys('{DOWN 5}')  # Scroll down 5 items
            time.sleep(self.expand_delay)
            print("Scrolled down 5 items using DOWN key")
            # Try Page Down a few times to see more items
            for i in range(3):
                send_keys('{PGDN}')
                time.sleep(self.expand_delay)
                print(f"Scrolled down with PgDn {i+1} time(s)")
                # Try to check if we have more items now
                try:
                    children_count = len(list(item.children()))
                    self.logger.info(f"After PgDn {i+1}, found {children_count} children")
                    print(f"After PgDn {i+1}, found {children_count} children")
                    # If we found a reasonable number of items, we can stop
                    if children_count > 3:
                        break
                except:
                    pass

            # Windows 11 sometimes needs Home/End keys
            send_keys('{END}')  # Go to end of list
            time.sleep(self.expand_delay)
            print("Scrolled to the end of the list using END key")
            send_keys('{HOME}')  # Go back to top
            time.sleep(self.expand_delay)
            print("Scrolled back to the top using HOME key")

            return True
        except Exception as e:
            self.logger.warning(f"Error while scrolling: {e}")
            print(f"Error while scrolling: {e}")
            return False

    def _ensure_dialog_visible(self, dialog):
        """Make sure dialog is visible and positioned correctly"""
        try:
            # Check if dialog exists and is positioned properly
            if dialog.exists() and dialog.rectangle().width() > 0:
                # Dialog is visible, no action needed
                print(f"Dialog '{dialog.window_text()}' is already visible and positioned correctly")
                return True

            # Try to set focus and move if needed
            dialog.set_focus()
            print(f"Setting focus to dialog '{dialog.window_text()}'")

            # Get screen dimensions
            try:
                screen_width = win32api.GetSystemMetrics(0)
                screen_height = win32api.GetSystemMetrics(1)

                # Reposition dialog if it's offscreen
                rect = dialog.rectangle()
                if rect.left < 0 or rect.top < 0 or rect.right > screen_width or rect.bottom > screen_height:
                    dialog.move_window(100, 100, rect.width(), rect.height())
                    print(f"Dialog '{dialog.window_text()}' repositioned to (100, 100) if it was offscreen")
            except ImportError:
                # Fall back if win32api isn't available
                dialog.set_focus()
                print(f"Dialog '{dialog.window_text()}' focus set, but cannot reposition due to win32api import error")
            return True
        except Exception as e:
            self.logger.error(f"Error ensuring dialog visibility: {e}")
            print(f"Error ensuring dialog visibility: {e}")
            return False

    def _get_tree_view(self, dialog):
        """Get tree view control with enhanced Windows 11 compatibility"""
        self.logger.info("Attempting to find tree view control")
        print("Attempting to find tree view control")

        # Try to identify the tree view using only class name first (most reliable)
        try:
            # Windows 11 more commonly uses SysTreeView32 without a specific title
            tree_control = dialog.child_window(class_name="SysTreeView32")
            if tree_control.exists():
                self.logger.info("Found tree view control by class name only")
                print("Found tree view control by class name only")
                return tree_control
        except Exception as e:
            self.logger.warning(f"Could not find tree view by class name only: {e}")
            print(f"Could not find tree view by class name only: {e}") 
        # Try different names used in Windows 10/11
        for name in ["Navigation Pane", "Tree View", "Choose project directory"]:
            try:
                self.logger.info(f"Trying to find tree view with title: {name}")
                print(f"Trying to find tree view with title: {name}")
                tree_control = dialog.child_window(title=name, class_name="SysTreeView32")
                if tree_control.exists():
                    self.logger.info(f"Found tree view with title: {name}")
                    print(f"Found tree view with title: {name}")
                    return tree_control
            except Exception as e:
                self.logger.warning(f"Could not find tree view with title {name}: {e}")
                print(f"Could not find tree view with title {name}: {e}")
        # Special handling for the SHBrowseForFolder control which contains the tree view
        try:
            self.logger.info("Trying to find via SHBrowseForFolder control")
            print("Trying to find via SHBrowseForFolder control")
            shell_control = dialog.child_window(class_name="SHBrowseForFolder ShellNameSpace Control")
            if shell_control.exists():
                self.logger.info("Found SHBrowseForFolder control, looking for tree view inside")
                print("Found SHBrowseForFolder control, looking for tree view inside")
                tree_control = shell_control.child_window(class_name="SysTreeView32")
                if tree_control.exists():
                    self.logger.info("Found tree view inside SHBrowseForFolder")
                    print("Found tree view inside SHBrowseForFolder")
                    return tree_control
        except Exception as e:
            self.logger.warning(f"Could not find tree view via SHBrowseForFolder: {e}")
            print(f"Could not find tree view via SHBrowseForFolder: {e}")
        # Last resort - try to find ANY SysTreeView32 control in the dialog
        try:
            self.logger.info("Last resort: trying to find ANY SysTreeView32 control")
            print("Last resort: trying to find ANY SysTreeView32 control")
            controls = dialog.children(class_name="SysTreeView32")
            if controls and len(controls) > 0:
                self.logger.info(f"Found {len(controls)} potential tree view controls, using first one")
                print(f"Found {len(controls)} potential tree view controls, using first one")
                print(f"First tree view control found: '{controls[0].window_text()}'")
                return controls[0]
        except Exception as e:
            self.logger.error(f"Failed to find ANY tree view control: {e}")
            print(f"Failed to find ANY tree view control: {e}")

        self.logger.error("Could not find tree view control in the dialog")
        print("Could not find tree view control in the dialog")
        return None

    def _get_this_pc_item(self, desktop_item, success_paths=None):
        """Get This PC node with enhanced Windows 11 compatibility"""
        self.logger.info("Attempting to find This PC item in desktop children")
        print("Attempting to find This PC item in desktop children")

        found_items = []
        try:
            all_tree_items = desktop_item.wrapper_object().children()
            self.logger.info(f"TRACKING: Found {len(all_tree_items)} total tree items to search")
            print(f"TRACKING: Found {len(all_tree_items)} total tree items to search")
            
            for item in all_tree_items:
                try:
                    item_text = item.text().lower()
                    # Only log items that might be relevant
                    if "pc" in item_text or "computer" in item_text:
                        self.logger.info(f"TRACKING: Found potential PC match: '{item.text()}'")
                        print(f"TRACKING: Found potential PC match: '{item.text()}'")
                        
                    if item_text == "this pc":
                        if success_paths is not None:
                            success_paths['this_pc_method'] = 'direct_search'
                        self.logger.info(f"SUCCESS PATH: Found exact This PC match via direct search: '{item.text()}'")
                        print(f"SUCCESS PATH: Found exact This PC match via direct search: '{item.text()}'")
                        return item
                except:
                    continue
        except Exception as e:
            self.logger.info(f"TRACKING: Direct search approach failed: {e}")
            print(f"TRACKING: Direct search approach failed: {e}")
        
        # If we found exact "This PC" matches, return the first one
        if found_items:
            self.logger.info(f"Returning exact match for This PC from {len(found_items)} found items")
            print(f"Returning exact match for This PC from {len(found_items)} found items")
            print(f"Exact matches found: {[item.text() for item in found_items]}")
            print(f"First exact match: '{found_items[0].text()}'")
            return found_items[0]
        
        # Traditional desktop children approach
        self.logger.info("TRACKING: Trying desktop children approach")
        print("TRACKING: Trying desktop children approach") 
        for child in desktop_item.children():
            child_text = child.text().lower()
            if any(pc_name.lower() in child_text for pc_name in ["this pc", "computer", "pc"]):
                if success_paths is not None:
                    success_paths['this_pc_method'] = 'desktop_children'
                self.logger.info(f"SUCCESS PATH: Found This PC via desktop children approach: '{child.text()}'")
                print(f"SUCCESS PATH: Found This PC via desktop children approach: '{child.text()}'")
                return child

        
        # Scrolling approach
        self.logger.info("TRACKING: Trying scrolling approach")
        print("TRACKING: Trying scrolling approach")
        self._scroll_if_needed(desktop_item)
        
        for child in desktop_item.children():
            child_text = child.text().lower()
            if any(pc_name.lower() in child_text for pc_name in ["this pc", "computer", "pc"]):
                self.logger.info(f"SUCCESS PATH: Found This PC after scrolling: '{child.text()}'")
                print(f"SUCCESS PATH: Found This PC after scrolling: '{child.text()}'")
                return child
    
        
        # Position-based approach
        self.logger.info("TRACKING: Trying position-based approach")
        print("TRACKING: Trying position-based approach")
        try:
            children = list(desktop_item.children())
            self.logger.info(f"TRACKING: Total immediate children of Desktop: {len(children)}")
            print(f"TRACKING: Total immediate children of Desktop: {len(children)}")

            # Log position range where This PC might be
            position_range = range(20, 26) if len(children) >= 26 else range(len(children))
            for pos in position_range:
                item = children[pos]
                self.logger.info(f"TRACKING: Position {pos} contains: '{item.text()}'")
                print(f"TRACKING: Position {pos} contains: '{item.text()}'")

            # Priority positions from screenshot
            priority_positions = [23, 24, 22, 25, 20, 21]
            for pos in priority_positions:
                if pos < len(children):
                    item = children[pos]
                    if "pc" in item.text().lower():
                        if success_paths is not None:
                            success_paths['this_pc_method'] = f'position_{pos}_text_match'
                        self.logger.info(f"SUCCESS PATH: Found This PC at position {pos}: '{item.text()}'")
                        print(f"SUCCESS PATH: Found This PC at position {pos}: '{item.text()}'")
                        return item
                    
                    # Even if text doesn't match, try known positions
                    if pos == 23 or pos == 24:
                        if success_paths is not None:
                            success_paths['this_pc_method'] = f'position_{pos}_known_layout'
                        self.logger.info(f"SUCCESS PATH: Using position {pos} as This PC based on known layout: '{item.text()}'")
                        print(f"SUCCESS PATH: Using position {pos} as This PC based on known layout: '{item.text()}'")
                        return item
        except Exception as e:
            self.logger.info(f"TRACKING: Position-based approach failed: {e}")
            print(f"TRACKING: Position-based approach failed: {e}")

        self.logger.error("TRACKING: All approaches to find This PC failed")
        print("TRACKING: All approaches to find This PC failed")
        return None
    
    def _find_drive_item(self, this_pc_item, target_drive, dialog=None, success_paths=None):
        """Find a specific drive with enhanced Windows 11 compatibility
        
        Args:
            this_pc_item: The This PC tree item
            target_drive: The drive to find (e.g., "C:" or "P:")
            dialog: Optional dialog window (if already known)
            
        Returns:
            The drive tree item or None if not found
        """
        self.logger.info(f"=== NAVIGATION TRACKING: Looking for drive: {target_drive} ===")
        print(f"=== NAVIGATION TRACKING: Looking for drive: {target_drive} ===")
        
        # Get the mapped name if it exists
        mapped_name = self.config.NETWORK_DRIVES.get(target_drive, None)
        self.logger.info(f"TRACKING: Target drive: '{target_drive}', Mapped name: '{mapped_name}'")
        print(f"TRACKING: Target drive: '{target_drive}', Mapped name: '{mapped_name}'")
    
        
        # Make sure This PC is expanded - safely get dialog without using parent()
        if dialog is None:
            try:
                # First try to find Browse For Folder dialog by title
                dialog = self._get_browse_dialog()
                if dialog and dialog.exists():
                    self.logger.info("Found Browse dialog by title")
                    print("Found Browse dialog by title")
                else:
                    # Try to find any active dialog from the application
                    dialogs = []
                    for title in ['Browse For Folder', 'Browse for Folder', 'Select Folder']:
                        try:
                            dlg = self.app.window(title=title)
                            if dlg.exists():
                                dialogs.append(dlg)
                        except:
                            pass
                    
                    # If we found any dialog, use the first one
                    if dialogs:
                        dialog = dialogs[0]
                        self.logger.info(f"Found dialog by title: {dialog.window_text()}")
                        print(f"Found dialog by title: {dialog.window_text()}")
                    else:
                        # Last resort - use top window
                        try:
                            dialog = self.app.top_window()
                            self.logger.info(f"Using top window as dialog: {dialog.window_text()}")
                            print(f"Using top window as dialog: {dialog.window_text()}")
                        except:
                            self.logger.warning("Could not find any dialog window")
                            print("Could not find any dialog window")
            except Exception as e:
                self.logger.warning(f"Error finding dialog: {e}")
                print(f"Error finding dialog: {e}")

        # Focus and expand This PC item
        try:
            this_pc_item.click_input()
            time.sleep(self.expand_delay * 2)  # Double delay for Windows 11
            this_pc_item.expand()
            time.sleep(self.expand_delay * 2)  # Double delay for Windows 11
        except Exception as e:
            self.logger.warning(f"Error focusing/expanding This PC: {e}")
            print(f"Error focusing/expanding This PC: {e}")
            # Try alternative expansion approach
            try:
                from pywinauto.keyboard import send_keys
                this_pc_item.click_input()
                print(f"Clicked on This PC item '{this_pc_item.text()}' to ensure focus")
                send_keys('{RIGHT}')  # Right arrow to expand
                print(f"Sent RIGHT key to expand '{this_pc_item.text()}'")
                time.sleep(self.expand_delay * 2)
            except Exception as e2:
                self.logger.warning(f"Alternative expansion failed: {e2}")
                print(f"Alternative expansion failed: {e2}")
                
        # Try to get all children of This PC
        drive_children = []
        try:
            drive_children = list(this_pc_item.children())
            self.logger.info(f"TRACKING: Found {len(drive_children)} children of This PC")
            print(f"TRACKING: Found {len(drive_children)} children of This PC")
            
            # Log all drive children
            for i, child in enumerate(drive_children):
                self.logger.info(f"TRACKING: Drive child {i}: '{child.text()}'")
                print(f"TRACKING: Drive child {i}: '{child.text()}'")
        except Exception as e:
            self.logger.warning(f"TRACKING: Error getting children of This PC: {e}")
            print(f"TRACKING: Error getting children of This PC: {e}")
            # Scrolling fallback code...
            if self._scroll_if_needed(this_pc_item):
                try:
                    drive_children = list(this_pc_item.children())
                    self.logger.info(f"After scrolling, found {len(drive_children)} children")
                    print(f"After scrolling, found {len(drive_children)} children")
                    # Log all drive children after scrolling
                except Exception as scroll_e:
                    self.logger.error(f"Error getting children after scrolling: {scroll_e}")
                    print(f"Error getting children after scrolling: {scroll_e}")
                    # If still no children found, we can't proceed
                    return None
        
        # Windows 11 specific fix: Try clicking on This PC again and expanding
        if len(drive_children) < 2:  # If we didn't find enough drives
            self.logger.info("Not enough drives found, trying Windows 11 specific fix")
            print("Not enough drives found, trying Windows 11 specific fix")
            # Ensure dialog is focused again
            try:
                # Click again and wait
                dialog.set_focus()
                this_pc_item.click_input()
                print(f"Clicked on This PC item '{this_pc_item.text()}' to ensure focus again")
                time.sleep(self.expand_delay * 3)
                
                # Try expanding multiple times with different methods
                this_pc_item.expand()
                print(f"Expanded This PC item '{this_pc_item.text()}' again")
                time.sleep(self.expand_delay)
                
                # Send key to expand (Windows 11 sometimes requires this)
                from pywinauto.keyboard import send_keys
                this_pc_item.click_input()
                send_keys('{RIGHT}')  # Right arrow to expand
                print(f"Sent RIGHT key to expand '{this_pc_item.text()}' again")
                time.sleep(self.expand_delay * 2)
                
                # Try to get children again
                try:
                    drive_children = list(this_pc_item.children())
                    self.logger.info(f"After Win11 fix, found {len(drive_children)} children")
                    print(f"After Win11 fix, found {len(drive_children)} children")
                except Exception as e:
                    self.logger.warning(f"Still couldn't get children after Win11 fix: {e}")
                    print(f"Still couldn't get children after Win11 fix: {e}")
            except Exception as e:
                self.logger.warning(f"Windows 11 specific fix failed: {e}")
                print(f"Windows 11 specific fix failed: {e}")
        
        # Look for drive matches with detailed logging
        for item in drive_children:
            drive_text = item.text()
            
            # Log match attempts more clearly
            exact_match = drive_text == target_drive
            mapped_exact_match = mapped_name and drive_text == mapped_name
            contains_match = target_drive in drive_text
            mapped_contains_match = mapped_name and mapped_name in drive_text

            self.logger.info(f"Checking drive '{drive_text}': exact_match={exact_match}, mapped_exact_match={mapped_exact_match}, contains_match={contains_match}, mapped_contains_match={mapped_contains_match}")

            print(f"Checking drive '{drive_text}': exact_match={exact_match}, mapped_exact_match={mapped_exact_match}, contains_match={contains_match}, mapped_contains_match={mapped_contains_match}")

            # Extract drive letter for Win11 style match
            win11_style_match = False
            if ":" in drive_text:
                drive_letter_parts = re.findall(r'([A-Za-z]:)', drive_text)
                print(f"Extracted drive letter parts: {drive_letter_parts}")
                self.logger.info(f"Extracted drive letter parts: {drive_letter_parts}")
                if drive_letter_parts and drive_letter_parts[0].upper() == target_drive.upper():
                    win11_style_match = True
                    self.logger.info(f"TRACKING: Windows 11 style match found in '{drive_text}'")
                    print(f"TRACKING: Windows 11 style match found in '{drive_text}'")
            
            # Log which match condition succeeded
            if exact_match:
                if success_paths is not None:
                    success_paths['drive_method'] = 'exact_match'
                self.logger.info(f"SUCCESS PATH: Found drive via exact match: '{drive_text}'")
                print(f"SUCCESS PATH: Found drive via exact match: '{drive_text}'")
                return item
            elif mapped_exact_match:
                if success_paths is not None:
                    success_paths['drive_method'] = 'mapped_exact_match'
                self.logger.info(f"SUCCESS PATH: Found drive via mapped exact match: '{drive_text}'")
                print(f"SUCCESS PATH: Found drive via mapped exact match: '{drive_text}'")
                return item
            elif contains_match:
                if success_paths is not None:
                    success_paths['drive_method'] = 'contains_match'
                self.logger.info(f"SUCCESS PATH: Found drive via contains match: '{drive_text}'")
                print(f"SUCCESS PATH: Found drive via contains match: '{drive_text}'")
                return item
            elif mapped_contains_match:
                if success_paths is not None:
                    success_paths['drive_method'] = 'mapped_contains_match'
                self.logger.info(f"SUCCESS PATH: Found drive via mapped contains match: '{drive_text}'")
                print(f"SUCCESS PATH: Found drive via mapped contains match: '{drive_text}'")
                return item
            elif win11_style_match:
                if success_paths is not None:
                    success_paths['drive_method'] = 'win11_style_match'
                self.logger.info(f"SUCCESS PATH: Found drive via Windows 11 style match: '{drive_text}'")
                print(f"SUCCESS PATH: Found drive via Windows 11 style match: '{drive_text}'")
                return item
        
        # If we didn't find a match, try scrolling to see more
        self.logger.info("Drive not found in visible items, scrolling to find more")
        print("Drive not found in visible items, scrolling to find more")
        if self._scroll_if_needed(this_pc_item):
            try:
                # Refresh list after scrolling
                drive_children = list(this_pc_item.children())
                
                # Try matching again after scrolling
                for item in drive_children:
                    drive_text = item.text()
                    
                    # Same matching strategies as above
                    exact_match = drive_text == target_drive
                    mapped_exact_match = mapped_name and drive_text == mapped_name
                    contains_match = target_drive in drive_text
                    mapped_contains_match = mapped_name and mapped_name in drive_text
                    
                    win11_style_match = False
                    if ":" in drive_text:
                        drive_letter_parts = re.findall(r'([A-Za-z]:)', drive_text)
                        if drive_letter_parts and drive_letter_parts[0].upper() == target_drive.upper():
                            win11_style_match = True
                    
                    if exact_match or mapped_exact_match or contains_match or mapped_contains_match or win11_style_match:
                        self.logger.info(f"Found drive after scrolling: '{drive_text}'")
                        return item
            except Exception as e:
                self.logger.warning(f"Error after scrolling for drives: {e}")
        
        self.logger.error(f"Could not find drive '{target_drive}' in This PC children")
        return None

    def is_process_complete(self, folder_path):
        """Check if mSeq has finished processing the folder"""
        # Check both possible dialog titles
        if (self.app.window(title="Low quality files skipped").exists() or
                self.app.window(title_re=".*quality.*skipped").exists()):
            print("Low quality files skipped dialog found, process is not complete")
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
                    print(f"Found dialog with title: {title}")
                    self.logger.info(f"Found dialog with title: {title}")
            except:
                pass

        # Last resort: Try with regex
        try:
            return self.app.window(title_re='Browse.*Folder')
        except:
            return None
        
    def navigate_folder_tree(self, dialog, path):
        """Navigate folder tree with enhanced Windows 11 compatibility"""
        self.logger.info(f"=== NAVIGATION TRACKING: Starting navigation to: {path} ===")
        print(f"=== NAVIGATION TRACKING: Starting navigation to: {path} ===")
        if not path or not isinstance(path, str):
            self.logger.error("Invalid path provided for navigation")
            print("Invalid path provided for navigation")
            return False

        # Initialize success tracking dictionary
        success_paths = {
            'this_pc_method': None,
            'drive_method': None,
            'folders': []
        }

        dialog.set_focus()
        self._ensure_dialog_visible(dialog)

        # Get tree view using the robust method
        tree_view = self._get_tree_view(dialog)
        if not tree_view:
            self.logger.error("Could not find TreeView control")
            print("Could not find TreeView control")
            return False
        self.logger.info("TreeView control found")
        print("TreeView control found")

        # Handle different path formats
        if ":" in path:
            # Path has a drive letter
            parts = path.split("\\")
            drive = parts[0]  # e.g., "P:"
            folders = parts[1:] if len(parts) > 1 else []
            self.logger.info(f"Path parsed: Drive={drive}, Folders={folders}")
            print(f"Path parsed: Drive={drive}, Folders={folders}")
        else:
            # Network path
            parts = path.split("\\")
            drive = "\\" + "\\".join(parts[:3])  # e.g., \\server\share
            folders = parts[3:] if len(parts) > 3 else []
            self.logger.info(f"Network path parsed: Share={drive}, Folders={folders}")
            print(f"Network path parsed: Share={drive}, Folders={folders}")

        # Enhanced navigation strategy for both Windows 10 and 11
        try:
            # Find Desktop root item
            desktop_item = None
            for item in tree_view.roots():
                if "Desktop" in item.text():
                    desktop_item = item
                    self.logger.info(f"Found Desktop: {item.text()}")
                    print(f"Found Desktop: {item.text()}")
                    break

            if not desktop_item:
                self.logger.error("Could not find Desktop root item in tree view")
                print("Could not find Desktop root item in tree view")
                return False

            # Ensure Desktop is expanded
            dialog.set_focus()
            desktop_item.click_input()
            time.sleep(self.expand_delay)
            desktop_item.expand()
            time.sleep(self.expand_delay)

            # Get This PC using enhanced method
            this_pc_item = self._get_this_pc_item(desktop_item, success_paths)
            if not this_pc_item:
                self.logger.error("Could not find This PC in Desktop children")
                print("Could not find This PC in Desktop children")
                return False

            # Find the target drive using the enhanced method (pass dialog)
            drive_item = self._find_drive_item(this_pc_item, drive, dialog, success_paths)
            if not drive_item:
                self.logger.error(f"Could not find drive '{drive}' in This PC children")
                print(f"Could not find drive '{drive}' in This PC children")
                return False

            # If we're only navigating to the drive level, we're done
            if not folders:
                self.logger.info("Navigation to drive level completed successfully")
                print("Navigation to drive level completed successfully")
                drive_item.click_input() # Ensure the drive is selected
                return True

            # Navigate through each subfolder
            current_item = drive_item
            
            # First click on the drive to ensure it's selected
            dialog.set_focus()
            current_item.click_input()
            time.sleep(self.click_delay)

            for i, folder in enumerate(folders):
                self.logger.info(f"Navigating to folder {i + 1}/{len(folders)}: {folder}")
                print(f"Navigating to folder {i + 1}/{len(folders)}: {folder}")

                # Expand current folder with extra wait time for Windows 11
                current_item.expand()
                time.sleep(self.expand_delay * 1.5)  # 50% longer delay for robust expansion

                # Get children of current folder with better error handling
                folder_children = []
                try:
                    folder_children = list(current_item.children())
                    self.logger.info(f"Current folder has {len(folder_children)} children")
                    print(f"Current folder has {len(folder_children)} children")
                    
                    # Log all children for INFOging
                    if len(folder_children) < 10:  # Only log if it's a reasonable number
                        for j, child in enumerate(folder_children):
                            self.logger.info(f"Child {j}: '{child.text()}'")
                except Exception as e:
                    self.logger.warning(f"Error getting folder children: {e}")
                    print(f"Error getting folder children: {e}")
                    # Try to scroll to see more children
                    if self._scroll_if_needed(current_item):
                        try:
                            folder_children = list(current_item.children())
                            self.logger.info(f"After scrolling, folder has {len(folder_children)} children")
                            print(f"After scrolling, folder has {len(folder_children)} children")
                        except Exception as e:
                            self.logger.error(f"Error getting folder children after scrolling: {e}")
                            print(f"Error getting folder children after scrolling: {e}")
                            return False
                    else:
                        return False

                # Look for exact folder match first
                next_item = None
                for child in folder_children:
                    child_text = child.text()
                    if child_text == folder:  # Exact match
                        if success_paths is not None:
                            success_paths['folders'].append('exact_match')
                        self.logger.info(f"SUCCESS PATH: Found exact match for folder '{folder}'")
                        print(f"SUCCESS PATH: Found exact match for folder '{folder}'")
                        dialog.set_focus()
                        child.click_input()
                        next_item = child
                        time.sleep(self.expand_delay)
                        self.logger.info(f"Found exact match for folder: {folder}")
                        print(f"Found exact match for folder: {folder}")
                        break

                # If exact match not found, try partial match
                if not next_item:
                    for child in folder_children:
                        child_text = child.text()
                        if folder.lower() in child_text.lower():  # Partial match (case insensitive)
                            if success_paths is not None:
                                success_paths['folders'].append('partial_match')
                            self.logger.info(f"SUCCESS PATH: Found partial match for folder '{folder}'")
                            print(f"SUCCESS PATH: Found partial match for folder '{folder}'")                   
                            dialog.set_focus()
                            child.click_input()
                            next_item = child
                            time.sleep(self.expand_delay)
                            self.logger.info(f"Found partial match for folder: {folder} -> {child_text}")
                            print(f"Found partial match for folder: {folder} -> {child_text}")
                            break

                # If still not found, try scrolling
                if not next_item:
                    self.logger.info(f"Folder '{folder}' not found in visible items, scrolling to find more")
                    print(f"Folder '{folder}' not found in visible items, scrolling to find more")
                    if self._scroll_if_needed(current_item):
                        # Refresh list after scrolling
                        try:
                            folder_children = list(current_item.children())

                            # Try again after scrolling
                            for child in folder_children:
                                child_text = child.text()
                                if child_text == folder or folder.lower() in child_text.lower():
                                    dialog.set_focus()
                                    child.click_input()
                                    next_item = child
                                    time.sleep(self.expand_delay)
                                    self.logger.info(f"Found folder after scrolling: {child_text}")
                                    print(f"Found folder after scrolling: {child_text}")
                                    break
                        except Exception as e:
                            self.logger.warning(f"Error after scrolling for folder: {e}")
                            print(f"Error after scrolling for folder: {e}")
                    else:
                        self.logger.error(f"Could not find folder '{folder}' even after scrolling")
                        print(f"Could not find folder '{folder}' even after scrolling")
                        return False

                # If folder still not found, we can't continue
                if not next_item:
                    self.logger.error(f"Could not find folder '{folder}' in the current directory")
                    print(f"Could not find folder '{folder}' in the current directory")
                    return False

                # Update current item for next iteration
                current_item = next_item

            # Make sure the final folder is selected
            current_item.click_input()
            
            # Navigation completed successfully
            self.logger.info("Navigation completed successfully")
            print("Navigation completed successfully")
            # Log navigation summary if we were tracking success paths
            if 'success_paths' in locals() and success_paths:
                self._log_navigation_summary(path, success_paths)
            
            # Analyze logs for success paths (uncomment these lines when you want to use this feature)
            summary = self.analyze_navigation_logs()
            self.logger.info(summary)
            return True

        except Exception as e:
            # Log the error but don't raise it - we want to continue even if navigation fails
            self.logger.error(f"Error during folder navigation: {e}")
            print(f"Error during folder navigation: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False

    def _fast_navigate_folder_tree(self, dialog, path):
        """Patched method with faster navigation for Windows 11"""
        import logging
        logger = logging.getLogger(__name__)

        # Add success tracking dictionary
        success_paths = {
            'this_pc_method': None,
            'drive_method': None,
            'folders': []
        }
        
        logger.info(f"=== NAVIGATION TRACKING (FAST MODE): Starting navigation to: {path} ===")
        print(f"=== NAVIGATION TRACKING (FAST MODE): Starting navigation to: {path} ===")
        dialog.set_focus()

        # Get tree view
        tree_view = self._get_tree_view(dialog)
        if not tree_view:
            logger.error("Could not find TreeView control")
            print("Could not find TreeView control")
            return False

        # Standard navigation with reduced delays
        try:
            # Parse the path
            if ":" in path:
                parts = path.split("\\")
                drive = parts[0]
                folders = parts[1:] if len(parts) > 1 else []
                logger.info(f"TRACKING: Path parsed: Drive={drive}, Folders={folders}")
                print(f"TRACKING: Path parsed: Drive={drive}, Folders={folders}")
            else:
                parts = path.split("\\")
                drive = "\\" + "\\".join(parts[:3])
                folders = parts[3:] if len(parts) > 3 else []
                logger.info(f"TRACKING: Network path parsed: Share={drive}, Folders={folders}")
                print(f"TRACKING: Network path parsed: Share={drive}, Folders={folders}")

            # Find Desktop
            desktop_item = None
            for root_item in tree_view.roots():
                if "Desktop" in root_item.text():
                    desktop_item = root_item
                    logger.info(f"TRACKING: Found Desktop item: {root_item.text()}")
                    print(f"TRACKING: Found Desktop item: {root_item.text()}")
                    break
            
            if not desktop_item:
                logger.error("Could not find Desktop root item")
                print("Could not find Desktop root item")
                return False

            # Ensure Desktop is expanded
            dialog.set_focus()
            desktop_item.click_input()
            time.sleep(self.expand_delay)
            desktop_item.expand()
            time.sleep(self.expand_delay)

            # Get This PC using enhanced method (handles lower positions)
            this_pc_item = self._get_this_pc_item(desktop_item, success_paths)
            if not this_pc_item:
                logger.error("Could not find This PC in Desktop children")
                print("Could not find This PC in Desktop children")
                return False
            
            logger.info(f"TRACKING: This PC found via method: {success_paths.get('this_pc_method', 'unknown')}")
            print(f"TRACKING: This PC found via method: {success_paths.get('this_pc_method', 'unknown')}")

            # Expand This PC
            this_pc_item.click_input()
            time.sleep(self.expand_delay)
            this_pc_item.expand()
            time.sleep(self.expand_delay * 2)  # Double delay for Windows 11

            # Find drive using enhanced method (doesn't rely on parent())
            drive_item = self._find_drive_item(this_pc_item, drive, dialog, success_paths)
            if not drive_item:
                logger.error(f"Could not find drive {drive}")
                print(f"Could not find drive {drive}")
                return False
            
            logger.info(f"TRACKING: Drive found via method: {success_paths.get('drive_method', 'unknown')}")
            print(f"TRACKING: Drive found via method: {success_paths.get('drive_method', 'unknown')}")
                    
            # Make sure the drive is selected
            drive_item.click_input()
            time.sleep(self.click_delay)

            # If we're only navigating to the drive level, we're done
            if not folders:
                logger.info("TRACKING: Navigation completed at drive level")
                print("TRACKING: Navigation completed at drive level")
                return True

            # Navigate folders with minimal delay
            current_item = drive_item
            for i, folder in enumerate(folders):
                logger.info(f"TRACKING: Navigating to folder {i+1}/{len(folders)}: {folder}")
                print(f"TRACKING: Navigating to folder {i+1}/{len(folders)}: {folder}")
                current_item.expand()
                time.sleep(self.expand_delay)

                folder_found = False
                folder_method = "unknown"
                try:
                    # Try exact match first
                    for child in current_item.children():
                        if child.text() == folder:
                            child.click_input()
                            current_item = child
                            folder_found = True
                            folder_method = "exact_match"
                            logger.info(f"TRACKING: Found exact match for folder: {folder}")
                            print(f"TRACKING: Found exact match for folder: {folder}")
                            break
                    
                    # Try partial match if exact match failed
                    if not folder_found:
                        for child in current_item.children():
                            if folder.lower() in child.text().lower():
                                child.click_input()
                                current_item = child
                                folder_found = True
                                folder_method = "partial_match"
                                logger.info(f"TRACKING: Found partial match for folder: {folder}")
                                print(f"TRACKING: Found partial match for folder: {folder}")
                                break
                    
                    # Try scrolling to find more
                    if not folder_found:
                        logger.info(f"TRACKING: Scrolling to find folder: {folder}")
                        print(f"TRACKING: Scrolling to find folder: {folder}")
                        self._scroll_if_needed(current_item)
                        for child in current_item.children():
                            if child.text() == folder or folder.lower() in child.text().lower():
                                child.click_input()
                                current_item = child
                                folder_found = True
                                folder_method = "after_scrolling"
                                logger.info(f"TRACKING: Found folder after scrolling: {folder}")
                                print(f"TRACKING: Found folder after scrolling: {folder}")
                                break
                    
                    # Record success method
                    if folder_found:
                        success_paths['folders'].append(folder_method)
                except Exception as e:
                    logger.error(f"Error navigating to folder {folder}: {e}")
                    print(f"Error navigating to folder {folder}: {e}")
                    return False

                if not folder_found:
                    logger.error(f"Could not find folder {folder}")
                    print(f"Could not find folder {folder}")
                    return False

            # Make sure final item is selected
            current_item.click_input()
            
            # Log navigation summary
            self._log_navigation_summary(path, success_paths)
            return True

        except Exception as e:
            logger.error(f"Error in fast navigation: {e}")
            print(f"Error in fast navigation: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False

    def _click_dialog_button(self, dialog_window, button_titles):
        """Click a button in a dialog with better OS compatibility"""
        from pywinauto.keyboard import send_keys

        if not dialog_window or not dialog_window.exists():
            self.logger.warning("Dialog not found for button click operation")
            print("Dialog not found for button click operation")
            return False

        # Try each button title in order
        for btn_title in button_titles:
            try:
                ok_button = dialog_window.child_window(title=btn_title, class_name="Button")
                if ok_button.exists():
                    ok_button.click_input()
                    print(f"Clicked button '{btn_title}'")
                    self.logger.info(f"Clicked button '{btn_title}'")
                    time.sleep(self.click_delay)  # OS-optimized delay
                    return True
            except Exception as e:
                self.logger.info(f"Button '{btn_title}' not found: {e}")
                print(f"Button '{btn_title}' not found: {e}")
                continue

        # Fallback to keyboard if no button found
        try:
            dialog_window.set_focus()
            send_keys('{ENTER}')
            print("No specific button found, pressed Enter as fallback")
            time.sleep(self.click_delay)  # OS-optimized delay
            return True
        except Exception as e:
            self.logger.error(f"Failed to send keyboard input: {e}")
            print(f"Failed to send keyboard input: {e}")
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
                print("Preferences dialog not found or not visible")
        except Exception as e:
            self.logger.error(f"Error handling preferences dialog: {e}")
            print(f"Error handling preferences dialog: {e}")

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
                        print("Windows 10 ShellView found")
                        if list_view.exists():
                            list_view.click_input()
                            list_view_found = True
                            print("Windows 10 list view found and clicked")
                except Exception as e:
                    self.logger.info(f"Windows 10 list view approach failed: {e}")
                    print(f"Windows 10 list view approach failed: {e}")

                # Windows 11 approach
                if not list_view_found:
                    print("Trying Windows 11 approach")
                    try:
                        list_view = copy_files_window.child_window(class_name="DirectUIHWND")
                        print("Windows 11 DirectUIHWND found")
                        self.logger.info("Windows 11 DirectUIHWND found")
                        if list_view.exists():
                            list_view.click_input()
                            list_view_found = True
                            print("Windows 11 list view found and clicked")
                    except Exception as e:
                        self.logger.info(f"Windows 11 list view approach failed: {e}")
                        print(f"Windows 11 list view approach failed: {e}")

                # Last resort - click in the middle of the dialog
                if not list_view_found:
                    try:
                        rect = copy_files_window.rectangle()
                        copy_files_window.click_input(coords=((rect.right - rect.left) // 2,
                                                              (rect.bottom - rect.top) // 2))
                        list_view_found = True
                        print("Clicked center of Copy Files dialog as last resort")
                    except Exception as e:
                        self.logger.warning(f"Center-click approach failed: {e}")
                        print(f"Center-click approach failed: {e}")

                if list_view_found:
                    # Select all files
                    send_keys('^a')  # Ctrl+A
                    time.sleep(self.click_delay)  # Wait for selection
                    print("All files selected in the list view")
                    # Click the Open button
                    self._click_dialog_button(copy_files_window, ["&Open", "Open"])
                    print("Clicked Open button in Copy Files dialog")
            else:
                self.logger.warning("Copy files dialog not found or not visible")
                print("Copy files dialog not found or not visible")
        except Exception as e:
            self.logger.error(f"Error handling copy files dialog: {e}")
            print(f"Error handling copy files dialog: {e}")

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
                print("Error dialog found")

                # If no specific button found, try any button
                if not button_found:
                    try:
                        ok_button = error_window.child_window(class_name="Button")
                        if ok_button.exists():
                            ok_button.click_input()
                            time.sleep(self.click_delay)
                            print("Clicked any available button in error dialog")
                        else:
                            # Last resort - press Enter
                            error_window.set_focus()
                            send_keys('{ENTER}')
                            time.sleep(self.click_delay)
                    except Exception as e:
                        self.logger.warning(f"Failed to handle error dialog with generic approach: {e}")
                        print(f"Failed to handle error dialog with generic approach: {e}")
            else:
                self.logger.info("No error dialog found or it disappeared")
                print("No error dialog found or it disappeared")
        except Exception as e:
            self.logger.error(f"Error handling error dialog: {e}")
            print(f"Error handling error dialog: {e}")

    def _handle_call_bases_dialog(self):
        """Handle Call bases dialog with OS awareness"""
        try:
            call_bases_window = self.app.window(title_re='Call bases.*')

            if call_bases_window and call_bases_window.exists():
                # Click Yes button
                button_found = self._click_dialog_button(call_bases_window, ["&Yes", "Yes"])
                print("Call bases dialog found")
                self.logger.info("Call bases dialog found")

                if not button_found:
                    # Try any button that might work
                    try:
                        yes_button = call_bases_window.child_window(class_name="Button")
                        if yes_button.exists():
                            yes_button.click_input()
                            print("Clicked any available button in call bases dialog")
                        else:
                            # Last resort - press Enter for default action
                            call_bases_window.set_focus()
                            send_keys('{ENTER}')
                    except Exception as e:
                        self.logger.warning(f"Failed to handle call bases dialog with generic approach: {e}")
                        print(f"Failed to handle call bases dialog with generic approach: {e}")
            else:
                self.logger.warning("Call bases dialog not found or not visible")
                print("Call bases dialog not found or not visible")
        except Exception as e:
            self.logger.error(f"Error handling call bases dialog: {e}")
            print(f"Error handling call bases dialog: {e}")

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
                            print("Clicked any available button in low quality dialog")
                        else:
                            # Last resort - press Enter
                            low_quality_window.set_focus()
                            send_keys('{ENTER}')
                            print("Pressed Enter in low quality dialog as last resort")

                    self.logger.info(f"Handled dialog: {title}")
                    print(f"Handled dialog: {title}")
                    return True

            return False
        except Exception as e:
            self.logger.error(f"Error handling low quality dialog: {e}")
            print(f"Error handling low quality dialog: {e}")
            return False

    def _wait_for_process_completion(self, folder_path, max_wait=None, interval=None):
        """Wait for mSeq processing to complete with OS-specific timeouts"""
        if max_wait is None:
            max_wait = OSCompatibilityManager.get_timeout("process_completion")
            print(f"Using default max wait time: {max_wait}s")

        if interval is None:
            interval = OSCompatibilityManager.get_timeout("polling_interval", 0.5)
            print(f"Using default polling interval: {interval}s")

        self.logger.info(f"Waiting for process completion (max_wait={max_wait}s, interval={interval}s)")
        print(f"Waiting for process completion (max_wait={max_wait}s, interval={interval}s)")

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
                    print("Read information dialog found")
                    read_window.close()
                    completed = True
                    self.logger.info("Closed read information dialog")
                    print("Closed read information dialog")
            except Exception as e:
                self.logger.info(f"Error checking read info dialog: {e}")
                print(f"Error checking read info dialog: {e}")

            # Check for txt files - most reliable completion indicator
            txt_count = 0
            for item in os.listdir(folder_path):
                if any(item.endswith(ext) for ext in self.config.TEXT_FILES):
                    txt_count += 1

            if txt_count >= 5:
                self.logger.info(f"Processing completed for {folder_path}: found {txt_count} text files")
                print(f"Processing completed for {folder_path}: found {txt_count} text files")
                completed = True

            if completed:
                return True

            time.sleep(interval)
            elapsed += interval

            # Log progress periodically
            if int(elapsed) % 5 == 0 and int(elapsed) > 0:
                self.logger.info(f"Still waiting for processing completion ({int(elapsed)}s elapsed)")
                print(f"Still waiting for processing completion ({int(elapsed)}s elapsed)")

        self.logger.warning(f"Timeout waiting for processing to complete for {folder_path}")
        print(f"Timeout waiting for processing to complete for {folder_path}")
        return True  # Return True anyway to continue with next folder

    def process_folder(self, folder_path):
        """Process a folder with mSeq using OS-aware optimizations"""
        self.logger.info(f"Processing folder: {folder_path}")
        print(f"Processing folder: {folder_path}")

        # Validate folder exists
        if not os.path.exists(folder_path):
            self.logger.warning(f"Folder does not exist: {folder_path}")
            print(f"Folder does not exist: {folder_path}")
            return False

        # Check for .ab1 files to process
        ab1_files = [f for f in os.listdir(folder_path) if f.endswith(self.config.ABI_EXTENSION)]
        if not ab1_files:
            self.logger.warning(f"No {self.config.ABI_EXTENSION} files found in {folder_path}, skipping processing")
            print(f"No {self.config.ABI_EXTENSION} files found in {folder_path}, skipping processing")
            return False

        self.logger.info(f"Found {len(ab1_files)} AB1 files to process")
        print(f"Found {len(ab1_files)} AB1 files to process")

        if self.use_fast_navigation:
            return self._fast_process_folder(folder_path)

        # Connect to mSeq
        self.app, self.main_window = self.connect_or_start_mseq()
        if not self.main_window:
            self.logger.error("Failed to connect to mSeq")
            print("Failed to connect to mSeq")
            return False

        self.main_window.set_focus()
        from pywinauto.keyboard import send_keys
        send_keys('^n')  # Ctrl+N for new project
        self.logger.info("Sent Ctrl+N to create new project")
        print("Sent Ctrl+N to create new project")

        # Wait for and handle Browse For Folder dialog
        browse_timeout = OSCompatibilityManager.get_timeout("browse_dialog")
        dialog_found = self.wait_for_dialog("browse_dialog", timeout=browse_timeout)
        print(f"Browse dialog wait result: {dialog_found}")

        if not dialog_found:
            self.logger.error("Browse For Folder dialog not found")
            print("Browse For Folder dialog not found")
            return False

        dialog_window = self._get_browse_dialog()
        if not dialog_window:
            self.logger.error("Failed to get Browse For Folder dialog window")
            print("Failed to get Browse For Folder dialog window")
            return False

        # Add a delay for the first browsing operation - needed on both Windows 10 and 11
        if self.first_time_browsing:
            self.first_time_browsing = False
            first_browse_delay = OSCompatibilityManager.get_timeout("first_browse_delay", 1.2)
            self.logger.info(f"First time browsing, adding delay of {first_browse_delay}s")
            print(f"First time browsing, adding delay of {first_browse_delay}s")
            time.sleep(first_browse_delay)
        else:
            subsequent_browse_delay = OSCompatibilityManager.get_timeout("subsequent_browse_delay", 0.5)
            print(f"Subsequent browsing, adding delay of {subsequent_browse_delay}s")
            time.sleep(subsequent_browse_delay)

        # Navigate to the target folder
        navigate_success = self.navigate_folder_tree(dialog_window, folder_path)
        if not navigate_success:
            self.logger.error(f"Navigation failed for {folder_path}")
            print(f"Navigation failed for {folder_path}")
            return False

        self.logger.info("Navigation successful, clicking OK")
        print("Navigation successful, clicking OK")

        # Find and click OK button with better error handling
        button_found = self._click_dialog_button(dialog_window, ["OK", "&OK", "Ok"])
        if not button_found:
            self.logger.warning("OK button not found by name, trying fallback methods")
            print("OK button not found by name, trying fallback methods")
            try:
                # Try to find any button
                buttons = dialog_window.children(class_name="Button")
                if buttons:
                    self.logger.info(f"Found {len(buttons)} buttons, clicking the first one")
                    print(f"Found {len(buttons)} buttons, clicking the first one")
                    buttons[0].click_input()
                else:
                    # Last resort - press Enter
                    self.logger.info("No buttons found, pressing Enter")
                    print("No buttons found, pressing Enter")
                    dialog_window.set_focus()
                    send_keys('{ENTER}')
            except Exception as e:
                self.logger.error(f"Error with fallback OK button handling: {e}")
                print(f"Error with fallback OK button handling: {e}")
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
            print(f"Successfully processed folder: {folder_path}")
        else:
            self.logger.warning(f"Processing may not have completed properly for: {folder_path}")
            print(f"Processing may not have completed properly for: {folder_path}")

        return True

    def _fast_process_folder(self, folder_path):
        """Patched method with faster processing and reduced waits"""
        self.logger.info(f"=== PROCESS TRACKING (FAST MODE): Processing folder: {folder_path} ===")
        print(f"=== PROCESS TRACKING (FAST MODE): Processing folder: {folder_path} ===")
        import time
        if not os.path.exists(folder_path):
            self.logger.error(f"Folder does not exist: {folder_path}")
            print(f"Folder does not exist: {folder_path}")
            return False

        # Check for .ab1 files
        ab1_files = [f for f in os.listdir(folder_path) if f.endswith('.ab1')]
        if not ab1_files:
            self.logger.error(f"No .ab1 files found in {folder_path}, skipping processing")
            print(f"No .ab1 files found in {folder_path}, skipping processing")
            return False
        
        self.logger.info(f"TRACKING: Found {len(ab1_files)} AB1 files to process")
        print(f"TRACKING: Found {len(ab1_files)} AB1 files to process")

        self.app, self.main_window = self.connect_or_start_mseq()
        self.main_window.set_focus()
        from pywinauto.keyboard import send_keys
        send_keys('^n')  # Ctrl+N for new project
        self.logger.info("TRACKING: Sent Ctrl+N to open new project")
        print("TRACKING: Sent Ctrl+N to open new project")

        # Wait for Browse dialog
        dialog_wait_result = self.wait_for_dialog("browse_dialog")
        self.logger.info(f"TRACKING: Browse dialog wait result: {dialog_wait_result}")
        print(f"TRACKING: Browse dialog wait result: {dialog_wait_result}")
        
        dialog_window = self._get_browse_dialog()
        if not dialog_window:
            self.logger.error("Failed to find Browse For Folder dialog")
            print("Failed to find Browse For Folder dialog")
            return False
        
        self.logger.info(f"TRACKING: Found Browse dialog: {dialog_window.window_text()}")
        print(f"TRACKING: Found Browse dialog: {dialog_window.window_text()}")

        # Navigate to folder with faster method
        navigate_success = self._fast_navigate_folder_tree(dialog_window, folder_path)
        self.logger.info(f"TRACKING: Fast navigation result: {navigate_success}")
        print(f"TRACKING: Fast navigation result: {navigate_success}")
        if not navigate_success:
            self.logger.error(f"Fast navigation failed for {folder_path}")
            print(f"Fast navigation failed for {folder_path}")
            return False

        # Click OK
        try:
            ok_button = dialog_window.child_window(title="OK", class_name="Button")
            if ok_button.exists():
                self.logger.info("TRACKING: Clicking OK button")
                print("TRACKING: Clicking OK button")
                ok_button.click_input()
            else:
                self.logger.info("TRACKING: OK button not found, pressing Enter")
                print("TRACKING: OK button not found, pressing Enter")
                dialog_window.set_focus()
                send_keys('{ENTER}')
        except Exception as e:
            self.logger.warning(f"TRACKING: Error clicking OK: {e}")
            print(f"TRACKING: Error clicking OK: {e}")
            dialog_window.set_focus()
            send_keys('{ENTER}')

        # Log progress through remaining dialogs
        dialog_results = {}

        # Handle preferences
        self.wait_for_dialog("preferences")
        pref_result = self.wait_for_dialog("preferences")
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
                for btn_title in ["&OK", "OK", "Ok"]:
                    try:
                        ok_button = pref_window.child_window(title=btn_title)
                        if ok_button.exists():
                            ok_button.click_input()
                            break
                    except:
                        pass
                else:
                    pref_window.set_focus()
                    send_keys('{ENTER}')
        except:
            pass
        dialog_results['preferences'] = pref_result

        # Handle Copy files
        self.wait_for_dialog("copy_files")
        copy_result = self.wait_for_dialog("copy_files") 
        try:
            copy_files_window = self.app.window(title_re='Copy.*sequence files')

            # Try to find list view and select all
            try:
                shell_view = copy_files_window.child_window(title="ShellView")
                print("TRACKING: Found ShellView for Copy Files dialog")
                list_view = shell_view.child_window(class_name="DirectUIHWND")
                print("TRACKING: Found DirectUIHWND for list view")
                if list_view.exists():
                    list_view.click_input()
            except:
                try:
                    list_view = copy_files_window.child_window(class_name="DirectUIHWND")
                    list_view.click_input()
                except:
                    rect = copy_files_window.rectangle()
                    copy_files_window.click_input(coords=((rect.right - rect.left) // 2,
                                                          (rect.bottom - rect.top) // 2))

            send_keys('^a')  # Select all

            # Click Open
            for btn_title in ["&Open", "Open"]:
                try:
                    open_button = copy_files_window.child_window(title=btn_title)
                    if open_button.exists():
                        open_button.click_input()
                        break
                except:
                    pass
            else:
                copy_files_window.set_focus()
                send_keys('{ENTER}')
        except:
            pass

        dialog_results['copy_files'] = copy_result

        # Handle Error dialog
        self.wait_for_dialog("error_window")
        error_result = self.wait_for_dialog("error_window")
        try:
            error_window = None
            for title in ['File error', 'Error']:
                try:
                    error_window = self.app.window(title=title)
                    if error_window.exists():
                        break
                except:
                    pass

            if error_window and error_window.exists():
                ok_button = error_window.child_window(class_name="Button")
                if ok_button.exists():
                    ok_button.click_input()
                else:
                    error_window.set_focus()
                    send_keys('{ENTER}')
        except:
            pass
        dialog_results['error_window'] = error_result

        # Handle Call bases
        self.wait_for_dialog("call_bases")
        call_bases_result = self.wait_for_dialog("call_bases")
        try:
            call_bases_window = self.app.window(title_re='Call bases.*')
            print("TRACKING: Found Call bases dialog")

            if call_bases_window and call_bases_window.exists():
                for btn_title in ["&Yes", "Yes"]:
                    try:
                        yes_button = call_bases_window.child_window(title=btn_title)
                        if yes_button.exists():
                            yes_button.click_input()
                            print("TRACKING: Clicked Yes button in Call bases dialog")
                            break
                    except:
                        pass
                else:
                    call_bases_window.set_focus()
                    send_keys('{ENTER}')
                    print("TRACKING: Pressed Enter in Call bases dialog as fallback")
        except:
            pass
        dialog_results['call_bases'] = call_bases_result
        

        # Log dialog results
        self.logger.info(f"TRACKING: Dialog wait results: {dialog_results}")
        print(f"TRACKING: Dialog wait results: {dialog_results}")

        # Add more detailed wait timing for process completion
        self.logger.info(f"TRACKING: Starting to wait for process completion")
        print(f"TRACKING: Starting to wait for process completion")
        process_start_time = time.time()

        # Wait for completion - use a more aggressive check
        import time
        max_wait = self.config.TIMEOUTS["process_completion"]
        interval = 0.5
        elapsed = 0
        completion_reason = None

        while elapsed < max_wait:
            # Check if process completed
            completed = False

            # Check for low quality dialog
            try:
                for title in ["Low quality files skipped", "Quality files skipped"]:
                    if self.app.window(title=title).exists():
                        self.logger.info(f"TRACKING: Found completion dialog: {title}")
                        print(f"TRACKING: Found completion dialog: {title}")
                        completion_reason = f"Dialog: {title}"
                        low_quality_window = self.app.window(title=title)
                        ok_button = low_quality_window.child_window(class_name="Button")
                        ok_button.click_input()
                        completed = True
                        break
            except Exception as e:
                self.logger.info(f"TRACKING: Error checking low quality dialog: {e}")
                print(f"TRACKING: Error checking low quality dialog: {e}")

            # Check for read info dialog
            try:
                if self.app.window(title_re='Read information for*').exists():
                    read_title = self.app.window(title_re='Read information for*').window_text()
                    self.logger.info(f"TRACKING: Found read info dialog: {read_title}")
                    print(f"TRACKING: Found read info dialog: {read_title}")
                    # Close the dialog
                    completion_reason = f"Dialog: {read_title}"
                    read_window = self.app.window(title_re='Read information for*')
                    print("TRACKING: Closing read information dialog")
                    read_window.close()
                    completed = True
            except Exception as e:
                self.logger.info(f"TRACKING: Error checking read info dialog: {e}")
                print(f"TRACKING: Error checking read info dialog: {e}")

            # Check for txt files
            txt_count = 0
            for item in os.listdir(folder_path):
                if (item.endswith('.raw.qual.txt') or
                        item.endswith('.raw.seq.txt') or
                        item.endswith('.seq.info.txt') or
                        item.endswith('.seq.qual.txt') or
                        item.endswith('.seq.txt')):
                    txt_count += 1

            if txt_count >= 5:
                self.logger.info(f"TRACKING: Found {txt_count} text files - process completed")
                print(f"TRACKING: Found {txt_count} text files - process completed")
                completion_reason = f"Text files: {txt_count}"
                completed = True

            if completed:
                process_elapsed = time.time() - process_start_time
                self.logger.info(f"TRACKING: Process completed in {process_elapsed:.2f}s - Reason: {completion_reason}")
                print(f"TRACKING: Process completed in {process_elapsed:.2f}s - Reason: {completion_reason}")
                return True

            # Get appropriate interval from the compatibility manager
            wait_interval = OSCompatibilityManager.get_timeout("polling_interval", interval)
            time.sleep(wait_interval)
            elapsed += wait_interval  # Make sure to update elapsed with the actual wait time

            # Log progress periodically
            if int(elapsed) % 5 == 0 and int(elapsed) > 0:
                self.logger.info(f"TRACKING: Still waiting for completion: {elapsed:.1f}s elapsed")
                print(f"TRACKING: Still waiting for completion: {elapsed:.1f}s elapsed")
                if completion_reason is not None:
                    self.logger.info(f"TRACKING: Current completion reason: {completion_reason}")
                    print(f"TRACKING: Current completion reason: {completion_reason}")

        self.logger.warning(f"TRACKING: Timeout after {elapsed:.1f}s waiting for processing to complete")
        print(f"TRACKING: Timeout after {elapsed:.1f}s waiting for processing to complete")
        print(f"Warning: Timeout waiting for processing to complete for {folder_path}")
        print("This may be normal if the folder has already been processed or has special files")
        return True  # Return True anyway to continue with next folder


    def _log_navigation_summary(self, path, success_paths):
        """Log a summary of the successful navigation approaches
        
        Args:
            path: The full path that was navigated
            success_paths: Dictionary of successful approaches, e.g.,
                {
                    'this_pc_method': 'direct_search',
                    'drive_method': 'win11_style_match',
                    'folders': ['exact_match', 'partial_match', 'after_scrolling']
                }
        """
        summary = [
            "======================================================",
            f"NAVIGATION SUMMARY FOR: {path}",
            "------------------------------------------------------",
            f"This PC found via: {success_paths.get('this_pc_method', 'unknown')}",
            f"Drive found via: {success_paths.get('drive_method', 'unknown')}",
        ]
        
        # Add folder methods if applicable
        folder_methods = success_paths.get('folders', [])
        if folder_methods:
            folder_parts = path.split('\\')[1:] if '\\' in path else []
            for i, (folder, method) in enumerate(zip(folder_parts, folder_methods)):
                if i < len(folder_methods):
                    summary.append(f"Folder '{folder}' found via: {method}")
        
        summary.append("======================================================")
        
        # Log each line separately for better readability
        for line in summary:
            self.logger.info(line)

    def analyze_navigation_logs(self, log_file=None):
        """Analyze the log file to extract which navigation methods succeeded
        
        This can be called after navigation completes to get a summary.
        
        Args:
            log_file: Path to the log file (uses default logger file if None)
            
        Returns:
            A string summary of which methods succeeded
        """
        import os
        
        # If no log file specified, try to find it
        if log_file is None:
            # Try to get logger handlers
            for handler in self.logger.handlers:
                if hasattr(handler, 'baseFilename'):
                    log_file = handler.baseFilename
                    break
        
        if not log_file or not os.path.exists(log_file):
            return "Could not find log file to analyze"
        
        # Read the log file
        success_paths = []
        try:
            with open(log_file, 'r') as f:
                log_content = f.read()
                
            # Extract success paths using regex
            import re
            success_matches = re.findall(r'SUCCESS PATH: (.*)', log_content)
            
            if not success_matches:
                return "No successful navigation paths found in log"
                
            return "Navigation Success Summary:\n" + "\n".join(success_matches)
        
        except Exception as e:
            return f"Error analyzing log file: {e}"



    def close(self):
        """Close the mSeq application with better error handling"""
        if self.app:
            self.logger.info("Attempting to close mSeq application")
            print("Attempting to close mSeq application")
            try:
                self.app.kill()
                self.logger.info("mSeq application killed successfully")
                print("mSeq application killed successfully")
            except Exception as e:
                self.logger.warning(f"Error killing mSeq process: {e}")
                print(f"Error killing mSeq process: {e}")
                # Try alternative approach
                if self.main_window and self.main_window.exists():
                    try:
                        self.main_window.close()
                        self.logger.info("mSeq window closed successfully")
                        print("mSeq window closed successfully")
                    except Exception as e2:
                        self.logger.error(f"Error closing mSeq window: {e2}")
                        print(f"Error closing mSeq window: {e2}")
        else:
            self.logger.info("No mSeq application to close")
            print("No mSeq application to close")