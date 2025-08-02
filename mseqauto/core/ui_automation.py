# ui_automation.py
import os
import time
import logging
import re
from pathlib import Path

from pywinauto import Application, timings
from pywinauto.keyboard import send_keys
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
import win32api
# Add parent directory to PYTHONPATH for imports
import os
import sys
sys.path.append(str(Path(__file__).parents[2]))
import warnings
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", module="pywinauto")

class MseqAutomation:
    """Streamlined automation class for controlling mSeq software"""


    def __init__(self, config, logger=None):
        """Initialize the automation with configuration settings"""
        self.config = config
        self.app = None
        self.main_window = None
        self.first_time_browsing = True

        # Set up logging
        if logger is None:
            # Import and set up logger if none provided
            import sys
            sys.path.append(str(Path(__file__).parents[1]))
            from utils.logger import setup_logger
            # Create logs directory in the same directory as this file
            log_dir = Path(__file__).resolve().parent.parent / "scripts" / "logs"
            self.logger = setup_logger("ui_automation", log_dir=log_dir)
        else:
            self.logger = logger

        # Basic timing values - increase for Windows 11
        self.is_win11 = self._is_windows_11()
        self.expand_delay = 0.3 if self.is_win11 else 0.2
        self.click_delay = 0.2 if self.is_win11 else 0.1

        # Define standard timeout values
        self.timeouts = {
            "connect": 10,
            "browse_dialog": 7.5 if self.is_win11 else 5,
            "preferences": 3,
            "copy_files": 3,
            "error_window": 5,
            "call_bases": 5,
            "low_quality": 3,
            "process_completion": 15,  # Increased timeout to catch Low quality dialog
            "read_info": 3
        }

        self.logger.info(f"MseqAutomation initialized (Windows 11: {self.is_win11})")

    def _is_windows_11(self):
        """Detect if running on Windows 11"""
        import platform
        if platform.system() != 'Windows':
            return False

        win_version = int(platform.version().split('.')[0])
        win_build = int(platform.version().split('.')[2]) if len(platform.version().split('.')) > 2 else 0
        return win_version >= 10 and win_build >= 22000

    def connect_or_start_mseq(self):
        """Connect to existing mSeq instance or start a new one"""
        try:
            # Try to connect to an existing instance
            self.app = Application(backend='win32').connect(title_re='[mM]seq.*', timeout=1)
            self.logger.info("Connected to existing mSeq instance")
        except (ElementNotFoundError, timings.TimeoutError):
            # Start a new instance
            self.logger.info("Starting new mSeq instance")
            start_cmd = f'cmd /c "cd /d {self.config.MSEQ_PATH} && {self.config.MSEQ_EXECUTABLE}"'

            try:
                self.app = Application(backend='win32').start(start_cmd, wait_for_idle=False)
                self.app.connect(title='mSeq', timeout=self.timeouts["connect"])
            except Exception as e:
                self.logger.error(f"Failed to start mSeq: {e}")
                raise
        except ElementAmbiguousError:
            # Handle multiple instances
            self.app = Application(backend='win32').connect(title_re='[mM]seq.*', found_index=0)
            self.logger.warning("Multiple mSeq windows found, connecting to first instance")

        # Get the main window
        for title_pattern in ['mSeq.*', 'Mseq.*']:
            try:
                self.main_window = self.app.window(title_re=title_pattern)
                if self.main_window.exists():
                    break
            except:
                pass

        if not self.main_window or not self.main_window.exists():
            self.logger.error("Could not find mSeq main window")
            raise RuntimeError("Could not find mSeq main window")

        return self.app, self.main_window

    def process_folder(self, folder_path):
        """Process a folder with mSeq - streamlined version based on successful path"""
        folder = Path(folder_path)
        if not folder.exists():
            self.logger.warning(f"Folder does not exist: {folder_path}")
            return False

        # Check for AB1 files
        ab1_files = [f.name for f in folder.iterdir() if f.suffix == self.config.ABI_EXTENSION]
        if not ab1_files:
            self.logger.warning(f"No AB1 files found in {folder_path}")
            return False

        self.logger.info(f"Processing folder with {len(ab1_files)} AB1 files: {folder_path}")

        # Close any existing Read information windows before starting
        self._close_all_read_info_dialogs()

        # Connect to mSeq and get main window
        self.app, self.main_window = self.connect_or_start_mseq() #type: ignore
        self.main_window.set_focus()

        # Start new project (Ctrl+N)
        send_keys('^n')
        self.logger.info("Started new mSeq project")

        # Wait for and handle Browse For Folder dialog
        browse_found, browse_dialog = self._wait_for_dialog("browse_dialog")
        if not browse_found or not browse_dialog:
            self.logger.error("Browse For Folder dialog not found")
            return False

        # Add a delay for the first browsing operation
        if self.first_time_browsing:
            self.first_time_browsing = False
            time.sleep(1.0)

        # Navigate to the target folder
        if not self._navigate_folder_tree(browse_dialog, folder_path):
            self.logger.error(f"Failed to navigate to {folder_path}")
            return False

        # Click OK button
        self._click_dialog_button(browse_dialog, ["OK", "&OK"])

        # Handle mSeq Preferences dialog
        pref_found, pref_dialog = self._wait_for_dialog("preferences")
        if pref_found and pref_dialog:
            self._click_dialog_button(pref_dialog, ["&OK", "OK"])

        # Handle Copy sequence files dialog
        copy_found, copy_dialog = self._wait_for_dialog("copy_files")
        if copy_found and copy_dialog:
            # Select all files
            self._select_all_files_in_dialog(copy_dialog)
            self._click_dialog_button(copy_dialog, ["&Open", "Open"])
          # Handle File error dialog (appears due to non-sequence files) or wdhandler error
        error_found, error_dialog = self._wait_for_dialog("error_window")
        if error_found and error_dialog:
            self.logger.info("Found error dialog, attempting to dismiss...")
            success = self._click_dialog_button(error_dialog, ["OK"])
            if success:
                self.logger.info("Successfully dismissed error dialog")
            else:
                self.logger.warning("Failed to dismiss error dialog - this may cause issues")
        else:
            # Check for wdhandler error dialog (alternative to File error)
            wdhandler_dialog = self._get_dialog_by_titles(['wdhandler'])
            if wdhandler_dialog:
                self.logger.warning("Detected wdhandler error dialog from mSeq, dismissing...")
                success = self._click_dialog_button(wdhandler_dialog, ["OK", "&OK"])
                if success:
                    self.logger.warning("wdhandler error prevents processing - skipping folder")
                    return False
                else:
                    self.logger.error("Failed to dismiss wdhandler dialog - process may be stuck")
                    return False

        # Handle Call bases dialog
        call_bases_found, call_bases_dialog = self._wait_for_dialog("call_bases")
        if call_bases_found and call_bases_dialog:
            self.logger.info("Found 'Call bases' dialog, clicking Yes...")
            success = self._click_dialog_button(call_bases_dialog, ["&Yes", "Yes"])
            if success:
                self.logger.info("Successfully clicked Yes on Call bases dialog")
            else:
                self.logger.warning("Failed to click Yes on Call bases dialog - this may prevent processing")
        else:
            self.logger.debug("No Call bases dialog found within timeout")

        # Handle Low quality files skipped dialog (may appear after Call bases)
        low_quality_found, low_quality_dialog = self._wait_for_dialog("low_quality")
        if low_quality_found and low_quality_dialog:
            self.logger.info("Found 'Low quality files skipped' dialog, clicking OK...")
            success = self._click_dialog_button(low_quality_dialog, ["OK"])
            if success:
                self.logger.info("Successfully dismissed Low quality files dialog")
            else:
                self.logger.warning("Failed to dismiss Low quality files dialog")
        else:
            self.logger.debug("No Low quality files dialog found within timeout")

        # Wait for processing to complete
        completion_success = self._wait_for_completion(folder_path)

        # Always close any Read information windows before returning
        self._close_all_read_info_dialogs()

        if completion_success:
            self.logger.info(f"Successfully processed {folder_path}")
        else:
            self.logger.warning(f"Processing may not have completed properly for {folder_path}")
            return False

        return True

    def _close_all_read_info_dialogs(self):
        """Close all Read information for... dialogs that might be open"""
        dialogs_closed = 0
        try:
            from pywinauto import findwindows

            # Find all Read information windows
            if self.app:
                read_windows = findwindows.find_elements(
                    title_re='Read information for.*',
                    process=self.app.process
                )

                for i, win in enumerate(read_windows):
                    try:
                        read_dialog = self.app.window(handle=win.handle)
                        read_dialog.close()
                        dialogs_closed += 1
                        time.sleep(0.1)  # Small delay between closes
                    except Exception as e:
                        self.logger.warning(f"Error closing read dialog {i+1}: {e}")

                if len(read_windows) > 0:
                    self.logger.info(f"Closed {dialogs_closed}/{len(read_windows)} read information dialogs")

                return len(read_windows) > 0
        except Exception as e:
            self.logger.warning(f"Error finding/closing read information dialogs: {e}")
            return False

    def _wait_for_dialog(self, dialog_type):
        """Wait for a specific dialog to appear and return both status and dialog object"""
        self.logger.debug(f"Waiting for {dialog_type} dialog...")
        timeout = self.timeouts.get(dialog_type, 5)

        try:
            if dialog_type == "browse_dialog":
                result = timings.wait_until(timeout=timeout, retry_interval=0.1,
                                        func=lambda: self.app.window(title_re='Browse.*Folder').exists(), #type: ignore
                                        value=True)
                if result:
                    dialog = self._get_browse_dialog()
                    return True, dialog
                return False, None
            elif dialog_type == "preferences":
                result = timings.wait_until(timeout=timeout, retry_interval=0.1,
                                        func=lambda: (self.app.window(title='Mseq Preferences').exists() or #type: ignore
                                                    self.app.window(title='mSeq Preferences').exists()), #type: ignore
                                        value=True)
                if result:
                    dialog = self._get_dialog_by_titles(['Mseq Preferences', 'mSeq Preferences'])
                    return True, dialog
                return False, None
            elif dialog_type == "copy_files":
                result = timings.wait_until(timeout=timeout, retry_interval=0.1,
                                        func=lambda: self.app.window(title_re='Copy.*sequence files').exists(), #type: ignore
                                        value=True)
                if result and self.app:
                    dialog = self.app.window(title_re='Copy.*sequence files')
                    return True, dialog
                return False, None
            elif dialog_type == "error_window":
                result = timings.wait_until(timeout=timeout, retry_interval=0.2,  # More frequent checks
                                        func=lambda: self.app.window(title_re='.*[Ee]rror.*').exists(), #type: ignore
                                        value=True)
                if result:
                    dialog = self._get_dialog_by_titles(['File error', 'Error'])
                    return True, dialog
                return False, None
            elif dialog_type == "call_bases":
                result = timings.wait_until(timeout=timeout, retry_interval=0.2,  # More frequent checks
                                        func=lambda: self.app.window(title_re='Call bases.*').exists(), #type: ignore
                                        value=True)
                if result and self.app:
                    dialog = self.app.window(title_re='Call bases.*')
                    return True, dialog
                return False, None
            elif dialog_type == "low_quality":
                result = timings.wait_until(timeout=timeout, retry_interval=0.3,
                                        func=lambda: self.app.window(title='Low quality files skipped').exists(), #type: ignore
                                        value=True)
                if result and self.app:
                    dialog = self.app.window(title='Low quality files skipped')
                    return True, dialog
                return False, None
            elif dialog_type == "read_info":
                result = timings.wait_until(timeout=timeout, retry_interval=0.1,
                                        func=lambda: self.app.window(title_re='Read information for.*').exists(), #type: ignore
                                        value=True)
                if result and self.app:
                    dialog = self.app.window(title_re='Read information for.*')
                    return True, dialog
                return False, None
            else:
                return False, None

        except timings.TimeoutError:
            return False, None

    def _get_browse_dialog(self):
        """Get browse dialog window with better reliability"""
        for title in ['Browse For Folder', 'Browse for Folder']:
            try:
                dialog = self.app.window(title=title) #type: ignore
                if dialog.exists():
                    return dialog
            except:
                pass

        # Last resort: try with regex
        try:
            return self.app.window(title_re='Browse.*Folder') #type: ignore
        except:
            return None

    def _get_dialog_by_titles(self, titles):
        """Try to get a dialog window by multiple possible titles"""
        if not self.app:
            return None

        for title in titles:
            try:
                if self.app is not None:
                    dialog = self.app.window(title=title) #type: ignore
                    if dialog.exists():
                        return dialog
            except Exception as e:
                self.logger.debug(f"Failed to find dialog with title '{title}': {e}")
                continue

        return None

    def _get_tree_view(self, dialog):
        """Get tree view control - simplified based on success path"""
        try:
            # Try most reliable method first
            tree = dialog.child_window(class_name="SysTreeView32")
            if tree.exists():
                return tree
        except:
            pass

        # Try with known titles as fallback
        for title in ["Choose project directory", "Navigation Pane", "Tree View"]:
            try:
                tree = dialog.child_window(title=title, class_name="SysTreeView32")
                if tree.exists():
                    return tree
            except:
                pass

        return None

    def _click_dialog_button(self, dialog, button_titles, max_retries=3):
        """Click a button in a dialog using multiple possible titles with simple retry logic"""
        for attempt in range(max_retries):
            for title in button_titles:
                try:
                    button = dialog.child_window(title=title, class_name="Button")
                    if button.exists() and button.is_enabled():
                        self.logger.debug(f"Found enabled button with title: {title} (attempt {attempt + 1}/{max_retries})")
                        button.click_input()
                        time.sleep(self.click_delay)

                        # Quick verification that dialog is gone
                        time.sleep(0.2)
                        if not dialog.exists():
                            self.logger.debug("Dialog dismissed successfully")
                            return True
                        else:
                            self.logger.debug("Dialog still exists after button click")
                            # If it's the last attempt, still consider it successful since we clicked the button
                            if attempt == max_retries - 1:
                                return True
                            # Otherwise, break out of title loop to retry
                            break
                except Exception as e:
                    self.logger.debug(f"Failed to click button '{title}': {e}")
                    continue

            # If we found and clicked a button but dialog still exists, wait briefly before retrying
            if attempt < max_retries - 1:
                time.sleep(0.3)

        self.logger.warning("Could not find any clickable button with provided titles")
        return False

    def _select_all_files_in_dialog(self, dialog):
        """Select all files in a dialog"""
        try:
            # Try Windows 10 approach first
            shell_view = dialog.child_window(title="ShellView", class_name="SHELLDLL_DefView")
            if shell_view.exists():
                list_view = shell_view.child_window(class_name="DirectUIHWND")
                if list_view.exists():
                    list_view.click_input()
                    send_keys('^a')  # Ctrl+A
                    return True
        except:
            pass

        try:
            # Try Windows 11 approach
            list_view = dialog.child_window(class_name="DirectUIHWND")
            if list_view.exists():
                list_view.click_input()
                send_keys('^a')  # Ctrl+A
                return True
        except:
            pass

        # Last resort - click in the middle and press Ctrl+A
        try:
            rect = dialog.rectangle()
            dialog.click_input(coords=((rect.right - rect.left) // 2, (rect.bottom - rect.top) // 2))
            time.sleep(self.click_delay)
            send_keys('^a')  # Ctrl+A
            return True
        except:
            return False

    def _navigate_folder_tree(self, dialog, path):
        """Navigate the folder tree - simplified based on success path"""
        dialog.set_focus()

        # Get tree view
        tree_view = self._get_tree_view(dialog)
        if not tree_view:
            self.logger.error("Could not find tree view control")
            return False

        # Convert Path object to string if necessary
        if hasattr(path, '__fspath__'):  # Check if it's a Path-like object
            path = str(path)

        # Parse path
        if ":" in path:
            parts = path.split("\\")
            drive = parts[0]  # e.g., "C:"
            folders = parts[1:] if len(parts) > 1 else []
        else:
            parts = path.split("\\")
            drive = "\\" + "\\".join(parts[:3])  # e.g., \\server\share
            folders = parts[3:] if len(parts) > 3 else []

        # Find Desktop in tree view roots
        desktop_item = None
        for item in tree_view.roots():
            if "Desktop" in item.text():
                desktop_item = item
                break

        if not desktop_item:
            self.logger.error("Could not find Desktop in tree view")
            return False

        # Click and expand Desktop
        desktop_item.click_input()
        time.sleep(self.click_delay)
        desktop_item.expand()
        time.sleep(self.expand_delay)

        # Find This PC under Desktop
        this_pc_item = None
        for child in desktop_item.children():
            if any(term in child.text().lower() for term in ["pc", "computer"]):
                this_pc_item = child
                break

        if not this_pc_item:
            self.logger.error("Could not find This PC under Desktop")
            return False

        # Click and expand This PC
        this_pc_item.click_input()
        time.sleep(self.click_delay)
        this_pc_item.expand()
        time.sleep(self.expand_delay)

        # Find the drive
        drive_item = None
        for child in this_pc_item.children():
            # Support multiple drive name formats
            drive_text = child.text()

            # Extract drive letter for better matching
            drive_letter = None
            if ":" in drive_text:
                drive_letter_parts = re.findall(r'([A-Za-z]:)', drive_text)
                if drive_letter_parts:
                    drive_letter = drive_letter_parts[0]

            # Try different matching approaches
            if (drive_text == drive or                      # Exact match
                drive in drive_text or                      # Contains match
                (drive_letter and drive_letter.upper() == drive.upper())):  # Drive letter match
                drive_item = child
                break

        if not drive_item:
            # Check for mapped network drives
            mapped_name = self.config.NETWORK_DRIVES.get(drive, None)
            if mapped_name:
                for child in this_pc_item.children():
                    if mapped_name in child.text():
                        drive_item = child
                        break

        if not drive_item:
            self.logger.error(f"Could not find drive {drive} in This PC")
            return False

        # Select the drive
        drive_item.click_input()
        time.sleep(self.click_delay)

        # If we only need to navigate to drive level, we're done
        if not folders:
            return True

        # Navigate through folder hierarchy
        current_item = drive_item
        for folder in folders:
            # Expand current folder
            current_item.expand()
            time.sleep(self.expand_delay)

            # Find the next folder
            next_item = None

            # Try exact match first
            for child in current_item.children():
                if child.text() == folder:
                    next_item = child
                    break

            # If not found, try partial match
            if not next_item:
                for child in current_item.children():
                    if folder.lower() in child.text().lower():
                        next_item = child
                        break

            if not next_item:
                self.logger.error(f"Could not find folder {folder}")
                return False

            # Select the folder
            next_item.click_input()
            time.sleep(self.click_delay)
            current_item = next_item

        # Make sure the final folder is selected
        current_item.click_input()
        return True

    def _wait_for_completion(self, folder_path):
        """Wait for mSeq processing to complete"""
        from pathlib import Path

        max_wait = self.timeouts["process_completion"]
        interval = 1.0
        elapsed = 0
        low_quality_handled = False

        # Reset the read info wait flag for this processing session
        if hasattr(self, '_read_info_wait_done'):
            delattr(self, '_read_info_wait_done')

        self.logger.info(f"Waiting for mSeq processing to complete (max: {max_wait}s)")
        self.logger.debug("Enhanced monitoring enabled for Low quality files skipped dialog")

        while elapsed < max_wait:
            # Check every 0.3 seconds for the blocking Low quality dialog (more frequent)
            if elapsed % 0.3 < interval:  # Check ~3 times per second
                self.logger.debug(f"Checking for Low quality dialog at {elapsed:.1f}s")
                low_quality_dialog = self._get_dialog_by_titles(['Low quality files skipped'])
                if low_quality_dialog and low_quality_dialog.exists():
                    self.logger.info("Found blocking 'Low quality files skipped' dialog during processing, clicking OK")
                    success = self._click_dialog_button(low_quality_dialog, ["OK"])
                    if success:
                        self.logger.info("Successfully clicked OK on Low quality dialog - waiting for processing to continue")
                        low_quality_handled = True
                        # Wait a moment for the read info dialog to appear and processing to start
                        time.sleep(0.5)
                    else:
                        self.logger.warning("Failed to click OK on Low quality dialog")

            # Check if read info dialog exists
            read_info_found, read_info_dialog = self._wait_for_dialog("read_info")

            if read_info_found:
                self.logger.debug(f"Read info dialog found at {elapsed:.1f}s")

                # Wait 0.5s after Read info dialog appears for text files to be created
                if not hasattr(self, '_read_info_wait_done'):
                    self.logger.debug("Read info dialog appeared, waiting 0.5s for text file creation")
                    time.sleep(0.5)
                    self._read_info_wait_done = True

                # Dialog exists, check for text files
                folder = Path(folder_path)
                txt_count = 0

                try:
                    for file_path in folder.iterdir():
                        if file_path.is_file():
                            # Check if file ends with any of the expected text file extensions
                            for ext in self.config.TEXT_FILES:
                                if file_path.name.endswith(ext):
                                    txt_count += 1
                                    break
                except Exception as e:
                    self.logger.error(f"Error reading folder contents: {e}")

                # Log progress every 10 seconds
                if elapsed % 10 == 0:
                    self.logger.info(f"Processing... found {txt_count}/5 text files (elapsed: {elapsed}s)")

                # Once we have all 5 text files, we're done
                if txt_count >= 5:
                    self.logger.info("All 5 text files found")

                    # Quick verification check (0.1s) to ensure files are stable
                    time.sleep(0.1)
                    final_txt_count = 0
                    try:
                        for file_path in folder.iterdir():
                            if file_path.is_file():
                                for ext in self.config.TEXT_FILES:
                                    if file_path.name.endswith(ext):
                                        final_txt_count += 1
                                        break
                    except Exception as e:
                        self.logger.error(f"Error in final file count: {e}")

                    if final_txt_count >= 5:
                        self.logger.info(f"Final confirmation: {final_txt_count} text files found, closing read dialogs")
                        dialogs_closed = self._close_all_read_info_dialogs()
                        if dialogs_closed:
                            self.logger.info("Processing completed successfully")
                        return True
                    else:
                        self.logger.debug(f"File count changed from {txt_count} to {final_txt_count}, continuing wait")
            else:
                # Read dialog not found
                if low_quality_handled:
                    self.logger.debug(f"Low quality handled but read info dialog not found yet at {elapsed:.1f}s - continuing to wait")
                else:
                    self.logger.debug(f"Read info dialog not found at {elapsed:.1f}s - checking for completion")
                    # Check if we have any text files even without the dialog
                    folder = Path(folder_path)
                    txt_count = 0
                    for file_path in folder.iterdir():
                        if file_path.is_file():
                            for ext in self.config.TEXT_FILES:
                                if file_path.name.endswith(ext):
                                    txt_count += 1
                                    break

                    # Only consider it complete if we have all 5 text files
                    # If we have exactly 4 files, the Low quality dialog is blocking the 5th (seq.info.txt)
                    if txt_count >= 5:
                        self.logger.info(f"Processing completed - found {txt_count} text files without dialog")
                        return True
                    elif txt_count == 4:
                        self.logger.info("Found exactly 4 text files - Low quality dialog is blocking seq.info.txt creation")
                        # Immediately check for Low quality dialog when we have exactly 4 files
                        low_quality_dialog = self._get_dialog_by_titles(['Low quality files skipped'])
                        if low_quality_dialog and low_quality_dialog.exists():
                            self.logger.info("Found blocking Low quality dialog with 4 files present, clicking OK")
                            success = self._click_dialog_button(low_quality_dialog, ["OK"])
                            if success:
                                self.logger.info("Successfully clicked OK on Low quality dialog - seq.info.txt should now be created")
                                low_quality_handled = True
                                time.sleep(0.5)  # Wait for seq.info.txt to be created
                            else:
                                self.logger.warning("Failed to click OK on Low quality dialog")
                        else:
                            self.logger.warning("Expected Low quality dialog with 4 files but dialog not found")
                    elif txt_count > 0 and txt_count < 4:
                        self.logger.debug(f"Found {txt_count} text files - processing still in progress")
                    elif elapsed > 5:  # Only warn about no files after some time has passed
                        self.logger.warning("No text files found after dialog disappeared")

            # Wait before checking again
            time.sleep(interval)
            elapsed += interval

        self.logger.warning(f"Processing timed out after {max_wait}s")
        return False

    def close(self):
        """Close the mSeq application"""
        if self.app:
            try:
                self.app.kill()
                self.logger.info("mSeq application closed")
            except Exception as e:
                self.logger.warning(f"Error closing mSeq: {e}")
                # Try alternative approach
                if self.main_window and self.main_window.exists():
                    try:
                        self.main_window.close()
                    except:
                        pass