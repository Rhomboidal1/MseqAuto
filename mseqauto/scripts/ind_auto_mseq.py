# ind_auto_mseq.py
import os
import sys
import tkinter as tk
from tkinter import filedialog
import subprocess
import re


def get_folder_from_user():
    """Simple folder selection dialog that works reliably"""
    print("Opening folder selection dialog...")
    root = tk.Tk()
    root.withdraw()
    root.update()  # Force an update to ensure dialog shows

    folder_path = filedialog.askdirectory(
        title="Select today's data folder to mseq orders",
        mustexist=True
    )

    root.destroy()
    return folder_path


def main():
    # Get folder path FIRST before any package imports
    data_folder = get_folder_from_user()

    if not data_folder:
        print("No folder selected, exiting")
        return

    # NOW import package modules
    from mseqauto.utils import setup_logger
    from mseqauto.config import MseqConfig
    from mseqauto.core import OSCompatibilityManager, FileSystemDAO, MseqAutomation, FolderProcessor

    # Setup logger
    logger = setup_logger("ind_auto_mseq")
    logger.info("Starting IND auto mSeq...")

    # Log that we already selected folder
    logger.info(f"Using folder: {data_folder}")
    data_folder = re.sub(r'/', '\\\\', data_folder)

    # Check for 32-bit Python requirement
    OSCompatibilityManager.py32_check(
        script_path=__file__,
        logger=logger
    )

    # Log OS environment information
    OSCompatibilityManager.log_environment_info(logger)

    # Initialize components
    logger.info("Initializing components...")
    config = MseqConfig()
    logger.info("Config loaded")

    # Use OS compatibility manager for timeouts
    logger.info("Using OS-specific timeouts")

    file_dao = FileSystemDAO(config)
    logger.info("FileSystemDAO initialized")

    ui_automation = MseqAutomation(config)
    logger.info("UI Automation initialized with OS-specific settings")

    processor = FolderProcessor(file_dao, ui_automation, config, logger=logger.info)
    logger.info("Folder processor initialized")

    # Run batch file to generate order key
    try:
        logger.info(f"Running batch file: {config.BATCH_FILE_PATH}")
        subprocess.run(config.BATCH_FILE_PATH, shell=True, check=True)
        logger.info("Batch file completed successfully")
    except subprocess.CalledProcessError:
        logger.error(f"Batch file {config.BATCH_FILE_PATH} failed to run")
        print(f"Error: Batch file {config.BATCH_FILE_PATH} failed to run")
        return

    def faster_navigate_folder_tree(self, dialog, path):
        """Patched method with faster navigation - no edit box attempt"""
        import logging
        logger = logging.getLogger(__name__)

        logger.info(f"Fast navigation to: {path}")
        dialog.set_focus()

        # Get tree view
        tree_view = self._get_tree_view(dialog)
        if not tree_view:
            logger.error("Could not find TreeView control")
            return False

        # Standard navigation with reduced delays
        try:
            # Parse the path
            if ":" in path:
                parts = path.split("\\")
                drive = parts[0]
                folders = parts[1:] if len(parts) > 1 else []
            else:
                parts = path.split("\\")
                drive = "\\" + "\\".join(parts[:3])
                folders = parts[3:] if len(parts) > 3 else []

            # Find Desktop
            for root_item in tree_view.roots():
                if "Desktop" in root_item.text():
                    desktop_item = root_item
                    break
            else:
                try:
                    desktop_item = tree_view.get_item('\\Desktop')
                except:
                    logger.error("Could not find Desktop")
                    return False

            # Find This PC
            desktop_item.click_input()
            desktop_item.expand()

            this_pc_item = None
            for child in desktop_item.children():
                if any(pc_name in child.text() for pc_name in ["PC", "Computer"]):
                    this_pc_item = child
                    break

            if not this_pc_item:
                logger.error("Could not find This PC")
                return False

            # Expand This PC
            this_pc_item.click_input()
            this_pc_item.expand()

            # Find drive
            drive_found = False
            mapped_name = self.config.NETWORK_DRIVES.get(drive, None)

            for item in this_pc_item.children():
                drive_text = item.text()
                if drive in drive_text or (mapped_name and mapped_name in drive_text):
                    item.click_input()
                    current_item = item
                    drive_found = True
                    break

            if not drive_found:
                logger.error(f"Could not find drive {drive}")
                return False

            # Navigate folders with minimal delay
            for folder in folders:
                current_item.expand()

                folder_found = False
                for child in current_item.children():
                    if child.text() == folder:
                        child.click_input()
                        current_item = child
                        folder_found = True
                        break

                if not folder_found:
                    logger.error(f"Could not find folder {folder}")
                    return False

            return True

        except Exception as e:
            logger.error(f"Error in fast navigation: {e}")
            return False

    # Apply the patch to the MSeqAutomation class
    MseqAutomation.navigate_folder_tree = faster_navigate_folder_tree

    # Create a patched process_folder method with less waiting
    def faster_process_folder(self, folder_path):
        """Patched method with faster processing and reduced waits"""
        if not os.path.exists(folder_path):
            print(f"Warning: Folder does not exist: {folder_path}")
            return False

        # Check for .ab1 files
        ab1_files = [f for f in os.listdir(folder_path) if f.endswith('.ab1')]
        if not ab1_files:
            print(f"No .ab1 files found in {folder_path}, skipping processing")
            return False

        self.app, self.main_window = self.connect_or_start_mseq()
        self.main_window.set_focus()
        from pywinauto.keyboard import send_keys
        send_keys('^n')  # Ctrl+N for new project

        # Wait for Browse dialog
        self.wait_for_dialog("browse_dialog")
        dialog_window = self._get_browse_dialog()

        if not dialog_window:
            print("Failed to find Browse For Folder dialog")
            return False

        # Navigate to folder with faster method
        navigate_success = self.navigate_folder_tree(dialog_window, folder_path)

        # Click OK
        try:
            ok_button = dialog_window.child_window(title="OK", class_name="Button")
            if ok_button.exists():
                ok_button.click_input()
            else:
                dialog_window.set_focus()
                send_keys('{ENTER}')
        except:
            dialog_window.set_focus()
            send_keys('{ENTER}')

        # Handle preferences
        self.wait_for_dialog("preferences")
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

        # Handle Copy files
        self.wait_for_dialog("copy_files")
        try:
            copy_files_window = self.app.window(title_re='Copy.*sequence files')

            # Try to find list view and select all
            try:
                shell_view = copy_files_window.child_window(title="ShellView")
                list_view = shell_view.child_window(class_name="DirectUIHWND")
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

        # Handle Error dialog
        self.wait_for_dialog("error_window")
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

        # Handle Call bases
        self.wait_for_dialog("call_bases")
        try:
            call_bases_window = self.app.window(title_re='Call bases.*')

            if call_bases_window and call_bases_window.exists():
                for btn_title in ["&Yes", "Yes"]:
                    try:
                        yes_button = call_bases_window.child_window(title=btn_title)
                        if yes_button.exists():
                            yes_button.click_input()
                            break
                    except:
                        pass
                else:
                    call_bases_window.set_focus()
                    send_keys('{ENTER}')
        except:
            pass

        # Wait for completion - use a more aggressive check
        import time
        max_wait = self.config.TIMEOUTS["process_completion"]
        interval = 0.5
        elapsed = 0

        while elapsed < max_wait:
            # Check if process completed
            completed = False

            # Check for low quality dialog
            try:
                for title in ["Low quality files skipped", "Quality files skipped"]:
                    if self.app.window(title=title).exists():
                        low_quality_window = self.app.window(title=title)
                        ok_button = low_quality_window.child_window(class_name="Button")
                        ok_button.click_input()
                        completed = True
                        break
            except:
                pass

            # Check for read info dialog
            try:
                if self.app.window(title_re='Read information for*').exists():
                    read_window = self.app.window(title_re='Read information for*')
                    read_window.close()
                    completed = True
            except:
                pass

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
                completed = True

            if completed:
                return True

            # Get appropriate interval from the compatibility manager
            wait_interval = OSCompatibilityManager.get_timeout("polling_interval", interval)
            time.sleep(wait_interval)
            elapsed += wait_interval  # Make sure to update elapsed with the actual wait time

        print(f"Warning: Timeout waiting for processing to complete for {folder_path}")
        print("This may be normal if the folder has already been processed or has special files")
        return True  # Return True anyway to continue with next folder

    # Apply the second patch
    MseqAutomation.process_folder = faster_process_folder

    ui_automation = MseqAutomation(config)
    logger.info("UI Automation initialized with optimized functions")

    from mseqauto.core import  FolderProcessor
    processor = FolderProcessor(file_dao, ui_automation, config, logger=logger.info)
    logger.info("Folder processor initialized")

    # Run batch file to generate order key
    try:
        logger.info(f"Running batch file: {config.BATCH_FILE_PATH}")
        subprocess.run(config.BATCH_FILE_PATH, shell=True, check=True)
        logger.info("Batch file completed successfully")
    except subprocess.CalledProcessError:
        logger.error(f"Batch file {config.BATCH_FILE_PATH} failed to run")
        print(f"Error: Batch file {config.BATCH_FILE_PATH} failed to run")
        return

    try:
        # Process BioI folders
        bio_folders = file_dao.get_folders(data_folder, r'bioi-\d+')
        logger.info(f"Found {len(bio_folders)} BioI folders")

        immediate_orders = file_dao.get_folders(data_folder, r'bioi-\d+_.+_\d+')
        logger.info(f"Found {len(immediate_orders)} immediate order folders")

        pcr_folders = file_dao.get_folders(data_folder, r'fb-pcr\d+_\d+')
        logger.info(f"Found {len(pcr_folders)} PCR folders")

        # Process BioI folders
        for i, folder in enumerate(bio_folders):
            logger.info(f"Processing BioI folder {i + 1}/{len(bio_folders)}: {os.path.basename(folder)}")
            processor.process_bio_folder(folder)

        # Determine if we're processing the IND Not Ready folder
        is_ind_not_ready = os.path.basename(data_folder) == config.IND_NOT_READY_FOLDER
        logger.info(f"Is IND Not Ready folder: {is_ind_not_ready}")

        # Process immediate orders
        for i, folder in enumerate(immediate_orders):
            logger.info(f"Processing order folder {i + 1}/{len(immediate_orders)}: {os.path.basename(folder)}")
            processor.process_order_folder(folder, data_folder)

        # Process PCR folders
        for i, folder in enumerate(pcr_folders):
            logger.info(f"Processing PCR folder {i + 1}/{len(pcr_folders)}: {os.path.basename(folder)}")
            processor.process_pcr_folder(folder)

        logger.info("All processing completed")
        print("")
        print("ALL DONE")
    except Exception as e:
        import traceback
        logger.error(f"Unexpected error: {e}")
        logger.error(traceback.format_exc())
        print(f"Unexpected error: {e}")
    finally:
        # Close mSeq application
        try:
            if ui_automation is not None:
                logger.info("Closing mSeq application")
                ui_automation.close()
        except Exception as e:
            logger.error(f"Error closing mSeq: {e}")

if __name__ == "__main__":
    main()