import os
import time
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys

def save_window_info(window, filename):
    """Save window control identifiers to a file"""
    try:
        # Create logs directory if it doesn't exist
        os.makedirs("ui_logs", exist_ok=True)
        
        # Redirect stdout to file temporarily to capture control_identifiers output
        with open(os.path.join("ui_logs", filename), 'w') as f:
            import sys
            original_stdout = sys.stdout
            sys.stdout = f
            
            # Print information about the window
            print(f"Window Title: {window.window_text()}")
            print(f"Window Class: {window.class_name()}")
            print(f"Window Rectangle: {window.rectangle()}")
            print("\nControl Identifiers:")
            window.print_control_identifiers()
            
            # Restore stdout
            sys.stdout = original_stdout
            
        print(f"Window info saved to ui_logs/{filename}")
    except Exception as e:
        print(f"Error saving window info: {e}")

def capture_all_windows():
    """Capture all visible windows on the desktop"""
    all_windows = Desktop(backend="win32").windows()
    for i, window in enumerate(all_windows):
        try:
            title = window.window_text() or f"Untitled_{i}"
            # Replace invalid filename characters
            safe_title = "".join(c for c in title if c.isalnum() or c in " _-").strip()
            if not safe_title:
                safe_title = f"Window_{i}"
            filename = f"{safe_title}_{window.class_name()}.txt"
            save_window_info(window, filename)
        except Exception as e:
            print(f"Error capturing window {i}: {e}")

def explore_mseq_dialogs():
    """Start mSeq and explore each dialog in the workflow"""
    try:
        # Try to connect to existing mSeq or start a new instance
        try:
            app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
            print("Connected to existing mSeq application")
        except:
            app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False)
            print("Started new mSeq application")
            app.connect(title='mSeq', timeout=10)
        
        # Capture the main window
        main_window = app.window(title_re='mSeq*')
        if not main_window.exists():
            main_window = app.window(title_re='Mseq*')
        
        save_window_info(main_window, "mSeq_MainWindow.txt")
        
        # Trigger the New Project dialog (Ctrl+N)
        main_window.set_focus()
        print("Pressing Ctrl+N to open New Project dialog...")
        send_keys('^n')
        time.sleep(1)
        
        # Try to find the Browse For Folder dialog
        browse_dialog = app.window(title='Browse For Folder')
        if browse_dialog.exists():
            print("Browse For Folder dialog found")
            save_window_info(browse_dialog, "Browse_For_Folder_Dialog.txt")
            
            # Click Cancel to close it
            cancel_button = browse_dialog.child_window(title="Cancel", class_name="Button")
            if cancel_button.exists():
                cancel_button.click_input()
                print("Clicked Cancel on Browse For Folder dialog")
        else:
            print("Browse For Folder dialog not found")
        
        # Capture any other dialogs that might be open
        dialogs = [
            (app.window(title='Mseq Preferences'), "Mseq_Preferences_Dialog.txt"),
            (app.window(title='Copy sequence files'), "Copy_Sequence_Files_Dialog.txt"),
            (app.window(title='File error'), "File_Error_Dialog.txt"),
            (app.window(title='Call bases?'), "Call_Bases_Dialog.txt"),
            (app.window(title='Low quality files skipped'), "Low_Quality_Files_Dialog.txt")
        ]
        
        for dialog, filename in dialogs:
            if dialog.exists():
                print(f"Found dialog: {dialog.window_text()}")
                save_window_info(dialog, filename)
                
                # Try to close the dialog
                try:
                    close_button = dialog.child_window(title="Cancel", class_name="Button")
                    if close_button.exists():
                        close_button.click_input()
                    else:
                        close_button = dialog.child_window(title="OK", class_name="Button")
                        if close_button.exists():
                            close_button.click_input()
                    print(f"Closed dialog: {dialog.window_text()}")
                except:
                    print(f"Couldn't close dialog: {dialog.window_text()}")
    
    except Exception as e:
        print(f"Error in explore_mseq_dialogs: {e}")

if __name__ == "__main__":
    print("Starting UI exploration...")
    
    # Option 1: Capture all windows on the desktop
    # capture_all_windows()
    
    # Option 2: Explore mSeq dialogs specifically
    explore_mseq_dialogs()
    
    print("UI exploration completed")