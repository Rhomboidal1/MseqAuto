# ui_inspector.py - Enhanced UI inspection tool
import os
import time
from pywinauto import Application, Desktop
from pywinauto.keyboard import send_keys
import traceback

def save_windows_hierarchy(parent_window, output_file, level=0):
    """Save the entire window hierarchy to a file"""
    try:
        with open(output_file, 'a', encoding='utf-8') as f:
            indent = "  " * level
            f.write(f"{indent}Window: {parent_window.window_text()} ({parent_window.class_name()})\n")
            f.write(f"{indent}Rectangle: {parent_window.rectangle()}\n")
            
            # Get all child windows
            try:
                children = parent_window.children()
                if children:
                    f.write(f"{indent}Children:\n")
                    for child in children:
                        try:
                            save_windows_hierarchy(child, output_file, level + 1)
                        except Exception as e:
                            f.write(f"{indent}  Error getting child details: {e}\n")
                else:
                    f.write(f"{indent}No children\n")
            except Exception as e:
                f.write(f"{indent}Error getting children: {e}\n")
            
            f.write("\n")
    except Exception as e:
        print(f"Error saving window hierarchy: {e}")

def explore_mseq_app():
    """Connect to mSeq and explore its UI structure"""
    os.makedirs("ui_logs", exist_ok=True)
    output_file = os.path.join("ui_logs", f"mseq_ui_structure_{time.strftime('%Y%m%d_%H%M%S')}.txt")
    
    print(f"Exploring mSeq UI - output will be saved to {output_file}")
    
    try:
        # Try to connect to existing mSeq or start a new instance
        try:
            app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
            print("Connected to existing mSeq application")
        except Exception as e:
            print(f"Connection error: {e}")
            print("Starting new mSeq instance...")
            app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False)
            app.connect(title='mSeq', timeout=10)
        
        # Get the main window
        if app.window(title_re='mSeq*').exists():
            main_window = app.window(title_re='mSeq*')
        else:
            main_window = app.window(title_re='Mseq*')
        
        print(f"Main window found: {main_window.window_text()}")
        
        # Save the UI hierarchy
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("mSeq UI Structure\n")
            f.write("================\n\n")
            f.write(f"Date/Time: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        save_windows_hierarchy(main_window, output_file)
        
        # Now trigger the Browse For Folder dialog
        main_window.set_focus()
        print("Pressing Ctrl+N...")
        send_keys('^n')
        
        # Wait for Browse For Folder dialog
        print("Waiting for Browse For Folder dialog...")
        try:
            import pywinauto.timings as timings
            timings.wait_until(timeout=5, retry_interval=0.1, 
                             func=lambda: app.window(title='Browse For Folder').exists(), 
                             value=True)
            
            browse_dialog = app.window(title='Browse For Folder')
            print("Browse For Folder dialog found")
            
            # Save dialog details
            with open(output_file, 'a', encoding='utf-8') as f:
                f.write("\n\nBrowse For Folder Dialog\n")
                f.write("======================\n\n")
            
            save_windows_hierarchy(browse_dialog, output_file)
            
            # Cancel the dialog
            cancel_button = browse_dialog.child_window(title="Cancel", class_name="Button")
            if cancel_button.exists():
                cancel_button.click_input()
                print("Closed Browse For Folder dialog")
            
        except timings.TimeoutError:
            print("ERROR: Browse For Folder dialog did not appear")
            with open(output_file, 'a', encoding='utf-8') as f:
                f.write("\n\nERROR: Browse For Folder dialog did not appear\n")
        
        print(f"UI exploration completed. See {output_file} for details.")
        
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write(f"\n\nError: {e}\n")
            f.write(traceback.format_exc())

if __name__ == "__main__":
    explore_mseq_app()