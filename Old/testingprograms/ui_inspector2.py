# ui_inspector.py - Enhanced version that properly saves window details
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
        # Start mSeq application
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
        
        # Press Ctrl+N to get Browse dialog
        main_window.set_focus()
        print("Pressing Ctrl+N...")
        send_keys('^n')
        
        # Wait for Browse For Folder dialog
        print("Waiting for Browse For Folder dialog...")
        import pywinauto.timings as timings
        timings.wait_until(timeout=5, retry_interval=0.1, 
                         func=lambda: app.window(title='Browse For Folder').exists(), 
                         value=True)
        
        browse_dialog = app.window(title='Browse For Folder')
        print("Browse For Folder dialog found")
        
        # Also save browse dialog details to the file
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write("\n\nBrowse For Folder Dialog\n")
            f.write("======================\n\n")
        
        save_windows_hierarchy(browse_dialog, output_file)
        
        # Print dialog info to console as well
        print("\nBrowse Dialog Structure:")
        browse_dialog.print_control_identifiers()
        
        # Get TreeView control
        tree_view = browse_dialog.child_window(class_name="SysTreeView32")
        
        # Log all top-level items in the tree to both console and file
        print("\nExploring TreeView structure:")
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write("\n\nTreeView Structure\n")
            f.write("=================\n\n")
            
            f.write("Root items:\n")
            print("Root items:")
            for root_item in tree_view.roots():
                print(f" - {root_item.text()}")
                f.write(f" - {root_item.text()}\n")
                
                # Try to expand and list children
                try:
                    root_item.expand()
                    time.sleep(0.5)
                    print(f"   Children of {root_item.text()}:")
                    f.write(f"   Children of {root_item.text()}:\n")
                    for child in root_item.children():
                        print(f"    - {child.text()}")
                        f.write(f"    - {child.text()}\n")
                        
                        # Try to expand one more level for things like This PC
                        if "PC" in child.text() or "Computer" in child.text():
                            try:
                                child.expand()
                                time.sleep(0.5)
                                print(f"      Children of {child.text()}:")
                                f.write(f"      Children of {child.text()}:\n")
                                for subchild in child.children():
                                    print(f"        - {subchild.text()}")
                                    f.write(f"        - {subchild.text()}\n")
                            except Exception as e:
                                print(f"      Error expanding {child.text()}: {e}")
                                f.write(f"      Error expanding {child.text()}: {e}\n")
                except Exception as e:
                    print(f"   Error expanding {root_item.text()}: {e}")
                    f.write(f"   Error expanding {root_item.text()}: {e}\n")
        
        # Cancel the dialog
        cancel_button = browse_dialog.child_window(title="Cancel", class_name="Button")
        if cancel_button.exists():
            cancel_button.click_input()
        
        print(f"UI exploration completed. See {output_file} for details.")
        
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        with open(output_file, 'a', encoding='utf-8') as f:
            f.write(f"\n\nError during exploration: {e}\n")
            f.write(traceback.format_exc())

if __name__ == "__main__":
    explore_mseq_app()