import os
import time
from pywinauto import Desktop
from pywinauto.application import Application

def dialog_explorer():
    """Wait for new dialog windows to appear and capture their details to a single log file"""
    print("Dialog Explorer started. Press Ctrl+C to stop.")
    print("Open dialogs in mSeq to capture their details...")
    
    # Create logs directory
    os.makedirs("dialog_logs", exist_ok=True)
    
    # Create a single log file with timestamp
    log_filename = f"dialog_explorer_log_{time.strftime('%Y%m%d_%H%M%S')}.txt"
    log_filepath = os.path.join("dialog_logs", log_filename)
    
    print(f"All dialog details will be written to: {log_filepath}")
    
    # Get initial set of windows
    initial_windows = set(window.window_text() for window in Desktop(backend="win32").windows())
    
    # Write header to log file
    with open(log_filepath, 'w') as logfile:
        logfile.write("=== Dialog Explorer Log ===\n")
        logfile.write(f"Started: {time.strftime('%Y-%m-%d %H:%M:%S')}\n\n")
    
    try:
        while True:
            # Get current windows
            current_windows = set(window.window_text() for window in Desktop(backend="win32").windows())
            
            # Find new windows
            new_windows = current_windows - initial_windows
            if new_windows:
                print(f"New windows detected: {new_windows}")
                
                # Open log file in append mode
                with open(log_filepath, 'a') as logfile:
                    # Capture details of new windows
                    for window_text in new_windows:
                        if window_text:  # Skip empty titles
                            try:
                                # Find the window by its text
                                window = Desktop(backend="win32").window(title=window_text)
                                
                                # Write a separator
                                logfile.write("\n" + "="*80 + "\n")
                                logfile.write(f"DIALOG DETECTED: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                                logfile.write("="*80 + "\n\n")
                                
                                # Save window details
                                logfile.write(f"Window Title: {window.window_text()}\n")
                                logfile.write(f"Window Class: {window.class_name()}\n")
                                logfile.write(f"Window Rectangle: {window.rectangle()}\n\n")
                                logfile.write("Controls:\n")
                                
                                # Get all controls with detailed properties
                                logfile.write("ID | Text | Class Name | Rectangle | Control Type | Automation ID\n")
                                logfile.write("-"*80 + "\n")
                                
                                for i, control in enumerate(window.descendants()):
                                    try:
                                        # Get more detailed information about each control
                                        text = control.window_text() or "[No Text]"
                                        class_name = control.class_name() or "[No Class]"
                                        rect = control.rectangle()
                                        
                                        # Try to get additional properties
                                        try:
                                            control_type = control.control_type() if hasattr(control, 'control_type') else "[Unknown]"
                                        except:
                                            control_type = "[Error]"
                                            
                                        try:
                                            automation_id = control.automation_id() if hasattr(control, 'automation_id') else "[Unknown]"
                                        except:
                                            automation_id = "[Error]"
                                        
                                        logfile.write(f"{i:02d} | {text} | {class_name} | {rect} | {control_type} | {automation_id}\n")
                                    except Exception as e:
                                        logfile.write(f"{i:02d} | [Error getting control details: {e}]\n")
                                
                                # Add specific check for TreeView controls
                                logfile.write("\nTreeView Controls:\n")
                                logfile.write("-"*80 + "\n")
                                
                                tree_views = [ctrl for ctrl in window.descendants() if ctrl.class_name() == "SysTreeView32"]
                                if tree_views:
                                    for i, tree in enumerate(tree_views):
                                        logfile.write(f"TreeView #{i+1}: {tree.window_text()} (Class: {tree.class_name()})\n")
                                        logfile.write(f"Rectangle: {tree.rectangle()}\n")
                                        
                                        # Try to get additional IDs
                                        try:
                                            logfile.write(f"Control ID: {tree.control_id()}\n")
                                        except:
                                            logfile.write("Control ID: [Error]\n")
                                            
                                        # Add extra details about this specific TreeView
                                        logfile.write("Getting children information...\n")
                                        try:
                                            # Try to get root items
                                            try:
                                                roots = list(tree.roots())
                                                logfile.write(f"Root items count: {len(roots)}\n")
                                                for j, root in enumerate(roots):
                                                    logfile.write(f"  Root {j+1}: {root.text()}\n")
                                            except Exception as e:
                                                logfile.write(f"Error getting roots: {e}\n")
                                        except Exception as e:
                                            logfile.write(f"Error exploring TreeView: {e}\n")
                                else:
                                    logfile.write("No TreeView controls found in this dialog.\n")
                                
                                print(f"Captured details for '{window_text}'")
                            except Exception as e:
                                logfile.write(f"\nError capturing window '{window_text}': {e}\n")
                                print(f"Error capturing window '{window_text}': {e}")
                
                # Update initial windows
                initial_windows = current_windows
            
            # Wait before checking again
            time.sleep(0.5)
    
    except KeyboardInterrupt:
        print("\nDialog Explorer stopped.")
        with open(log_filepath, 'a') as logfile:
            logfile.write(f"\n\nDialog Explorer stopped: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")

if __name__ == "__main__":
    dialog_explorer()