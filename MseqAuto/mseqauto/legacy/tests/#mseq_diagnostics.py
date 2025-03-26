# mseq_diagnostics.py
import psutil
from pywinauto import Application
import sys
import time

def diagnose_window_properties():
    """Print diagnostic information about all windows potentially related to mSeq"""
    print("\n=== WINDOW DIAGNOSTIC INFORMATION ===")
    
    # Get all running processes with "j" or "mseq" in name
    mseq_processes = []
    for proc in psutil.process_iter(['pid', 'name', 'exe']):
        try:
            if proc.info['exe'] and ('j.exe' in proc.info['exe'].lower() or 'mseq' in proc.info['exe'].lower()):
                mseq_processes.append(proc)
                print(f"Found potential mSeq process: PID={proc.pid}, EXE={proc.info['exe']}")
        except Exception as e:
            print(f"Error checking process: {e}")
            pass
    
    # Try to find windows by process ID
    if mseq_processes:
        for proc in mseq_processes:
            try:
                app = Application(backend='win32').connect(process=proc.pid)
                windows = app.windows()
                for window in windows:
                    print(f"\nWindow details for PID {proc.pid}:")
                    print(f"  Title: {window.window_text()}")
                    print(f"  Class: {window.element_info.class_name}")
                    print(f"  Handle: {window.handle}")
                    print(f"  Rectangle: {window.rectangle()}")
                    try:
                        controls = window.children()
                        control_classes = set(c.element_info.class_name for c in controls)
                        print(f"  Control classes: {control_classes}")
                    except Exception as e:
                        print(f"  Could not retrieve controls: {e}")
            except Exception as e:
                print(f"Could not connect to process {proc.pid}: {e}")
    
    # Also search for any window with 'mseq' in title
    print("\nSearching for windows with 'mseq' in title:")
    try:
        app = Application()
        windows = []
        # Using Desktop().windows() to get all windows
        from pywinauto.application import Desktop
        desktop = Desktop(backend='win32')
        all_windows = desktop.windows()
        
        # Filter for mseq in title
        for window in all_windows:
            title = window.window_text().lower()
            if 'mseq' in title:
                windows.append(window)
                
        for window in windows:
            print(f"\nWindow with 'mseq' in title:")
            print(f"  Title: {window.window_text()}")
            print(f"  Class: {window.element_info.class_name}")
            print(f"  Handle: {window.handle}")
            print(f"  Rectangle: {window.rectangle()}")
            try:
                pid = window.process_id()
                proc = psutil.Process(pid)
                print(f"  Process: {proc.name()} (PID: {proc.pid})")
                print(f"  Executable: {proc.exe()}")
            except Exception as e:
                print(f"  Could not get process info: {e}")
    except Exception as e:
        print(f"Error searching for windows with 'mseq' in title: {e}")
    
    # Search for J software windows
    print("\nSearching for potential J Software windows:")
    try:
        # Look for windows with classes that might be related to J Software
        j_windows = desktop.windows(class_name_re='.*[Jj].*')
        for window in j_windows:
            # Skip explorer windows and other common windows
            if window.element_info.class_name in ['CabinetWClass', 'ExploreWClass', '#32770']:
                continue
                
            print(f"\nPotential J Software window found:")
            print(f"  Title: {window.window_text()}")
            print(f"  Class: {window.element_info.class_name}")
            print(f"  Handle: {window.handle}")
            try:
                pid = window.process_id()
                proc = psutil.Process(pid)
                print(f"  Process: {proc.name()} (PID: {proc.pid})")
                print(f"  Executable: {proc.exe()}")
            except Exception as e:
                print(f"  Could not get process info: {e}")
    except Exception as e:
        print(f"Error searching for J Software windows: {e}")
    
    print("=== END DIAGNOSTIC INFORMATION ===\n")

def main():
    print("Starting mSeq diagnostic tool")
    print("Please make sure mSeq is running before continuing.")
    input("Press Enter to continue...")
    
    diagnose_window_properties()
    
    print("\nDiagnostic complete. Results are displayed above.")

if __name__ == "__main__":
    main()