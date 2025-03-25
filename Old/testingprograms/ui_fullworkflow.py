import os
import time
from pywinauto import Application, timings
from pywinauto.keyboard import send_keys

def test_mseq_workflow(test_folder_path):
    """Test the full mSeq workflow with a test folder"""
    print(f"Testing mSeq workflow with folder: {test_folder_path}")
    
    try:
        # Try to connect to existing mSeq or start a new instance
        try:
            app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
            print("Connected to existing mSeq application")
        except:
            app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False)
            print("Started new mSeq application")
            app.connect(title='mSeq', timeout=10)
        
        # Get the main window
        main_window = app.window(title_re='mSeq*')
        if not main_window.exists():
            main_window = app.window(title_re='Mseq*')
        
        main_window.set_focus()
        print("Step 1: Pressing Ctrl+N to open New Project dialog...")
        send_keys('^n')
        
        # Handle Browse For Folder dialog
        print("Step 2: Waiting for Browse For Folder dialog...")
        timings.wait_until(timeout=5, retry_interval=0.1, 
                         func=lambda: app.window(title='Browse For Folder').exists(), 
                         value=True)
        
        browse_dialog = app.window(title='Browse For Folder')
        print("Browse For Folder dialog found")
        
        # First time browsing needs extra time
        time.sleep(1.2)
        
        # Navigate to the test folder
        print(f"Step 3: Navigating to folder: {test_folder_path}")
        
        # Print the TreeView structure for debugging
        tree_view = browse_dialog.child_window(class_name="SysTreeView32")
        
        # Get the virtual folder name
        import win32com.client
        shell = win32com.client.Dispatch("Shell.Application")
        namespace = shell.Namespace(0x11)  # CSIDL_DRIVES
        virtual_folder = namespace.Title
        print(f"Virtual folder name: {virtual_folder}")
        
        # Navigate through folder path
        desktop_path = f'\\Desktop\\{virtual_folder}'
        print(f"Starting navigation from: {desktop_path}")
        
        item = tree_view.get_item(desktop_path)
        folders = test_folder_path.split('\\')
        
        # Detailed navigation with logging
        for folder in folders:
            print(f"Looking for folder: {folder}")
            
            # Handle network drive mappings
            if 'P:' in folder:
                folder = folder.replace('P:', r'ABISync (P:)')
            if 'H:' in folder:
                import getpass
                username = getpass.getuser()
                folder = folder.replace('H:', f'{username} (\\\\w2k16\\users) (H:)')
            
            children_found = False
            for child in item.children():
                print(f" - Available child: {child.text()}")
                children_found = True
                if child.text() == folder:
                    print(f"   Found match: {folder}")
                    browse_dialog.set_focus()
                    item = child
                    item.click_input()
                    time.sleep(0.3)  # Give UI time to update
                    break
            
            if not children_found:
                print(f"No children found for {item.text()}")
        
        print("Step 4: Clicking OK button...")
        ok_button = browse_dialog.child_window(title="OK", class_name="Button")
        if ok_button.exists():
            ok_button.click_input()
        else:
            print("ERROR: OK button not found!")
        
        # Handle Mseq Preferences dialog
        print("Step 5: Waiting for Mseq Preferences dialog...")
        timings.wait_until(timeout=5, retry_interval=0.1, 
                         func=lambda: app.window(title='Mseq Preferences').exists(), 
                         value=True)
        
        pref_window = app.window(title='Mseq Preferences')
        print("Mseq Preferences dialog found")
        ok_button = pref_window.child_window(title="&OK", class_name="Button")
        ok_button.click_input()
        
        # Handle Copy sequence files dialog
        print("Step 6: Waiting for Copy sequence files dialog...")
        timings.wait_until(timeout=5, retry_interval=0.1, 
                         func=lambda: app.window(title='Copy sequence files').exists(), 
                         value=True)
        
        copy_files_window = app.window(title='Copy sequence files')
        print("Copy sequence files dialog found")
        
        shell_view = copy_files_window.child_window(title="ShellView", class_name="SHELLDLL_DefView")
        list_view = shell_view.child_window(class_name="DirectUIHWND")
        list_view.click_input()
        
        print("Selecting all files with Ctrl+A...")
        send_keys('^a')
        
        open_button = copy_files_window.child_window(title="&Open", class_name="Button")
        open_button.click_input()
        
        # Handle File error dialog
        print("Step 7: Waiting for File error dialog...")
        timings.wait_until(timeout=20, retry_interval=0.5, 
                         func=lambda: app.window(title='File error').exists(), 
                         value=True)
        
        error_window = app.window(title='File error')
        print("File error dialog found")
        error_ok_button = error_window.child_window(class_name="Button")
        error_ok_button.click_input()
        
        # Handle Call bases dialog
        print("Step 8: Waiting for Call bases dialog...")
        timings.wait_until(timeout=10, retry_interval=0.3, 
                         func=lambda: app.window(title='Call bases?').exists(), 
                         value=True)
        
        call_bases_window = app.window(title='Call bases?')
        print("Call bases dialog found")
        yes_button = call_bases_window.child_window(title="&Yes", class_name="Button")
        yes_button.click_input()
        
        # Wait for processing to complete
        print("Step 9: Waiting for processing to complete...")
        
        def process_complete():
            if app.window(title="Low quality files skipped").exists():
                return True
            
            # Check for output files
            count = 0
            for item in os.listdir(test_folder_path):
                if os.path.isfile(os.path.join(test_folder_path, item)):
                    if (item.endswith('.raw.qual.txt') or 
                        item.endswith('.raw.seq.txt') or 
                        item.endswith('.seq.info.txt') or 
                        item.endswith('.seq.qual.txt') or 
                        item.endswith('.seq.txt')):
                        count += 1
            return count == 5
        
        timings.wait_until(timeout=45, retry_interval=0.2, 
                         func=process_complete, 
                         value=True)
        
        # Handle Low quality files skipped dialog if it appears
        if app.window(title="Low quality files skipped").exists():
            print("Step 10: Handling Low quality files skipped dialog...")
            low_quality_window = app.window(title='Low quality files skipped')
            ok_button = low_quality_window.child_window(class_name="Button")
            ok_button.click_input()
        
        # Handle Read information dialog
        print("Step 11: Waiting for Read information dialog...")
        timings.wait_until(timeout=5, retry_interval=0.1, 
                         func=lambda: app.window(title_re='Read information for*').exists(), 
                         value=True)
        
        if app.window(title_re='Read information for*').exists():
            read_window = app.window(title_re='Read information for*')
            read_window.close()
        
        print("Workflow test completed successfully!")
        
    except Exception as e:
        print(f"Error in test_mseq_workflow: {e}")

if __name__ == "__main__":
    # Specify a test folder path that exists on your system
    test_folder = r"P:\Data\Testing\Win11Test\BioI-21845\BioI-21845_Mollafarajzadeh_147546"  # Example - change to a real path
    test_mseq_workflow(test_folder)