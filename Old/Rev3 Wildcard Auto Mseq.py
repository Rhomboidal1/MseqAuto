from tkinter import filedialog
from os import listdir, path as OsPath
from re import sub
from time import sleep
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
from pywinauto.keyboard import send_keys
from pywinauto import timings

firstTimeBrowsingMseq = True

#function gets the virtual folder name, often named This PC, but could be different per machine.
def get_virtual_folder_name():
    import win32com.client
    shell = win32com.client.Dispatch("Shell.Application")
    namespace = shell.Namespace(0x11)  # CSIDL_DRIVES
    folder_name = namespace.Title
    return folder_name


def GetPlateFolders(path):
    folderPaths = []
    for item in listdir(path):
        if OsPath.isdir(f'{path}\\{item}'):
                folderPaths.append(f'{path}\\{item}')
    return folderPaths

#function looks for any fsa files in the folder. If there are, the folder is not to be Mseqed.
def CheckForfsaFiles(folder):
    fsa = False
    for item in listdir(folder):
        if item.endswith('.fsa'):
            fsa =True
            break
    return fsa
    

def CheckForMseq(orderfolder):
    wasMSeqed = False
    currentProj = []
    mseqSet = {'chromat_dir', 'edit_dir', 'phd_dir', 'mseq4.ini'}
    for item in listdir(orderfolder):
        if item in mseqSet:
            currentProj.append(item)
    projSet = set(currentProj)
    if mseqSet == projSet: 
        wasMSeqed = True
    return wasMSeqed

def CheckFor5txt(orderPath):
    fivetxts = 0
    all5 = False
    for item in listdir(orderPath):
        if OsPath.isfile(f'{orderPath}\\{item}'):
            #print(item)
            if item.endswith('.raw.qual.txt'):
                fivetxts +=1
            elif item.endswith('.raw.seq.txt'):
                fivetxts +=1
            elif item.endswith('.seq.info.txt'):
                fivetxts +=1
            elif item.endswith('.seq.qual.txt'):
                fivetxts +=1
            elif item.endswith('.seq.txt'):
                fivetxts +=1
    all5 = True if (fivetxts == 5) else False
    return all5


#function returns both the app object and its main window.
def GetMseq():
    from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
    from pywinauto import Application, timings
    
    print("Attempting to connect to or start mSeq...")
    
    try:
        print("Trying to connect to 'Mseq*'")
        app = Application(backend='win32').connect(title_re='Mseq*', timeout=1)
    except (ElementNotFoundError, timings.TimeoutError):
        try:
            print("Trying to connect to 'mSeq*'")
            app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
        except (ElementNotFoundError, timings.TimeoutError):
            print("No existing mSeq found, starting a new instance...")
            # This is the key line that works in the original version - keep it as one chained call
            app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False).connect(title='mSeq', timeout=10)
        except ElementAmbiguousError:
            print("Multiple 'mSeq*' windows found, connecting to first one")
            app = Application(backend='win32').connect(title_re='mSeq*', found_index=0, timeout=1)
            app2 = Application(backend='win32').connect(title_re='mSeq*', found_index=1, timeout=1)
            app2.kill()
    except ElementAmbiguousError:
        print("Multiple 'Mseq*' windows found, connecting to first one")
        app = Application(backend='win32').connect(title_re='Mseq*', found_index=0, timeout=1)
        app2 = Application(backend='win32').connect(title_re='Mseq*', found_index=1, timeout=1)
        app2.kill()
    
    print("Finding main window...")
    if app.window(title_re='mSeq*').exists()==False:
        mainWindow = app.window(title_re='Mseq*')
        print("Using 'Mseq*' window")
    else:
        mainWindow = app.window(title_re='mSeq*')
        print("Using 'mSeq*' window")

    print("mSeq started successfully")
    return app, mainWindow
    
def IsBrowseDialogOpen(myapp):
    # Check for different possible dialog titles in Windows 11
    return (myapp.window(title='Browse For Folder').exists() or 
            myapp.window(title='Select Folder').exists() or
            myapp.window(title='Choose Folder').exists())

def IsPreferencesOpen(myapp):
    return (myapp.window(title='Mseq Preferences').exists() or 
            myapp.window(title='mSeq Preferences').exists() or
            myapp.window(title='MSEQ Preferences').exists())

def IsCopyFilesOpen(myapp):
    return (myapp.window(title='Copy sequence files').exists() or 
            myapp.window(title='Copy Sequence Files').exists() or 
            myapp.window(title='Select Files').exists())

def IsErrorWindowOpen(myapp):
    return (myapp.window(title='File error').exists() or 
            myapp.window(title='File Error').exists() or
            myapp.window(title='Error').exists())

def IsCallBasesOpen(myapp):
    return (myapp.window(title='Call bases?').exists() or 
            myapp.window(title='Call Bases?').exists())

def IsReadInfoOpen(myapp):
    return (myapp.window(title_re='Read information for*').exists() or 
            myapp.window(title_re='Read Information For*').exists())

def FivetxtORlowQualityWindow(myApp, orderpath):
    #from pywinauto.application import Application
    #app = GetMseq()
    if myApp.window(title="Low quality files skipped").exists():
        return True
    if CheckFor5txt(orderpath):
        return True
    return False

def click_button_by_text(window, button_text, fallback_classes=None):
    """Try multiple methods to find and click a button."""
    if fallback_classes is None:
        fallback_classes = ["Button"]
    
    # Try by exact text
    for cls in fallback_classes:
        try:
            button = window.child_window(title=button_text, class_name=cls)
            button.click_input()
            return True
        except:
            pass
    
    # Try by case-insensitive text
    for cls in fallback_classes:
        try:
            buttons = window.children(class_name=cls)
            for button in buttons:
                if button.texts() and button_text.lower() in button.texts()[0].lower():
                    button.click_input()
                    return True
        except:
            pass
    
    # Try by partial match
    for cls in fallback_classes:
        try:
            buttons = window.children(class_name=cls)
            for button in buttons:
                if button.texts() and any(partial in button.texts()[0].lower() for partial in button_text.lower().split()):
                    button.click_input()
                    return True
        except:
            pass
    
    return False

def MseqOrder(orderpath):
    import time
    from pywinauto.keyboard import send_keys
    from pywinauto import timings
    
    # Add an overall timeout for the entire operation
    start_time = time.time()
    max_operation_time = 180  # 3 minutes
    
    try:
        app, mseqMainWindow = GetMseq()
        mseqMainWindow.set_focus()
        send_keys('^n')

        # Wait for Browse dialog with retry
        timings.wait_until(timeout=5, retry_interval=0.1, func=lambda: IsBrowseDialogOpen(app), value=True)
        
        dialogWindow = app.window(title='Browse For Folder')
        if not dialogWindow.exists():
            # Try alternative titles in Windows 11
            for title in ['Select Folder', 'Choose Folder']:
                dialogWindow = app.window(title=title)
                if dialogWindow.exists():
                    break
            else:
                raise Exception("Could not find browse dialog window")
        
        global firstTimeBrowsingMseq
        if firstTimeBrowsingMseq:
            firstTimeBrowsingMseq = False
            sleep(1.2)  # mseq and file system is sleepy, ok.. needs some time to wake.
        else:
            sleep(0.5)
        
        nav_success = NavigateToFolder(dialogWindow, orderpath)
        if not nav_success:
            print(f"Failed to navigate to folder: {orderpath}")
            raise Exception(f"Could not navigate to {orderpath}")
        
        # Try multiple ways to find and click OK button
        if not click_button_by_text(dialogWindow, "OK", ["Button"]):
            # If OK button not found, try alternatives
            for button_text in ["Select Folder", "Open", "Choose"]:
                if click_button_by_text(dialogWindow, button_text, ["Button"]):
                    break
            else:
                # Last resort - try to find any button that might be the OK button
                buttons = dialogWindow.children(class_name="Button")
                if buttons:
                    buttons[0].click_input()

        # Wait for preferences window
        timings.wait_until(timeout=5, retry_interval=0.1, func=lambda: IsPreferencesOpen(app), value=True)
        
        mseqPrefWindow = None
        for title in ['Mseq Preferences', 'mSeq Preferences', 'MSEQ Preferences']:
            if app.window(title=title).exists():
                mseqPrefWindow = app.window(title=title)
                break
        
        if not mseqPrefWindow:
            mseqPrefWindow = app.window(title_re='.*[Pp]references.*')
        
        # Try to click OK button in preferences window
        if not click_button_by_text(mseqPrefWindow, "&OK", ["Button"]):
            if not click_button_by_text(mseqPrefWindow, "OK", ["Button"]):
                # Last resort - try to find any button that might be the OK button
                buttons = mseqPrefWindow.children(class_name="Button")
                if buttons:
                    buttons[-1].click_input()  # Often the OK button is the last button

        # Wait for copy files dialog
        timings.wait_until(timeout=5, retry_interval=0.1, func=lambda: IsCopyFilesOpen(app), value=True)
        
        copySeqFilesWindow = None
        for title in ['Copy sequence files', 'Copy Sequence Files', 'Select Files']:
            if app.window(title=title, class_name='#32770').exists():
                copySeqFilesWindow = app.window(title=title, class_name='#32770')
                break
        
        if not copySeqFilesWindow:
            copySeqFilesWindow = app.window(class_name='#32770')
        
        # Try multiple methods to select all files
        try:
            # Original method
            shellViewFiles = copySeqFilesWindow.child_window(title="ShellView", class_name="SHELLDLL_DefView")
            listView = shellViewFiles.child_window(class_name="DirectUIHWND")
            listView.click_input()
            send_keys('^a')
        except:
            # Alternative methods for Windows 11
            try:
                # Try any list view control
                for class_name in ["DirectUIHWND", "SysListView32"]:
                    try:
                        listView = copySeqFilesWindow.descendant(class_name=class_name)
                        listView.click_input()
                        send_keys('^a')
                        break
                    except:
                        continue
            except:
                # Last resort - click in the middle and Ctrl+A
                copySeqFilesWindow.click_input(coords=(100, 100))
                sleep(0.2)
                send_keys('^a')
        
        # Find and click Open button
        if not click_button_by_text(copySeqFilesWindow, "&Open", ["Button"]):
            if not click_button_by_text(copySeqFilesWindow, "Open", ["Button"]):
                # Try alternatives
                for button_text in ["OK", "Select", "Continue"]:
                    if click_button_by_text(copySeqFilesWindow, button_text, ["Button"]):
                        break

        # Wait for file error window (this always happens because of mseq.ini)
        timings.wait_until(timeout=20, retry_interval=0.5, func=lambda: IsErrorWindowOpen(app), value=True)
        
        fileErrorWindow = None
        for title in ['File error', 'File Error', 'Error']:
            if app.window(title=title).exists():
                fileErrorWindow = app.window(title=title)
                break
        
        if fileErrorWindow:
            # Click OK on error window
            if not click_button_by_text(fileErrorWindow, "OK", ["Button"]):
                # Try clicking any button
                buttons = fileErrorWindow.children(class_name="Button")
                if buttons:
                    buttons[0].click_input()
        
        # Wait for call bases dialog
        timings.wait_until(timeout=10, retry_interval=0.3, func=lambda: IsCallBasesOpen(app), value=True)
        
        callBasesWindow = None
        for title in ['Call bases?', 'Call Bases?']:
            if app.window(title=title).exists():
                callBasesWindow = app.window(title=title)
                break
        
        if callBasesWindow:
            # Click Yes on call bases window
            if not click_button_by_text(callBasesWindow, "&Yes", ["Button"]):
                if not click_button_by_text(callBasesWindow, "Yes", ["Button"]):
                    # Try clicking any button that might be Yes
                    buttons = callBasesWindow.children(class_name="Button")
                    if buttons:
                        buttons[0].click_input()
        
        # Wait for processing to complete
        timings.wait_until(timeout=45, retry_interval=0.2, func=lambda: FivetxtORlowQualityWindow(app, orderpath), value=True)
        
        # Handle low quality files window if it appears
        if app.window(title="Low quality files skipped").exists():
            lowQualityWindow = app.window(title='Low quality files skipped')
            if not click_button_by_text(lowQualityWindow, "OK", ["Button"]):
                # Try clicking any button
                buttons = lowQualityWindow.children(class_name="Button")
                if buttons:
                    buttons[0].click_input()
        
        # Check for read info window
        timings.wait_until(timeout=5, retry_interval=0.1, func=lambda: IsReadInfoOpen(app), value=True)
        
        # Close the read info window if it exists
        readWindow = None
        if app.window(title_re='Read information for*').exists():
            readWindow = app.window(title_re='Read information for*')
        elif app.window(title_re='Read Information For*').exists():
            readWindow = app.window(title_re='Read Information For*')
        
        if readWindow:
            readWindow.close()
        
        return True
        
    except Exception as e:
        print(f"Error processing {orderpath}: {str(e)}")
        # Check if we've been running too long
        if time.time() - start_time > max_operation_time:
            print(f"Operation timed out for {orderpath}")
        
        # Try to recover
        try:
            for window in app.windows():
                try:
                    if window.exists():
                        window.close()
                except:
                    pass
            app.kill()
        except:
            pass
        
        return False

def select_all_files_in_copy_dialog(copySeqFilesWindow):
    """Try multiple methods to select all files in the copy dialog."""
    # Method 1: Using the shell view control (original method)
    try:
        shellView = copySeqFilesWindow.child_window(title="ShellView", class_name="SHELLDLL_DefView")
        listView = shellView.child_window(class_name="DirectUIHWND")
        listView.click_input()
        send_keys('^a')
        return True
    except:
        pass
    
    # Method 2: Using Windows 11's updated explorer UI
    try:
        # Try to find any list view control and select all
        for class_name in ["DirectUIHWND", "SysListView32", "ListView"]:
            try:
                listView = copySeqFilesWindow.descendant(class_name=class_name)
                listView.click_input()
                send_keys('^a')
                return True
            except:
                continue
    except:
        pass
    
    # Method 3: Click in the center of the dialog then Ctrl+A
    try:
        copySeqFilesWindow.click_input(coords=(100, 100))  # Click somewhere in the middle
        sleep(0.1)
        send_keys('^a')
        return True
    except:
        return False
        
def NavigateToFolder(browsefolderswindow, path):
    print(f"Navigating to folder: {path}")
    browsefolderswindow.set_focus()
    
    # Find the tree view using multiple strategies
    try:
        treeView = browsefolderswindow.child_window(auto_id="100", class_name="SysTreeView32")
        print("Found tree view using auto_id")
    except:
        try:
            shellNameSpace = browsefolderswindow.child_window(class_name="SHBrowseForFolder ShellNameSpace Control")
            treeView = shellNameSpace.child_window(class_name="SysTreeView32")
            print("Found tree view through ShellNameSpace control")
        except:
            # Fallback to the original method
            treeView = browsefolderswindow.child_window(class_name="SysTreeView32")
            print("Found tree view using class_name only")
    
    # Build the path to start from - always starts from Desktop\This PC
    virtualFolder = get_virtual_folder_name()
    print(f"Virtual folder name: {virtualFolder}")
    desktopPath = f'\\Desktop\\{virtualFolder}'
    
    # Try to find the starting item
    try:
        item = treeView.get_item(desktopPath)
        print(f"Found starting item using path: {desktopPath}")
    except:
        try:
            item = treeView.get_item('\\Desktop')
            print("Found starting item using path: \\Desktop")
        except:
            # Last resort - try to find the first item that looks like a desktop or computer
            print("Trying alternative method to find starting point")
            for child in treeView.items():
                if "Desktop" in child.text() or virtualFolder in child.text():
                    item = child
                    print(f"Found starting item: {child.text()}")
                    break
            else:
                print("ERROR: Could not find Desktop or starting item")
                return False
    
    # Split our target path into individual folder names
    data = path.split('\\')
    print(f"Navigating through: {data}")
    
    # For each folder in our path...
    for folder in data:
        if not folder:  # Skip empty path segments
            continue
            
        # Handle network drives
        if 'P:' in folder:
            folder = sub('P:', r'ABISync (P:)', folder)
        if 'H:' in folder:
            folder = sub('H:', r'Tyler (\\\\w2k16\\users) (H:)', folder)
        
        print(f"Looking for folder: {folder}")
        
        # Look through all items at the current level
        found = False
        for child in item.children():
            if child.text() == folder:
                print(f"Found folder: {folder}")
                browsefolderswindow.set_focus()
                item = child
                item.click_input()
                found = True
                break
        
        # If not found, try clicking the parent to refresh and try again
        if not found:
            print(f"Folder not found on first try: {folder}")
            item.click_input()
            sleep(0.5)
            
            for child in item.children():
                if child.text() == folder:
                    print(f"Found folder after refresh: {folder}")
                    browsefolderswindow.set_focus()
                    item = child
                    item.click_input()
                    found = True
                    break
                    
            if not found:
                # Try case-insensitive matching
                for child in item.children():
                    if folder.lower() == child.text().lower():
                        print(f"Found folder with case-insensitive match: {child.text()}")
                        browsefolderswindow.set_focus()
                        item = child
                        item.click_input()
                        found = True
                        break
                
                if not found:
                    print(f"WARNING: Could not find folder: {folder}")
                    return False
    
    print("Navigation complete")
    return True

# Main program
if __name__ == "__main__":
    # Ask user to pick a folder
    dataFolderPath = filedialog.askdirectory(title="Select today's data folder. To mseq folders.")
    dataFolderPath = sub(r'/','\\\\', dataFolderPath)
    
    # Get all folders in the selected directory
    PlateFolders = GetPlateFolders(dataFolderPath)

    fsaFiles = False
    from pywinauto import Application
    from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
    from pywinauto.keyboard import send_keys
    from pywinauto import timings
    
    # For each folder...
    for folder in PlateFolders:
        print(folder)
        
        # Check if it's already been processed
        mseqed = CheckForMseq(folder)
        # Check if it has FSA files (which we don't process)
        fsaFiles = CheckForfsaFiles(folder)
        
        # If not already processed and no FSA files, process it
        if not mseqed and not fsaFiles:
            MseqOrder(folder)