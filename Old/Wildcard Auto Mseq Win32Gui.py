import warnings
import sys
sys.coinit_flags = 2  # Move this before any COM imports

warnings.filterwarnings("ignore", message="Apply externally defined coinit_flags*", category=UserWarning)
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", category=UserWarning)

from tkinter import filedialog
from os import listdir, path as OsPath
from re import sub
import commctrl
import win32api
import win32gui
import win32con
import win32com.client
from time import sleep
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
from pywinauto.keyboard import send_keys
from pywinauto import timings

firstTimeBrowsingMseq = True

# This function gets the name that Windows uses for "This PC" in File Explorer
# It uses the Windows COM interface because this name can be different on different systems
def get_virtual_folder_name():
    shell = win32com.client.Dispatch("Shell.Application")
    namespace = shell.Namespace(0x11)  # 0x11 is the code for "My Computer"/"This PC"
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

def find_dialog_controls(hwnd):
    """Find key controls in the browse dialog window"""
    results = {}
    
    def callback(child_hwnd, results):
        class_name = win32gui.GetClassName(child_hwnd)
        try:
            text = win32gui.GetWindowText(child_hwnd)
            # Debug info
            print(f"Found control - Class: {class_name}, Text: {text}")
        except:
            text = ""
            
        # Store handles for controls we care about
        if class_name == "SysTreeView32":
            results['tree'] = child_hwnd
        elif text == "OK" and class_name == "Button":
            results['ok'] = child_hwnd
        
        return True
    
    # Enumerate child windows to find controls
    win32gui.EnumChildWindows(hwnd, callback, results)
    return results

def set_focus_and_click(hwnd):
    """Set focus to window and simulate click"""
    if not hwnd:
        return False
        
    try:
        # Bring window to front
        win32gui.SetForegroundWindow(hwnd)
        sleep(0.1)
        
        # Get window rect for clicking center
        rect = win32gui.GetWindowRect(hwnd)
        center_x = (rect[0] + rect[2]) // 2
        center_y = (rect[1] + rect[3]) // 2
        
        # Send click messages
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, 0, center_y << 16 | center_x)
        sleep(0.1)
        win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, center_y << 16 | center_x)
        return True
    except:
        return False

def navigate_shell_folder(shell, target_path):
    """Navigate to folder using Shell.Application"""
    print(f"Starting shell navigation to: {target_path}")
    
    # First try with Desktop (0x10)
    try:
        # Split path into parts
        parts = [p for p in target_path.split('\\') if p]
        print(f"Path parts: {parts}")
        
        # Start from desktop
        try:
            folder = shell.NameSpace(0x10)  # CSIDL_DESKTOP
            print("Got Desktop folder using 0x10")
        except:
            folder = shell.NameSpace(0)  # Try alternate Desktop method
            print("Got Desktop folder using 0")
            
        if not folder:
            print("Could not get Desktop folder")
            return False
            
        # Navigate each part
        for part in parts:
            print(f"Navigating to part: {part}")
            # Handle network drives
            if part == 'P:':
                part = 'ABISync (P:)'
            elif part == 'H:':
                part = 'Tyler (\\\\w2k16\\users) (H:)'
            
            print(f"Looking for: {part}")
            # Look for folder in current location
            found = False
            try:
                items = list(folder.Items())
                for item in items:
                    if item.Name == part:
                        print(f"Found {part}")
                        folder = item.GetFolder
                        found = True
                        break
                
                if not found:
                    print(f"Could not find {part}")
                    return False
            except Exception as e:
                print(f"Error while looking for {part}: {str(e)}")
                return False
                
        print("Successfully navigated to target")
        return True
        
    except Exception as e:
        print(f"Shell navigation error: {str(e)}")
        return False
def NavigateToFolder(browsefolderswindow, path):
    from pywinauto.keyboard import send_keys
    from time import sleep
    import os

    print(f"Attempting to navigate to: {path}")
    
    # Try multiple navigation strategies
    navigation_strategies = [
        # Strategy 1: Direct path typing
        lambda: _navigate_by_path_typing(browsefolderswindow, path),
        
        # Strategy 2: Tree view navigation with extended error handling
        lambda: _navigate_by_treeview(browsefolderswindow, path),
        
        # Strategy 3: Keyboard-based navigation
        lambda: _navigate_by_keyboard(browsefolderswindow, path)
    ]

    for strategy in navigation_strategies:
        try:
            if strategy():
                print("Navigation successful")
                return True
        except Exception as e:
            print(f"Navigation strategy failed: {str(e)}")
    
    print("All navigation strategies failed")
    return False

def _navigate_by_path_typing(browsefolderswindow, path):
    """Navigate by directly typing the full path"""
    from pywinauto.keyboard import send_keys
    from time import sleep

    try:
        # Try to find and focus the address/path edit control
        edit_controls = browsefolderswindow.children(class_name='Edit')
        
        for edit in edit_controls:
            try:
                edit.set_focus()
                edit.click_input()
                sleep(0.3)
                send_keys('^a')  # Select all existing text
                send_keys('{DELETE}')  # Clear existing text
                send_keys(path)  # Type full path
                sleep(0.5)
                send_keys('{ENTER}')
                sleep(1.0)
                return True
            except Exception as e:
                print(f"Path typing on edit control failed: {str(e)}")
    except Exception as e:
        print(f"Finding edit controls failed: {str(e)}")
    
    return False

def _navigate_by_treeview(browsefolderswindow, path):
    """Navigate through treeview with more robust error handling"""
    from time import sleep
    
    try:
        # Get the SysTreeView32 control
        treeView = browsefolderswindow.child_window(class_name="SysTreeView32")
        
        # Break down the path components
        path_components = path.split('\\')
        
        # Start from desktop or This PC
        desktop_variants = [
            f'\\Desktop\\{get_virtual_folder_name()}', 
            '\\Desktop\\This PC', 
            '\\This PC'
        ]
        
        current_item = None
        for desktop_path in desktop_variants:
            try:
                current_item = treeView.get_item(desktop_path)
                break
            except:
                continue
        
        if not current_item:
            print("Could not find desktop/This PC in treeview")
            return False
        
        # Navigate through path components
        for component in path_components:
            # Handle drive name translations
            if 'P:' in component:
                component = component.replace('P:', 'ABISync (P:)')
            if 'H:' in component:
                component = component.replace('H:', 'Tyler (\\\\w2k16\\users) (H:)')
            
            # Find and click the matching child
            child_found = False
            for child in current_item.children():
                if child.text() == component:
                    browsefolderswindow.set_focus()
                    child.click_input()
                    current_item = child
                    child_found = True
                    sleep(0.5)  # Give time for UI to update
                    break
            
            if not child_found:
                print(f"Could not find folder: {component}")
                return False
        
        return True
    
    except Exception as e:
        print(f"Treeview navigation failed: {str(e)}")
        return False

def _navigate_by_keyboard(browsefolderswindow, path):
    """Navigate using keyboard shortcuts as a fallback"""
    from pywinauto.keyboard import send_keys
    from time import sleep

    try:
        # Try various keyboard shortcuts to access address bar or path
        keyboard_shortcuts = [
            ('%d', "Alt+D"),    # Address bar in many Windows dialogs
            ('^l', "Ctrl+L"),   # Another common address bar shortcut
            ('{F4}', "F4"),     # Dropdown/combo box navigation
        ]

        for shortcut, description in keyboard_shortcuts:
            try:
                browsefolderswindow.set_focus()
                sleep(0.5)
                send_keys(shortcut)
                sleep(0.5)
                send_keys('^a')  # Select all
                send_keys(path)  # Type path
                send_keys('{ENTER}')
                sleep(1.0)
                return True
            except Exception as e:
                print(f"{description} navigation failed: {str(e)}")
    
    except Exception as e:
        print(f"Keyboard navigation failed: {str(e)}")
    
    return False
'''
def NavigateToFolder(browsefolderswindow, path):
    """Test version to diagnose tree view focus"""
    try:
        print(f"Attempting to navigate to: {path}")
        hwnd = browsefolderswindow.handle
        print(f"Browse window handle: {hwnd}")
        
        # First ensure browse dialog has focus
        win32gui.SetForegroundWindow(hwnd)
        sleep(0.5)
        print("Set browse dialog as foreground")
        
        # Find tree view control
        controls = find_dialog_controls(hwnd)
        if 'tree' not in controls:
            print("Could not find tree view control")
            return False
            
        tree_hwnd = controls['tree']
        print(f"Found tree view handle: {tree_hwnd}")
        
        # Method 1: Try to click tree view directly
        try:
            # Get tree view rectangle
            tree_rect = win32gui.GetWindowRect(tree_hwnd)
            center_x = (tree_rect[0] + tree_rect[2]) // 2
            center_y = (tree_rect[1] + tree_rect[3]) // 2
            
            # Click in center of tree view
            win32api.SetCursorPos((center_x, center_y))
            sleep(0.1)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
            sleep(0.1)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
            sleep(0.5)
            print("Clicked in tree view center")
            
            # Try to get selected item info
            # We'll need to add the correct message constants here
            
        except Exception as e:
            print(f"Direct click failed: {str(e)}")
            
        # Method 2: Try tab method
        try:
            # Send two tabs
            win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
            sleep(0.1)
            win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
            sleep(0.2)
            win32api.keybd_event(win32con.VK_TAB, 0, 0, 0)
            sleep(0.1)
            win32api.keybd_event(win32con.VK_TAB, 0, win32con.KEYEVENTF_KEYUP, 0)
            sleep(0.5)
            print("Sent two tabs")
            
        except Exception as e:
            print(f"Tab method failed: {str(e)}")
            
        # For testing, return False so we can see the output
        return False
        
    except Exception as e:
        print(f"Navigation error: {str(e)}")
        return False

def NavigateToFolder(browsefolderswindow, path):
    """Navigate using the actual tree view control"""
    try:
        print(f"Attempting to navigate to: {path}")
        hwnd = browsefolderswindow.handle
        
        # Find controls and print all found
        controls = find_dialog_controls(hwnd)
        print(f"Found controls: {list(controls.keys())}")
        
        if 'tree' not in controls:
            print("Could not find tree view control")
            return False
            
        tree_hwnd = controls['tree']
        
        # Get tree view items
        try:
            tree_items = win32gui.SendMessage(tree_hwnd, win32con.TVM_GETNEXTITEM, win32con.TVGN_ROOT, 0)
            while tree_items:
                item_text = win32gui.SendMessage(tree_hwnd, win32con.TVM_GETITEM, 0, tree_items)
                print(f"Tree item: {item_text}")
                tree_items = win32gui.SendMessage(tree_hwnd, win32con.TVM_GETNEXTITEM, win32con.TVGN_NEXT, tree_items)
        except Exception as e:
            print(f"Error reading tree: {str(e)}")
        
        return False  # For now, just print structure
            
    except Exception as e:
        print(f"Navigation error: {str(e)}")
        return False
'''
#function returns both the app object and its main window.
def GetMseq():
    try:
        app = Application(backend='win32').connect(title_re='Mseq*', timeout=1)
        #return app
    except (ElementNotFoundError, timings.TimeoutError):
        try:
            app = Application(backend='win32').connect(title_re='mSeq*', timeout=1)
            #return app
        except (ElementNotFoundError, timings.TimeoutError):
            app = Application(backend='win32').start('cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', wait_for_idle=False).connect(title='mSeq', timeout=10)
            #return app
        except ElementAmbiguousError:
            app = Application(backend='win32').connect(title_re='mSeq*', found_index=0, timeout=1)
            app2 = Application(backend='win32').connect(title_re='mSeq*', found_index=1, timeout=1)
            app2.kill()
            #return app
    except ElementAmbiguousError:
        app = Application(backend='win32').connect(title_re='Mseq*', found_index=0, timeout=1)
        app2 = Application(backend='win32').connect(title_re='Mseq*', found_index=1, timeout=1)
        app2.kill()
        #return app
    
    if app.window(title_re='mSeq*').exists()==False:
        mainWindow = app.window(title_re='Mseq*')
    else:
        mainWindow = app.window(title_re='mSeq*')

    return app, mainWindow
       
def IsBrowseDialogOpen(myapp):
    #app = GetMseq()
    return myapp.window(title='Browse For Folder').exists()

def IsPreferencesOpen(myapp):
    return myapp.window(title='Mseq Preferences').exists()

def IsCopyFilesOpen(myapp):
    return myapp.window(title='Copy sequence files').exists()

def IsErrorWindowOpen(myapp):
    return myapp.window(title='File error').exists()

def IsCallBasesOpen(myapp):
    return myapp.window(title='Call bases?').exists()

def IsReadInfoOpen(myapp):
    return myapp.window(title_re='Read information for*').exists()

def FivetxtORlowQualityWindow(myApp, orderpath):
    #from pywinauto.application import Application
    #app = GetMseq()
    if myApp.window(title="Low quality files skipped").exists():
        return True
    if CheckFor5txt(orderpath):
        return True
    return False
# This function handles the entire mSeq process
def MseqOrder(orderpath):
    try:
        print(f"Starting mSeq process for {orderpath}")
        
        # Get or start mSeq application and get its main window
        app, mseqMainWindow = GetMseq()
        mseqMainWindow.set_focus()
        print("Got mSeq window")
        
        # Press Ctrl+N to start a new project
        send_keys('^n')
        print("Sent Ctrl+N")

        # Wait for browse dialog
        if not timings.wait_until(timeout=5, retry_interval=.1, 
                          func=lambda: IsBrowseDialogOpen(app), value=True):
            print("Browse dialog failed to appear")
            return False
        print("Browse dialog appeared")
        
        # Get a reference to the browse dialog window
        dialogWindow = app.window(title='Browse For Folder')
        
        # First time needs extra delay for system to "wake up"
        global firstTimeBrowsingMseq
        if firstTimeBrowsingMseq:
            firstTimeBrowsingMseq = False
            sleep(1.2)
        else:
            sleep(.5)
        
        # Try to navigate to our target folder
        if not NavigateToFolder(dialogWindow, orderpath):
            print("Navigation failed")
            return False
        print("Navigation succeeded")
        
        # Wait for the Preferences window
        if not timings.wait_until(timeout=10, retry_interval=.5, 
                          func=lambda: IsPreferencesOpen(app), value=True):
            print("Preferences window failed to appear")
            return False
        print("Preferences window appeared")

        mseqPrefWindow = app.window(title='Mseq Preferences')
        okPrefButton = mseqPrefWindow.child_window(title="&OK", class_name="Button")
        okPrefButton.click_input()
        print("Clicked OK on preferences")

        # Handle the file selection dialog
        if not timings.wait_until(timeout=5, retry_interval=.1, 
                          func=lambda: IsCopyFilesOpen(app), value=True):
            print("Copy files dialog failed to appear")
            return False
        print("Copy files dialog appeared")

        copySeqFilesWindow = app.window(title='Copy sequence files', class_name='#32770')
        shellViewFiles = copySeqFilesWindow.child_window(title="ShellView", 
                                                    class_name="SHELLDLL_DefView")
        listView = shellViewFiles.child_window(class_name="DirectUIHWND")
        listView.click_input()
        send_keys('^a')
        print("Selected all files")
        
        copySeqOpenButton = copySeqFilesWindow.child_window(title="&Open", 
                                                       class_name="Button")
        copySeqOpenButton.click_input()
        print("Clicked Open")

        # Handle the error dialog
        if not timings.wait_until(timeout=20, retry_interval=.5, 
                          func=lambda: IsErrorWindowOpen(app), value=True):
            print("File error dialog failed to appear")
            return False
        print("File error dialog appeared")

        fileErrorWindow = app.window(title='File error')
        fileErrorOkButton = fileErrorWindow.child_window(class_name="Button")
        fileErrorOkButton.click_input()
        print("Clicked OK on file error")

        # Handle Call bases dialog
        if not timings.wait_until(timeout=10, retry_interval=.3, 
                          func=lambda: IsCallBasesOpen(app), value=True):
            print("Call bases dialog failed to appear")
            return False
        print("Call bases dialog appeared")

        callBasesWindow = app.window(title='Call bases?')
        callBasesYesButton = callBasesWindow.child_window(title="&Yes", 
                                                     class_name="Button")
        callBasesYesButton.click_input()
        print("Clicked Yes on call bases")

        # Wait for processing
        if not timings.wait_until(timeout=45, retry_interval=.2, 
                          func=lambda: FivetxtORlowQualityWindow(app, orderpath), 
                          value=True):
            print("Processing timed out")
            return False
        print("Processing completed")
        
        # Handle low quality warning if it appears
        if app.window(title="Low quality files skipped").exists():
            lowQualityWindow = app.window(title='Low quality files skipped')
            lowQualityOkButton = lowQualityWindow.child_window(class_name="Button")
            lowQualityOkButton.click_input()
            print("Handled low quality warning")

        # Handle read info window
        if not timings.wait_until(timeout=5, retry_interval=.1, 
                          func=lambda: IsReadInfoOpen(app), value=True):
            print("Read info check timed out")
        elif app.window(title_re='Read information for*').exists():
            readWindow = app.window(title_re='Read information for*')
            readWindow.close()
            print("Closed read info window")

        return True

    except Exception as e:
        print(f"Error in MseqOrder: {str(e)}")
        # Try to recover by sending Escape key
        try:
            send_keys('{ESC}')
            sleep(0.5)
        except:
            pass
        return False

# Main program
if __name__ == "__main__":
    # Ask user to pick a folder
    dataFolderPath = filedialog.askdirectory(title="Select today's data folder. To mseq folders.")
    dataFolderPath = sub(r'/','\\\\', dataFolderPath)
    
    # Get all folders in the selected directory
    PlateFolders = GetPlateFolders(dataFolderPath)

    fsaFiles = False
    
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