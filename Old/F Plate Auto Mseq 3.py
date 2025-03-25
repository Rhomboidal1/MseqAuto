import warnings
import sys
sys.coinit_flags = 2
#warnings.filterwarnings("ignore", category=UserWarning, module="pywinauto")
warnings.filterwarnings("ignore", message="Apply externally defined coinit_flags*", category=UserWarning)
warnings.filterwarnings("ignore", message="Revert to STA COM threading mode", category=UserWarning)
import os
import subprocess

# Check if running in 64-bit Python
if sys.maxsize > 2**32:
    # Path to 32-bit Python
    py32_path = r"C:\Python312-32\python.exe"
    
    if os.path.exists(py32_path):
        # Get the full path of the current script
        script_path = os.path.abspath(__file__)
        
        # Re-run this script with 32-bit Python and exit current process
        subprocess.run([py32_path, script_path])
        sys.exit(0)
    else:
        print("32-bit Python not found at", py32_path)
        print("Continuing with 64-bit Python (may cause issues)")
        
from tkinter import filedialog
from os import listdir, path as OsPath
from re import search, sub
from time import sleep as TimeSleep

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
                if search('p\\d+.+', item.lower()) and item.lower()[:1] == 'p':
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
    #from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
    #from pywinauto import Application
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

def MseqOrder(orderpath):
    try:
        print(f"Starting mSeq for {orderpath}")
        app, mseqMainWindow = GetMseq()
        mseqMainWindow.set_focus()
        send_keys('^n')

        # Wait for Browse dialog with slightly increased timeout
        timings.wait_until(timeout=8, retry_interval=.2, func=lambda: IsBrowseDialogOpen(app), value=True)
        dialogWindow = app.window(title='Browse For Folder')
        
        # Keep the original sleep timing logic that was working before
        global firstTimeBrowsingMseq
        if firstTimeBrowsingMseq:
            firstTimeBrowsingMseq = False
            TimeSleep(1.2)  # mseq and file system is sleepy, ok.. needs some time to wake.
        else:
            TimeSleep(.5)
        
        # Use the original NavigateToFolder without modifying its internal logic
        NavigateToFolder(dialogWindow, orderpath)
        
        # Try to click OK button, with a fallback
        try:
            okDiaglogButton = app.BrowseForFolder.child_window(title="OK", class_name="Button")
            okDiaglogButton.click_input()
        except Exception as e:
            print(f"Error clicking OK button: {str(e)}")
            try:
                # Try the Enter key as a backup
                send_keys('{ENTER}')
            except:
                pass

        # Increase preference window timeout slightly
        timings.wait_until(timeout=8, retry_interval=.2, func=lambda: IsPreferencesOpen(app), value=True)
        mseqPrefWindow = app.window(title='Mseq Preferences')
        okPrefButton = mseqPrefWindow.child_window(title="&OK", class_name="Button")
        okPrefButton.click_input()
        
        # Continue with mostly original code but with slightly increased timeouts
        timings.wait_until(timeout=8, retry_interval=.2, func=lambda: IsCopyFilesOpen(app), value=True)
        copySeqFilesWindow = app.window(title='Copy sequence files', class_name='#32770')
        
        # Try the original file selection method first
        try:
            shellViewFiles = copySeqFilesWindow.child_window(title="ShellView", class_name="SHELLDLL_DefView")
            listView = shellViewFiles.child_window(class_name="DirectUIHWND")
            listView.click_input()
            send_keys('^a')
        except Exception as e:
            print(f"Error using original file selection method: {str(e)}")
            # Try a simpler approach as fallback
            try:
                # Just press Ctrl+A and hope for the best
                copySeqFilesWindow.set_focus()
                TimeSleep(0.5)
                send_keys('^a')
            except:
                pass
                
        # Click Open button
        copySeqOpenButton = copySeqFilesWindow.child_window(title="&Open", class_name="Button")
        copySeqOpenButton.click_input()

        # Keep original timeouts for these steps
        timings.wait_until(timeout=20, retry_interval=.5, func=lambda: IsErrorWindowOpen(app), value=True)
        fileErrorWindow = app.window(title='File error')
        fileErrorOkButton = fileErrorWindow.child_window(class_name="Button")
        fileErrorOkButton.click_input()
        
        timings.wait_until(timeout=10, retry_interval=.3, func=lambda: IsCallBasesOpen(app), value=True)
        callBasesWindow = app.window(title='Call bases?')
        callBasesYesButton = callBasesWindow.child_window(title="&Yes", class_name="Button")
        callBasesYesButton.click_input()

        # Keep the original long timeout for processing
        timings.wait_until(timeout=45, retry_interval=.2, func=lambda: FivetxtORlowQualityWindow(app, orderpath), value=True)
        
        if app.window(title="Low quality files skipped").exists():
            lowQualityWindow = app.window(title='Low quality files skipped')
            lowQualityOkButton = lowQualityWindow.child_window(class_name="Button")
            lowQualityOkButton.click_input()
        
        timings.wait_until(timeout=5, retry_interval=.1, func=lambda: IsReadInfoOpen(app), value=True)
        if app.window(title_re='Read information for*').exists():
            readWindow = app.window(title_re='Read information for*')
            readWindow.close()
            
        print(f"Successfully completed mSeq for {orderpath}")
        return True
        
    except Exception as e:
        print(f"Error in MseqOrder function: {str(e)}")
        # If there's an error, try to recover by pressing Escape to close any open dialogs
        try:
            send_keys('{ESC}')
            TimeSleep(0.5)
        except:
            pass
        return False

#function clicks through the treeview control to get to the folder the user selected.
def get_windows_version():
    """Detect whether the system is running Windows 10 or Windows 11"""
    import platform
    import sys
    
    # Get Windows version info
    windows_version = sys.getwindowsversion()
    build_number = windows_version.build
    
    # Windows 11 is Windows 10 version 21H2 build 22000 or higher
    if build_number >= 22000:
        return "Windows 11"
    else:
        return "Windows 10"

# Define both navigation functions
def NavigateToFolder_Win10(browsefolderswindow, path):
    """Original navigation function that works on Windows 10"""
    browsefolderswindow.set_focus()
    treeView = browsefolderswindow.child_window(class_name="SysTreeView32")
    
    desktopPath = f'\\Desktop\\{get_virtual_folder_name()}'
    item = treeView.get_item(desktopPath)
    data = path.split('\\')
    for folder in data:
        if 'P:' in folder:
            folder = sub('P:', r'ABISync (P:)', folder)
        if 'H:' in folder:
            folder = sub('H:', r'Tom (\\\\w2k16\\users) (H:)', folder)
        for child in item.children():
            if child.text() == folder:
                browsefolderswindow.set_focus()
                item = child
                item.click_input()
                break

def NavigateToFolder_Win11(browsefolderswindow, path):
    """Enhanced navigation function for Windows 11"""
    browsefolderswindow.set_focus()
    treeView = browsefolderswindow.child_window(class_name="SysTreeView32")
    
    # Try different possible desktop paths
    possible_paths = [
        f'\\Desktop\\{get_virtual_folder_name()}',
        '\\Desktop\\This PC',
        '\\This PC'
    ]
    
    item = None
    for desktop_path in possible_paths:
        try:
            item = treeView.get_item(desktop_path)
            print(f"Successfully found navigation starting point: {desktop_path}")
            break
        except Exception as e:
            print(f"Failed to find {desktop_path}: {str(e)}")
    
    if not item:
        print("Could not find any valid starting point for navigation")
        # As a last resort, try to directly type the path
        try:
            send_keys(path)
            return
        except:
            return
    
    # Navigate through the folder structure
    data = path.split('\\')
    for folder in data:
        # Handle network drives
        if 'P:' in folder:
            folder = sub('P:', r'ABISync (P:)', folder)
        if 'H:' in folder:
            folder = sub('H:', r'Tom (\\\\w2k16\\users) (H:)', folder)
        
        found_child = False
        for child in item.children():
            if child.text() == folder:
                browsefolderswindow.set_focus()
                item = child
                item.click_input()
                found_child = True
                print(f"Found and selected folder: {folder}")
                break
        
        if not found_child:
            print(f"Could not find folder {folder} in the tree view")
            # Try to search for partial matches
            for child in item.children():
                if folder.lower() in child.text().lower():
                    browsefolderswindow.set_focus()
                    item = child
                    item.click_input()
                    found_child = True
                    print(f"Found partial match: {child.text()} for {folder}")
                    break

# Choose the correct function based on Windows version
def NavigateToFolder(browsefolderswindow, path):
    """Wrapper function that calls the appropriate NavigateToFolder function based on OS version"""
    windows_version = get_windows_version()
    print(f"Detected {windows_version}")
    
    if windows_version == "Windows 11":
        NavigateToFolder_Win11(browsefolderswindow, path)
    else:
        NavigateToFolder_Win10(browsefolderswindow, path)
#End functions 
########################################################################################################


########################################################################################################
#Main code
if __name__ == "__main__":
    #dataFolderPath = filedialog.askdirectory(parent=dialog, title="Select today's data folder. To mseq orders.")
    dataFolderPath = filedialog.askdirectory(title="Select today's data folder. To mseq plates.")
    #dialog.destroy()
    dataFolderPath = sub(r'/','\\\\', dataFolderPath)
    PlateFolders  = GetPlateFolders(dataFolderPath)  #get bioIfolders that have been sorted already
    #print(bioIFolders)
    fsaFiles = False
    from pywinauto import Application
    from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
    from pywinauto.keyboard import send_keys
    from pywinauto import timings
    
    
    #loop through each bioI folder
    for folder in PlateFolders:
        print(folder)
        
        #mseq ab1 files.
        mseqed = CheckForMseq(folder)
        fsaFiles = CheckForfsaFiles(folder)    
        if not mseqed and not fsaFiles:
            MseqOrder(folder)