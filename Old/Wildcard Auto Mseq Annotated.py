from tkinter import filedialog
from os import listdir, path as OsPath
from re import sub
from time import sleep

firstTimeBrowsingMseq = True

# This function gets the name that Windows uses for "This PC" in File Explorer
# It uses the Windows COM interface because this name can be different on different systems
def get_virtual_folder_name():
    import win32com.client
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

# This is the main function that tries to navigate through Windows' folder selection dialog
# browsefolderswindow = the folder selection window that opens when you click "browse"
# path = the full path we're trying to navigate to (e.g. "P:\Data\SomeFolder")

def NavigateToFolder(browsefolderswindow, path):
    # Make sure the browse window is active/focused
    browsefolderswindow.set_focus()
    
    # Find the tree view control (the left-side panel showing folders) in the browse window
    treeView = browsefolderswindow.child_window(class_name="SysTreeView32")
    
    # Build the path to start from - always starts from Desktop\This PC
    desktopPath = f'\\Desktop\\{get_virtual_folder_name()}'
    
    # Get the "This PC" item in the tree view
    item = treeView.get_item(desktopPath)
    
    # Split our target path into individual folder names
    data = path.split('\\')
    
    # For each folder in our path...
    for folder in data:
        # Handle network drives - replace drive letters with their full names
        if 'P:' in folder:
            folder = sub('P:', r'ABISync (P:)', folder)
        if 'H:' in folder:
            folder = sub('H:', r'Tom (\\\\w2k16\\users) (H:)', folder)
            
        # Look through all items at the current level
        for child in item.children():
            if child.text() == folder:
                # When we find our folder, focus the window and click it
                browsefolderswindow.set_focus()
                item = child
                item.click_input()
                break

# This function handles the entire mSeq process
def MseqOrder(orderpath):
    # Get or start mSeq application and get its main window
    app, mseqMainWindow = GetMseq()
    mseqMainWindow.set_focus()
    
    # Press Ctrl+N to start a new project
    send_keys('^n')

    # Wait up to 5 seconds for the browse dialog to appear
    # Checks every 0.1 seconds using IsBrowseDialogOpen function
    timings.wait_until(timeout=5, retry_interval=.1, 
                      func=lambda: IsBrowseDialogOpen(app), value=True)
    
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
    NavigateToFolder(dialogWindow, orderpath)
    
    # Find and click the OK button in the browse dialog
    okDialogButton = app.BrowseForFolder.child_window(title="OK", class_name="Button")
    okDialogButton.click_input()

    # Wait for the Preferences window to appear
    timings.wait_until(timeout=5, retry_interval=.1, 
                      func=lambda: IsPreferencesOpen(app), value=True)
    mseqPrefWindow = app.window(title='Mseq Preferences')
    # Click OK in preferences window
    okPrefButton = mseqPrefWindow.child_window(title="&OK", class_name="Button")
    okPrefButton.click_input()

    # Handle the file selection dialog
    timings.wait_until(timeout=5, retry_interval=.1, 
                      func=lambda: IsCopyFilesOpen(app), value=True)
    copySeqFilesWindow = app.window(title='Copy sequence files', class_name='#32770')
    
    # Get the shell view control (the area showing files) in the file dialog
    shellViewFiles = copySeqFilesWindow.child_window(title="ShellView", 
                                                    class_name="SHELLDLL_DefView")
    
    # Get the actual list view within the shell view
    listView = shellViewFiles.child_window(class_name="DirectUIHWND")
    
    # Click in the list view area and press Ctrl+A to select all files
    listView.click_input()
    send_keys('^a')
    
    # Click the Open button
    copySeqOpenButton = copySeqFilesWindow.child_window(title="&Open", 
                                                       class_name="Button")
    copySeqOpenButton.click_input()

    # Handle the error dialog that appears because we selected mseq.ini
    timings.wait_until(timeout=20, retry_interval=.5, 
                      func=lambda: IsErrorWindowOpen(app), value=True)
    fileErrorWindow = app.window(title='File error')
    fileErrorOkButton = fileErrorWindow.child_window(class_name="Button")
    fileErrorOkButton.click_input()

    # Handle the "Call bases?" dialog
    timings.wait_until(timeout=10, retry_interval=.3, 
                      func=lambda: IsCallBasesOpen(app), value=True)
    callBasesWindow = app.window(title='Call bases?')
    callBasesYesButton = callBasesWindow.child_window(title="&Yes", 
                                                     class_name="Button")
    callBasesYesButton.click_input()

    # Wait for processing to complete - either 5 txt files appear or we get low quality warning
    timings.wait_until(timeout=45, retry_interval=.2, 
                      func=lambda: FivetxtORlowQualityWindow(app, orderpath), 
                      value=True)
    
    # Handle the low quality warning if it appears
    if app.window(title="Low quality files skipped").exists():
        lowQualityWindow = app.window(title='Low quality files skipped')
        lowQualityOkButton = lowQualityWindow.child_window(class_name="Button")
        lowQualityOkButton.click_input()

    # Close the read information window if it appears
    timings.wait_until(timeout=5, retry_interval=.1, 
                      func=lambda: IsReadInfoOpen(app), value=True)
    if app.window(title_re='Read information for*').exists:
        readWindow = app.window(title_re='Read information for*')
        readWindow.close()

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