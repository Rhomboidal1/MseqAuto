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
    #from pywinauto.application import Application
    #from pywinauto.application import findwindows
    #from pywinauto.keyboard import send_keys
    #from pywinauto import timings
    app, mseqMainWindow = GetMseq()
    mseqMainWindow.set_focus()
    send_keys('^n')

    timings.wait_until(timeout=5, retry_interval=.1, func=lambda: IsBrowseDialogOpen(app), value=True)
    dialogWindow = app.window(title='Browse For Folder')
    global firstTimeBrowsingMseq
    if firstTimeBrowsingMseq:
        firstTimeBrowsingMseq = False
        sleep(1.2) #mseq and file system is sleepy, ok.. needs some time to wake.
    else:
        sleep(.5)
    
    NavigateToFolder(dialogWindow, orderpath)
    okDialogButton = app.BrowseForFolder.child_window(title="OK", class_name="Button")
    okDialogButton.click_input()

    timings.wait_until(timeout=5, retry_interval=.1, func=lambda: IsPreferencesOpen(app), value=True)
    mseqPrefWindow = app.window(title='Mseq Preferences')
    okPrefButton = mseqPrefWindow.child_window(title="&OK", class_name="Button")
    okPrefButton.click_input()

    timings.wait_until(timeout=5, retry_interval=.1, func=lambda: IsCopyFilesOpen(app), value=True)
    copySeqFilesWindow = app.window(title='Copy sequence files', class_name='#32770')
    shellViewFiles = copySeqFilesWindow.child_window(title="ShellView", class_name="SHELLDLL_DefView")      #get the shell view control from the file dialog window
    listView = shellViewFiles.child_window(class_name="DirectUIHWND")                       #get this thing, whatever it is (it's like a UI element displaying the list of files in the selected directory)
    listView.click_input()                                                                  #click on that thing, whatever it is (clicks right in the center of the window)
    send_keys('^a')     
    #fileComboBox = copySeqFilesWindow.child_window(class_name="ComboBoxEx32")
    #fileEditBox = fileComboBox.child_window(class_name='Edit')
    #fileEditBox.set_edit_text(fileString)
    copySeqOpenButton = copySeqFilesWindow.child_window(title="&Open", class_name="Button")
    copySeqOpenButton.click_input()

    timings.wait_until(timeout=20, retry_interval=.5, func=lambda: IsErrorWindowOpen(app), value=True)
    fileErrorWindow = app.window(title='File error')                                        #since we ^a in the shell view, we'll always get the mseq.ini file - not a seq file type so we get this file error window every time.
    fileErrorOkButton = fileErrorWindow.child_window(class_name="Button")                   #get the ok button in the error window
    fileErrorOkButton.click_input()  
    
    timings.wait_until(timeout=10, retry_interval=.3, func=lambda: IsCallBasesOpen(app), value=True)   
    callBasesWindow = app.window(title='Call bases?')
    callBasesYesButton = callBasesWindow.child_window(title="&Yes", class_name="Button")
    callBasesYesButton.click_input()
    #phredWindow = app.window(title='Phred')                                                 #three cmd windows
    #phd2fastaWindow = app.window(title='phd2fasta')
    #crossmatch = app.window(title="Crossmatch")
    timings.wait_until(timeout=45, retry_interval=.2, func=lambda: FivetxtORlowQualityWindow(app, orderpath), value=True)   #wait until all bases are called. one of two things happens - either a 5th txt file shows up in the order folder or the low quality read window opens.
    if app.window(title="Low quality files skipped").exists():
        lowQualityWindow = app.window(title='Low quality files skipped')
        lowQualityOkButton = lowQualityWindow.child_window(class_name="Button")
        lowQualityOkButton.click_input()

    timings.wait_until(timeout=5, retry_interval=.1, func=lambda: IsReadInfoOpen(app), value=True)
    if app.window(title_re='Read information for*').exists:
        readWindow = app.window(title_re='Read information for*')
        readWindow.close()
def NavigateToFolder(browsefolderswindow, path):
    # Make sure the browse window is active/focused
    browsefolderswindow.set_focus()

    # Find the tree view control (the left-side panel showing folders) in the browse window
    # Using AutomationId is the most reliable way if available
    try:
        treeView = browsefolderswindow.child_window(auto_id="100", class_name="SysTreeView32")
    except pywinauto.findwindows.ElementNotFoundError:
        # Fallback to the original method if AutomationId is not found
        treeView = browsefolderswindow.child_window(class_name="SysTreeView32")

    # Build the path to start from - always starts from Desktop\This PC
    desktopPath = f'\\Desktop\\{get_virtual_folder_name()}'

    # Get the "Desktop" item in the tree view
    item = treeView.get_item(desktopPath)

    # Split our target path into individual folder names
    data = path.split('\\')

    # For each folder in our path...
    for folder in data:
        # Handle network drives - replace drive letters with their full names
        if 'P:' in folder:
            folder = sub('P:', r'ABISync (P:)', folder)
        if 'H:' in folder:
            folder = sub('H:', r'Tyler (\\\\w2k16\\users) (H:)', folder)  #Use name from inspect.exe

        # Look through all items at the current level
        for child in item.children():
            # Compare the child's name with the folder we are looking for
            if child.text() == folder:
                # When we find our folder, focus the window and click it
                browsefolderswindow.set_focus()
                item = child
                item.click_input()
                break
'''
#function clicks through the treeview control to get to the folder the user selected.
def NavigateToFolder(browsefolderswindow, path):
    browsefolderswindow.set_focus()
    treeView = browsefolderswindow.child_window(class_name="SysTreeView32")
    
    desktopPath = f'\\Desktop\\{get_virtual_folder_name()}'
    #print(desktopPath)
    item = treeView.get_item(desktopPath)
    data = path.split('\\')
    for folder in data:
        if 'P:' in folder:
            #folder = sub('P:', r'abisync (\\\\w2k16) (P:)', folder)         OLD DRIVE NAME
            folder = sub('P:', r'ABISync (P:)', folder)                     #NEW DRIVE NAME
        if 'H:' in folder:
            folder = sub('H:', r'Tom (\\\\w2k16\\users) (H:)', folder)
        for child in item.children():
            #print(child.text())
            if child.text() == folder:
                browsefolderswindow.set_focus()
                item = child
                item.click_input()              #the last folder will be clicked on, so the next mseq function should be ready to just click ok.
                break
'''   
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