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
from time import sleep


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
    okDiaglogButton = app.BrowseForFolder.child_window(title="OK", class_name="Button")
    okDiaglogButton.click_input()

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