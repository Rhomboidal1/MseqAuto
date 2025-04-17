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
        
from subprocess import CalledProcessError, run as SubprocessRun
from tkinter import filedialog
from os import listdir, path as OsPath, mkdir
from shutil import move as ShutilMove
from re import search, sub, findall
from time import time, sleep as TimeSleep
from datetime import date
from numpy import loadtxt, where

#function that gets today's date in format m-d-yyyy


def get_formatted_date():
    today = date.today()
    return f"{int(today.strftime('%m'))}-{int(today.strftime('%d'))}-{today.strftime('%Y')}"

#Function returns a version of the string without anything contained in {} ex: {01A}Sample_Primer{5_25}{I-12345} returns Sample_Primer
def CleanBracesFormat(aList):
    i = 0
    newList = []
    while i < len(aList):
        newList.append(sub('{.*?}', '', NeutralizeSuffixes(aList[i])))
        #newList.append(sub('{.*?}', '', aList[i]))
        i+=1
    return newList


#End Functions used for formatting
##############################################################################################################


##############################################################################################################
#Function gets the number from the folder name: '01-01-1999_BioI-12345_1' returns '12345'
def GetINumberFromFolderName(folderName):
    '''need to slice as the char befor bioi and '''
    #BioString = folderName[folderName.lower().find('bio'):]
    
    if search('bioi-\\d+', str(folderName).lower()):
        BioString = search('bioi-\\d+', str(folderName).lower()).group(0)
        Inumber = search('\\d+', BioString).group(0)
        return Inumber
    else:
        return None

#function finds and returns just the order number of the BioI folder name
def GetOrderNumberFromFolderName(folder):
    orderNum = ''
    if search('.+_\\d+', folder):
        orderNum = search('_\\d+', folder).group(0)
        orderNum = search('\\d+', orderNum).group(0)
        return orderNum
    else:
        return ''
    
    
def GetNumbers(files):
    numbers = []
    for file in files:
        numbers.append(GetINumberFromFolderName(file))
    return numbers


'''
def GetSubKey(INumbers):
    subkey = []
    i = 0
    while i < len(key):
        if key[i][0] in INumbers: 
            subkey.append(key[i].tolist())
        i+=1
    return subkey
'''
#End functions for today's I numbers and subkey
##############################################################################################################

#Begin functions for reinject things
##############################################################################################################
#Function changes characters in file name just as they are changed when making txt files from excel spreadsheets.
def AdjustABIChars(fileName):
    newFileName = fileName
    newFileName = newFileName.replace(' ', '')
    newFileName = newFileName.replace('+', '&')
    newFileName = newFileName.replace("*", "-")
    newFileName = newFileName.replace("|", "-")
    newFileName = newFileName.replace("/", "-")
    newFileName = newFileName.replace("\\", "-")
    newFileName = newFileName.replace(":", "-")
    newFileName = newFileName.replace("\"", "")
    newFileName = newFileName.replace("\'", "")
    newFileName = newFileName.replace("<", "-")
    newFileName = newFileName.replace(">", "-")
    newFileName = newFileName.replace("?", "")
    newFileName = newFileName.replace(",", "")
    return newFileName

def NeutralizeSuffixes(fileName):
    newFileName = fileName
    newFileName = newFileName.replace('_Premixed', '')
    newFileName = newFileName.replace('_RTI', '')
    return newFileName



def AdjustFullKeyToABIChars():
    i = 0
    while i < len(key):
        key[i][3] = AdjustABIChars(key[i][3])
        i+=1

##############################################################################################################

#Begin functions for determining order completion
##############################################################################################################
def GetOrderList(number):
    orderSampleNames = []
    orderIndexes = where((key == str(number)))
    i = 0 
    while i < len(orderIndexes[0]):
        adjustedName = AdjustABIChars(key[orderIndexes[0][i], 3])
        orderSampleNames.append(adjustedName)
        i+=1
    return orderSampleNames
        
def GetAllab1Files(folder):
    ab1FilePaths = []
    for item in listdir(folder):
        if OsPath.isfile(f'{folder}\\{item}'):
            #print(item)
            if item.endswith('.ab1'):
                ab1FilePaths.append(f'{folder}\\{item}')
    return ab1FilePaths

def OrderInReinjects(order):
    inReinjectList = False
    for rxn in order:
        if rxn in completeReinjectList:
            inReinjectList = True
            break
    return inReinjectList

#End functions for determining order completion
##############################################################################################################

#function gets the virtual folder name, often named This PC, but could be different per machine.
def get_virtual_folder_name():
    import win32com.client
    shell = win32com.client.Dispatch("Shell.Application")
    namespace = shell.Namespace(0x11)  # CSIDL_DRIVES
    folder_name = namespace.Title
    return folder_name

def GetInumberFolders(path):
    INumberPaths = []
    for item in listdir(path):                              #for all folders in the user-selected daily data folder                                           
        if OsPath.isdir(f'{path}\\{item}'):                    #if item is a folder
            if search('bioi-\\d+', item.lower()) and (item.lower()[:4]=='bioi') and (search('reinject', item.lower())==None) and (search('bioi-\\d+_.+_\\d+', item.lower())==None):       
            #if item looks like a BioI- folder. Only want to get bioI folders that have been sorted - looking like "BioI-20000" and don't include order folders here. We will handle order folders that are in the main folder separately.
                INumberPaths.append(f'{path}\\{item}')
    return  INumberPaths

#returns a list of paths to order folders in the given path
def GetImmediateOrders(path):
    orderPaths = []
    for item in listdir(path):
        if OsPath.isdir(f'{path}\\{item}'):
            if search('bioi-\\d+_.+_\\d+', item.lower()) and not 'reinject' in item.lower():
                orderPaths.append(f'{path}\\{item}')
    return orderPaths

#function gets pcr folders given a directory.
def GetPCRFolders(path):
    pcrpaths = []
    for item in listdir(path):
        if OsPath.isdir(f'{path}\\{item}'):
            if search('fb-pcr\\d+_\\d+', item.lower()):
                pcrpaths.append(f'{path}\\{item}')
    return pcrpaths



#function gets full paths of the order folders given a BioI folder. 
# returns list of full paths of order folders. 
def GetOrderFolders(plateFolder):
    orderFolderPaths = []

    #I don't think this would capture any by the way the main loop works. Also we'll handle immediate orders(not in order folders) 
    #if search('bioi-\\d+_.+_\\d', plateFolder.lower()): #if "plateFolder" itself is an order
    #    orderFolderPaths.append(f'{plateFolder}')

    for item in listdir(plateFolder):
        if OsPath.isdir(f'{plateFolder}\\{item}'):
            if search('bioi-\\d+_.+_\\d+', item.lower()) and (search('reinject', item.lower())==None):            #if item looks like an order folder BioI-20000_YoMama_123456 Also don't get any that say reinject.
                orderFolderPaths.append(f'{plateFolder}\\{item}')   #add to list
    return orderFolderPaths



def GetDestination(order, path):
    destPath = ''

    #order folders look like: BioI-12345_Name_123456
    if search('bioi-\\d+', order.lower()):
        iNum = search('bioi-\\d+', order.lower()).group(0)
        iNum = search('\\d+', iNum).group(0)

        dataFolderPath = sub(r'/','\\\\', path)
        dayData = dataFolderPath.split('\\')
        dayDataPath = '\\'.join(dayData[:len(dayData)-1])

        for item in listdir(dayDataPath):
            if search(f'bioi-{iNum}', item.lower()) and (search('reinject', item.lower())==None):
                destPath = f'{dayDataPath}\\{item}\\'
                break

    if destPath == '':
        destPath = dayDataPath + '\\'
    
    return destPath

#function checks the order folder to tell if it was mseqed yet, if any ab1 files have braces, and ensures there are ab1 files present. 
def CheckOrder(orderfolder):
    wasMSeqed = False
    areBraces = False
    ab1s = False
    currentProj = []
    mseqSet = {'chromat_dir', 'edit_dir', 'phd_dir', 'mseq4.ini'}
    for item in listdir(orderfolder):
        if item in mseqSet:
            currentProj.append(item)
        if ('{' in item or '}' in item) and item.endswith('.ab1'): 
            areBraces = True
        if item.endswith('.ab1'):
            ab1s = True
    projSet = set(currentProj)

    if mseqSet == projSet: 
        wasMSeqed = True

    return wasMSeqed, areBraces, ab1s


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
    """Connect to existing mSeq instance or start a new one"""
    from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
    from pywinauto import Application, timings
    
    try:
        app = Application(backend='win32').connect(title_re='[mM]seq.*', timeout=1)  # Try existing instance
    except (ElementNotFoundError, timings.TimeoutError):
        try:
            app = Application(backend='win32').start(
                'cmd /c "cd /d C:\\DNA\\Mseq4\\bin && C:\\DNA\\Mseq4\\bin\\j.exe -jprofile mseq.ijl"', 
                wait_for_idle=False)
            app.connect(title='mSeq', timeout=10)
        except Exception as e:
            print(f"Failed to start mSeq: {e}")
            raise
    except ElementAmbiguousError:
        app = Application(backend='win32').connect(title_re='[mM]seq.*', found_index=0)  # Multiple instances found
    
    # Find main window with either capitalization
    main_window = None
    for title_pattern in ['mSeq.*', 'Mseq.*']:
        try:
            main_window = app.window(title_re=title_pattern)
            if main_window.exists(): break
        except: pass
    
    if not main_window or not main_window.exists():
        raise RuntimeError("Could not find mSeq main window")
       
    return app, main_window    
 
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

def wait_for_dialog(app, dialog_type, timeout=5):
    """Wait for a specific dialog to appear"""
    retry_interval = 0.3 if dialog_type in ["error_window", "call_bases"] else 0.1
    
    if dialog_type == "browse_dialog":
        return timings.wait_until(timeout=timeout, retry_interval=retry_interval,
                                func=lambda: IsBrowseDialogOpen(app), value=True)
    elif dialog_type == "preferences":
        return timings.wait_until(timeout=timeout, retry_interval=retry_interval,
                                func=lambda: IsPreferencesOpen(app), value=True)
    elif dialog_type == "copy_files":
        return timings.wait_until(timeout=timeout, retry_interval=retry_interval,
                                func=lambda: IsCopyFilesOpen(app), value=True)
    elif dialog_type == "error_window":
        return timings.wait_until(timeout=timeout, retry_interval=retry_interval,
                                func=lambda: IsErrorWindowOpen(app), value=True)
    elif dialog_type == "call_bases":
        return timings.wait_until(timeout=timeout, retry_interval=retry_interval,
                                func=lambda: IsCallBasesOpen(app), value=True)
    return False

def get_dialog_by_titles(app, titles):
    """Get a dialog window by multiple possible titles"""
    for title in titles:
        try:
            dialog = app.window(title=title)
            if dialog.exists():
                return dialog
        except:
            pass
    return None

def click_dialog_button(dialog, button_titles):
    """Click a button in a dialog using multiple possible titles"""
    for title in button_titles:
        try:
            button = dialog.child_window(title=title, class_name="Button")
            if button.exists():
                button.click_input()
                TimeSleep(0.2)
                return True
        except:
            pass
    return False

def select_all_files_in_dialog(dialog):
    """Select all files in a dialog"""
    try:
        # Try Windows 10 approach first
        shell_view = dialog.child_window(title="ShellView", class_name="SHELLDLL_DefView")
        if shell_view.exists():
            list_view = shell_view.child_window(class_name="DirectUIHWND")
            if list_view.exists():
                list_view.click_input()
                send_keys('^a')  # Ctrl+A
                return True
    except:
        pass
    
    try:
        # Try Windows 11 approach
        list_view = dialog.child_window(class_name="DirectUIHWND")
        if list_view.exists():
            list_view.click_input()
            send_keys('^a')  # Ctrl+A
            return True
    except:
        pass
    
    # Last resort
    try:
        rect = dialog.rectangle()
        dialog.click_input(coords=((rect.right - rect.left) // 2, (rect.bottom - rect.top) // 2))
        TimeSleep(0.2)
        send_keys('^a')  # Ctrl+A
        return True
    except:
        return False
def wait_until_completion(app, orderpath, timeout=45):
    """Wait for mSeq processing to complete"""
    from time import time 
    start_time = time.time()
    interval = 0.2
    
    while time.time() - start_time < timeout:
        # Check for low quality dialog
        if app.window(title="Low quality files skipped").exists():
            return True
            
        # Check for 5 text files (alternative completion signal)
        if CheckFor5txt(orderpath):
            return True
            
        time.sleep(interval)
    
    return False

def close_all_read_info_dialogs(app):
    """Close all Read information dialogs"""
    try:
        from pywinauto import findwindows
        # Find all Read information windows
        read_windows = findwindows.find_elements(
            title_re='Read information for.*', 
            process=app.process
        )
        for win in read_windows:
            try:
                read_dialog = app.window(handle=win.handle)
                read_dialog.close()
                TimeSleep(0.1)
            except Exception as e:
                print(f"Error closing read dialog: {e}")
        return len(read_windows) > 0
    except Exception as e:
        # Fallback approach
        try:
            readWindow = app.window(title_re='Read information for*')
            if readWindow.exists():
                readWindow.close()
                return True
        except:
            pass
        return False
        
def MseqOrder(orderpath):
    """Process folder with mSeq using more reliable methods"""
    from pywinauto.keyboard import send_keys
    
    # Connect to mSeq and get main window
    app, mseqMainWindow = GetMseq()
    mseqMainWindow.set_focus()
    
    # Start new project (Ctrl+N)
    send_keys('^n')
    print("Starting new project")
    
    # Handle Browse For Folder dialog
    wait_for_dialog(app, "browse_dialog")
    dialog_window = get_dialog_by_titles(app, ['Browse For Folder', 'Browse for Folder'])
    if not dialog_window:
        dialog_window = app.window(title_re='Browse.*Folder')
    
    if not dialog_window:
        print("Could not get Browse For Folder dialog")
        return False
    
    # Navigate to the target folder
    TimeSleep(0.5)  # Add delay for first browsing operation
    if not NavigateToFolder(dialog_window, orderpath):
        print(f"Failed to navigate to {orderpath}")
        return False
    
    # Click OK button
    click_dialog_button(dialog_window, ["OK", "&OK"])
    
    # Handle mSeq Preferences dialog
    wait_for_dialog(app, "preferences")
    pref_dialog = get_dialog_by_titles(app, ['Mseq Preferences', 'mSeq Preferences'])
    if pref_dialog:
        click_dialog_button(pref_dialog, ["&OK", "OK"])
    
    # Handle Copy sequence files dialog
    wait_for_dialog(app, "copy_files")
    copy_dialog = app.window(title_re='Copy.*sequence files')
    if copy_dialog:
        select_all_files_in_dialog(copy_dialog) # Select all files
        click_dialog_button(copy_dialog, ["&Open", "Open"])
    wait_for_dialog(app, "error_window") # Handle File error dialog
    error_dialog = get_dialog_by_titles(app, ['File error', 'Error'])
    if error_dialog:
        click_dialog_button(error_dialog, ["OK"])

    wait_for_dialog(app, "call_bases") # Handle Call bases dialog
    call_bases_dialog = app.window(title_re='Call bases.*')
    if call_bases_dialog:
        click_dialog_button(call_bases_dialog, ["&Yes", "Yes"])

    # Wait for processing to complete
    wait_until_completion(app, orderpath)

    # Close any Read information windows before returning
    close_all_read_info_dialogs(app)
    print(f"Successfully processed {orderpath}")
    return True

def NavigateToFolder(browsefolderswindow, path):
    """Navigate through folder tree with improved reliability"""
    import re
    browsefolderswindow.set_focus()
    treeView = browsefolderswindow.child_window(class_name="SysTreeView32")
    if not treeView.exists(): return False
    # Get Desktop item and expand it
    item = treeView.get_item('\\Desktop')
    item.click_input()
    TimeSleep(0.2)
    item.expand()
    TimeSleep(0.3)
    # Find This PC under Desktop (name varies by Windows version)
    this_pc_item = None
    for child in item.children():
        if any(term in child.text().lower() for term in ["pc", "computer"]):
            this_pc_item = child
            break
    
    if not this_pc_item: return False
    # Expand This PC
    this_pc_item.click_input()
    TimeSleep(0.2)
    this_pc_item.expand()
    TimeSleep(0.3)
    # Parse path and handle drive first
    data = path.split('\\')
    if len(data) > 0:
        drive = data[0]
        # Handle mapped network drives
        if 'P:' in drive: drive = sub('P:', r'ABISync (P:)', drive)
        if 'H:' in drive: drive = sub('H:', r'Tom (\\\\w2k16\\users) (H:)', drive)
        # Find the drive in This PC children
        drive_item = None
        for child in this_pc_item.children():
            drive_text = child.text()
            # Extract drive letter for better matching
            drive_letter = None
            if ":" in drive_text:
                drive_letter_parts = re.findall(r'([A-Za-z]:)', drive_text)
                if drive_letter_parts: drive_letter = drive_letter_parts[0]
            # Try multiple matching approaches
            if (drive_text == drive or drive in drive_text or 
                (drive_letter and drive_letter.upper() == drive.upper()[:2])):
                drive_item = child
                break
        if not drive_item: return False
        # Navigate folder hierarchy
        current_item = drive_item
        current_item.click_input()
        TimeSleep(0.2)
        for folder in data[1:]:
            if not folder: continue  # Skip empty segments
            current_item.expand()
            TimeSleep(0.3)
            # Find next folder (exact match first, then partial)
            next_item = None
            for child in current_item.children():
                if child.text() == folder:
                    next_item = child
                    break
            if not next_item:
                for child in current_item.children():
                    if folder.lower() in child.text().lower():
                        next_item = child
                        break
            if not next_item: return False
            next_item.click_input()
            TimeSleep(0.2)
            current_item = next_item
    # Ensure final folder is selected
    current_item.click_input()
    return True
    
#End functions 
########################################################################################################


########################################################################################################
#Main code
if __name__ == "__main__":
    batPath = "P:\\generate-data-sorting-key-file.bat"
    try:
        # SubprocessRun(batPath, shell=True, check=True)
        keyFilePathName = 'C:\\P\\order_key.txt'
        reinjectListFileName = f'Reinject List_{get_formatted_date()}.xlsx'
        reinjectListFilePath = f'C:\\P:\\Data\\Reinjects\\{reinjectListFileName}'
        completeReinjectList = []
        reinjectPrepList = []
        completeRawReinjectList = []
        key = loadtxt(keyFilePathName, dtype=str, delimiter='\t')
        AdjustFullKeyToABIChars

        dataFolderPath = filedialog.askdirectory(title="Select today's data folder. To mseq orders.")
        dataFolderPath = sub(r'/','\\\\', dataFolderPath)
        bioIFolders  = GetInumberFolders(dataFolderPath)  #get bioIfolders that have been sorted already
        immediateOrderFolders = GetImmediateOrders(dataFolderPath)
        pcrFolders = GetPCRFolders(dataFolderPath)

        from pywinauto import Application
        from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
        from pywinauto.keyboard import send_keys
        from pywinauto import timings
        

        #loop through each bioI folder
        for bioIFolder in bioIFolders:
            orderFolders = GetOrderFolders(bioIFolder) #get order folders which will be compressed.
            for orderFolder in orderFolders:
                if not 'andreev' in orderFolder.lower():                            #don't mseq dmitri's orders
                    orderNumber = GetOrderNumberFromFolderName(orderFolder)
                    orderKey = GetOrderList(orderNumber)
                    allab1Files = GetAllab1Files(orderFolder)
                    #mseq ab1 files.
                    mseqed, braces, yesab1 = CheckOrder(orderFolder)
                    if not mseqed and not braces:
                        if len(orderKey)==len(allab1Files): #if there are the right number of files, mseq
                            if yesab1:
                                MseqOrder(orderFolder)
                                print(f'Mseq completed      {OsPath.basename(orderFolder)}')
                        else:
                            if OsPath.exists(f'{dataFolderPath}\\IND Not Ready')==False:
                                mkdir(f'{dataFolderPath}\\IND Not Ready')
                            destination = f'{dataFolderPath}\\IND Not Ready\\'
                            ShutilMove(orderFolder, destination)
                            print(f'Order moved      {OsPath.basename(orderFolder)}')
                else:  #for Dmitri, we don't mseq but still want to check for completion to move his order folder if needed.
                    orderNumber = GetOrderNumberFromFolderName(orderFolder)
                    orderKey = GetOrderList(orderNumber)
                    allab1Files = GetAllab1Files(orderFolder)
                    if len(orderKey) != len(allab1Files):
                        if OsPath.exists(f'{dataFolderPath}\\IND Not Ready')==False:
                            mkdir(f'{dataFolderPath}\\IND Not Ready')
                        destination = f'{dataFolderPath}\\IND Not Ready\\'
                        ShutilMove(orderFolder, destination)
                        print(f'Order moved      {OsPath.basename(orderFolder)}')


        #this loop will handle order folders that are in the day data folder if that's what's selected. 
        #otherwise, it will handle order folders in the IND Not Ready - move them out to their appropriate I number folders after mseq
        INDNotReadyIsTheSelectedFolder = True if OsPath.basename(dataFolderPath)=='IND Not Ready' else False
        for orderFolder in immediateOrderFolders:
            if not 'andreev' in orderFolder.lower():                            #don't mseq dmitri's orders
                orderNumber = GetOrderNumberFromFolderName(orderFolder)
                orderKey = GetOrderList(orderNumber)
                allab1Files = GetAllab1Files(orderFolder)
                #mseq ab1 files.
                mseqed, braces, yesab1 = CheckOrder(orderFolder)
                if not mseqed and not braces:
                    if len(orderKey)==len(allab1Files): #if there are the right number of files, mseq
                        if yesab1:
                            MseqOrder(orderFolder)
                            print(f'Mseq completed      {OsPath.basename(orderFolder)}')
                        if INDNotReadyIsTheSelectedFolder:
                            #if we're mseqing the IND Not Ready folder, and we've just completed and mseqed the order, we move it back to where it belongs
                            destination = GetDestination(OsPath.basename(orderFolder), dataFolderPath)
                            app, mseqMainWindow = GetMseq()
                            mseqMainWindow = None
                            app.kill()  
                            TimeSleep(.1)
                            ShutilMove(orderFolder, destination)

                    else:
                        if INDNotReadyIsTheSelectedFolder:
                            #if user has selected IND Not Ready, and the order is still not complete, we won't be mseqing, and won't move it, it will stay where it is 
                            #in the IND Not Ready Folder
                            pass
                        else:
                            if OsPath.exists(f'{dataFolderPath}\\IND Not Ready')==False:
                                mkdir(f'{dataFolderPath}\\IND Not Ready')
                            destination = f'{dataFolderPath}\\IND Not Ready\\'
                            ShutilMove(orderFolder, destination)
                            print(f'Order moved      {OsPath.basename(orderFolder)}')

                elif mseqed and not braces:
                    if INDNotReadyIsTheSelectedFolder:
                        #if we're mseqing the IND Not Ready folder, and we've just completed and mseqed the order, we move it back to where it belongs
                        destination = GetDestination(OsPath.basename(orderFolder), dataFolderPath)
                        app, mseqMainWindow = GetMseq()
                        mseqMainWindow = None
                        app.kill()  
                        TimeSleep(.1)
                        ShutilMove(orderFolder, destination)
            else: #if we have a Dmitri Order in the IND Not Ready folder, we still don't mseq but want to check
                if INDNotReadyIsTheSelectedFolder:
                    orderNumber = GetOrderNumberFromFolderName(orderFolder)
                    orderKey = GetOrderList(orderNumber)
                    allab1Files = GetAllab1Files(orderFolder)
                    if len(orderKey)==len(allab1Files): #the order folder has the right number of files, then we move it back to its I number folder
                        destination = GetDestination(OsPath.basename(orderFolder), dataFolderPath)
                        ShutilMove(orderFolder, destination)
                    else: 
                        #leave it there
                        pass
                else: 
                    #if it's in the selected folder, just leave it 
                    pass



                            
        for pcrFolder in pcrFolders:
            #print(pcrFolder)
            mseqed, braces, yesab1 = CheckOrder(pcrFolder)
            if not mseqed and not braces and yesab1:
                MseqOrder(pcrFolder)
                print(f'Mseq completed      {OsPath.basename(pcrFolder)}')
            else:
                print(f'Mseq NOT completed      {OsPath.basename(pcrFolder)} ')
        
        print('')
        print('ALL DONE')

        #close mseq
        app, mseqMainWindow = GetMseq()
        mseqMainWindow = None
        app.kill()

    except CalledProcessError:
        print(f'Batch file, {batPath} didn\'t work.')
    except FileNotFoundError:
        print('File not found.')