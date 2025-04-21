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
from re import search, sub
from time import sleep as TimeSleep
from datetime import date
from numpy import loadtxt, where

#function that gets today's date in format m-d-yyyy
def GetFormattedDate():
    today = date.today()
    #gets rid of leading 
    if today.strftime('%m')[:1]=='0':
        m = today.strftime('%m')[1:]
    else:
        m = today.strftime('%m')

    if today.strftime('%d')[:1]=='0':
        d = today.strftime('%d')[1:]
    else:
        d = today.strftime('%d')
    y = today.strftime('%Y')

    formatteDate= f'{m}-{d}-{y}'
    return formatteDate

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
    
    #mseqMainWindow = app.window(title_re='mSeq*')
    
    app, mseqMainWindow = GetMseq()
    mseqMainWindow.set_focus()
    send_keys('^n')

    timings.wait_until(timeout=5, retry_interval=.1, func=lambda: IsBrowseDialogOpen(app), value=True)
    dialogWindow = app.window(title='Browse For Folder')
    TimeSleep(.5)
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

    timings.wait_until(timeout=10, retry_interval=.3, func=lambda: IsErrorWindowOpen(app), value=True)
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
    #desktopPath = f'\\Desktop\\This PC'
    item = treeView.get_item(desktopPath)
    data = path.split('\\')
    for folder in data:
        if 'P:' in folder:
            #folder = sub('P:', r'abisync (\\\\w2k16) (P:)', folder         OLD DRIVE NAME
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
    
#End functions 
########################################################################################################


########################################################################################################
#Main code
if __name__ == "__main__":
    batPath = "P:\\generate-data-sorting-key-file.bat"
    try:
        SubprocessRun(batPath, shell=True, check=True)
        keyFilePathName = 'P:\\order_key.txt'
        reinjectListFileName = f'Reinject List_{GetFormattedDate()}.xlsx'
        reinjectListFilePath = f'P:\\Data\\Reinjects\\{reinjectListFileName}'
        completeReinjectList = []
        reinjectPrepList = []
        completeRawReinjectList = []
        key = loadtxt(keyFilePathName, dtype=str, delimiter='\t')
        AdjustFullKeyToABIChars

        #root = tk.Tk()
        #dialog = tk.Toplevel(root)
        #dataFolderPath = filedialog.askdirectory(parent=dialog, title="Select today's data folder. To mseq orders.")
        dataFolderPath = filedialog.askdirectory(title="Select today's data folder. To mseq orders.")
        #dialog.destroy()
        dataFolderPath = sub(r'/','\\\\', dataFolderPath)
        bioIFolders  = GetInumberFolders(dataFolderPath)  #get bioIfolders that have been sorted already
        immediateOrderFolders = GetImmediateOrders(dataFolderPath)
        #print(bioIFolders)
        pcrFolders = GetPCRFolders(dataFolderPath)

        from pywinauto import Application
        from pywinauto.findwindows import ElementNotFoundError, ElementAmbiguousError
        from pywinauto.keyboard import send_keys
        from pywinauto import timings
        

        #loop through each bioI folder
        for bioIFolder in bioIFolders:
            orderFolders = GetOrderFolders(bioIFolder) #get order folders which will be compressed.
            #print(orderFolders)
            for orderFolder in orderFolders:
                #print(orderFolder)
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