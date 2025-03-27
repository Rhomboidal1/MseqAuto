from subprocess import CalledProcessError, run as SubprocessRun
from datetime import date, datetime, timedelta
from os import listdir, scandir, path as OsPath, replace, mkdir
from re import search, sub
from numpy import loadtxt, where, array
from tkinter import filedialog, Button, Label, Tk
from glob import glob
from pylightxl import readxl
from sys import exit as SysExit

#FORMATTING FUNCTIONS
#Function returns a version of the string without anything contained in {} ex: {01A}Sample_Primer{5_25}{I-12345} returns Sample_Primer
def CleanBracesFormat(aList):
    i = 0
    newList = []
    while i < len(aList):
        newList.append(sub('{.*?}', '', NeutralizeSuffixes(aList[i])))
        #newList.append(sub('{.*?}', '', aList[i]))
        i+=1
    return newList

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

#Need to compare sample_primer names to names adjusted for our "Premixed" and "RTI" endings.
def NeutralizeSuffixes(fileName):
    newFileName = fileName
    newFileName = newFileName.replace('_Premixed', '')
    newFileName = newFileName.replace('_RTI', '')
    return newFileName

def AdjustKeyToABIChars():
    i = 0
    while i < len(subKey):
        subKey[i][3] = AdjustABIChars(subKey[i][3])
        i+=1
def AdjustFullKeyToABIChars():
    i = 0
    while i < len(key):
        key[i][3] = AdjustABIChars(key[i][3])
        i+=1
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

def previous_business_day(date):
    previous_day = date - timedelta(days=1)
    while previous_day.weekday() > 4:  # Monday to Friday are 0-4
        previous_day -= timedelta(days=1)
    return previous_day
#End Functions used for formatting
##############################################################################################################



##############################################################################################################
#FUNCTIONS USED FOR INTERPRETING INFO AND SORTING FILES
#Function gets only the information needed based on the I numbers that were in the selected folder.
def GetSubKey(INumbers):
    subkey = []
    i = 0
    while i < len(key):
        if key[i][0] in INumbers: 
            subkey.append(key[i].tolist())
        i+=1
    return subkey

#Function determines if PCR#### is in the file name
def GetPCRNumber(fileName):
    #return str(search(f'PCR\\d+', fileName).group())
    if search('{pcr\\d+.+}', fileName.lower()):
         pcr = str(search("{pcr\\d+.+}", fileName.lower()).group()).upper()  #get the bracket information {PCR123_exp1}
         pcr = str(search("PCR\\d+", pcr).group())                               #then get only PCR123 out of that string.
    else:
        pcr = ''
    return pcr

#Function needs to get a list of the reaction names from reinject*.txt files(that's cells B6 through B101 in excel)     
def GetReinjectList(INumbers):
    i = 0
    reinjectList = []
    #print(len(INumbers))
    while i < len(INumbers):
        #print(f'G:\\Lab\\Spreadsheets\\*{str(INumbers[i])}*reinject*txt')
        if glob(f'G:\\Lab\\Spreadsheets\\*{str(INumbers[i])}*reinject*txt'):
            fileName = glob(f'G:\\Lab\\Spreadsheets\\*{str(INumbers[i])}*reinject*txt')[0]
            #data = np.loadtxt(fileName, dtype=str, delimiter='\t')
            data = loadtxt(fileName, dtype=str, delimiter='\t')
            j = 5
            while j < 101:
                reinjectList.append(data[j,1])
                j+=1
             
        if glob(f'G:\\Lab\\Spreadsheets\\Individual Uploaded to ABI\\*{str(INumbers[i])}*reinject*txt'):
            #there could be cases where there are multiple versions of this reinject .txt. could have a warning or something for user input to select which file they want to use.  
            fileName = glob(f'G:\\Lab\\Spreadsheets\\Individual Uploaded to ABI\\*{str(INumbers[i])}*reinject*txt')[0]
            #data = np.loadtxt(fileName, dtype=str, delimiter='\t')
            data = loadtxt(fileName, dtype=str, delimiter='\t')
            j = 5
            while j < 101:
                reinjectList.append(data[j,1])
                j+=1
        i+=1


    return reinjectList

#function looks through today's I number txt files and compares the data/sample/soon-to-be-file names to the P drive reinject list to determine 
#reinjects. This is needed because there could be reinjects on a BioI-### plate rather than the BioI-###_Reinject plate. These still need to be detected and sorted appropriately.
def GetReinjectsFromPartialPlates(INumbers):
    reinjectList = []
    i = 0
    #print(len(INumbers))
    while i < len(INumbers):
        #print(f'G:\\Lab\\Spreadsheets\\*{str(INumbers[i])}*reinject*txt')
        if glob(f'G:\\Lab\\Spreadsheets\\*{str(INumbers[i])}*txt'):
            fileNames = glob(f'G:\\Lab\\Spreadsheets\\*{str(INumbers[i])}*txt')
            for file in fileNames:
                #data = np.loadtxt(file, dtype=str, delimiter='\t')
                data = loadtxt(file, dtype=str, delimiter='\t')
                j = 5
                while j < 101:
                    #if data[j,1] is in masterreinjectlist nparray then append
                    #print(data[j,1][5:])
                    #keyIndex = np.where(reinjectPrepList == data[j,1][5:])[0]
                    keyIndex = where(reinjectPrepList == data[j,1][5:])[0]
                    if len(keyIndex) > 0:
                        reinjectList.append(data[j,1])
                    j+=1
             
        if glob(f'G:\\Lab\\Spreadsheets\\Individual Uploaded to ABI\\*{str(INumbers[i])}*txt'):
            #there could be cases where there are multiple versions of this reinject .txt. could have a warning or something for user input to select which file they want to use.  
            fileNames = glob(f'G:\\Lab\\Spreadsheets\\Individual Uploaded to ABI\\*{str(INumbers[i])}*txt')
            for file in fileNames:
                #print(file)
                #data = np.loadtxt(file, dtype=str, delimiter='\t')
                data = loadtxt(file, dtype=str, delimiter='\t')
                j = 5
                while j < 101:
                    #if data[j,1] is in masterreinjectlist nparray then append
                    #print(data[j,1][5:])
                    #keyIndex = np.where(reinjectPrepList == str(data[j,1][5:]))[0]
                    keyIndex = where(reinjectPrepList == str(data[j,1][5:]))[0]
                    #print(keyIndex)
                    if len(keyIndex) > 0:
                        reinjectList.append(data[j,1])
                    j+=1
             
        i+=1
    return reinjectList


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

#function returns two lists: a list of I numbers and a list of directory paths for those I numbers
def GetTodaysInumbersFromFolder(path):
    INumbers = []
    INumberPaths = []
    
    for item in listdir(path):                              #for all folders in the user-selected daily data folder
        #print(item)                                             
        if OsPath.isdir(f'{path}\\{item}'):                    #if item is a folder
            if search('bioi-\\d+', item.lower()):       #if item looks like a BioI- folder. This is true pre or post sorting and renaming.
                INumbers.append(GetINumberFromFolderName(item))
            if search('.+bioi-\\d+.+', item.lower()) and not 'reinject' in item.lower():      #this is a BioI- folder before it has been sorted. Avoid reinject folders. 
                INumberPaths.append(f'{path}\\{item}')
    return INumbers, INumberPaths


#function finds the most recent I number based on the folders contained in P:\Data\Individuals
def GetMostRecentINumberFromDataInd(path):
    try: 
        current_timestamp = datetime.now().timestamp()
        cutoff_timestamp = current_timestamp - (7 * 24 * 3600)  # 24 hours * 3600 seconds
        folders = []
        
        with scandir(path) as entries:
            for entry in entries:
                
                if entry.is_dir():
                    last_modified_timestamp = entry.stat().st_mtime
                    if last_modified_timestamp >= cutoff_timestamp:
                        #print(entry.name)
                        folders.append(entry.name)

        # Sort the list by last modified timestamp (newest to oldest)
        sorted_folder_info_list = sorted(folders, reverse=True)
        #print(sorted_folder_info_list)
        inum = GetINumberFromFolderName(sorted_folder_info_list[0])
        return inum
    except Exception as e:
        return f"Error: {str(e)}"

def GetINumbersGreaterThan(txtFiles, Inumber):
    inumsGreaterThan = []
    for file in txtFiles:
        inum = GetINumberFromFolderName(file)
        if inum and Inumber:  # Check that both values are not None
            try:
                if int(inum) > int(Inumber):
                    inumsGreaterThan.append(file)
            except ValueError:
                continue  # Skip if conversion to int fails
    return inumsGreaterThan

#Function returns a list of names for the order folders. 
def GetOrderFolderNames(Inumber):
    folderList = []
    usekey = False
    #InumIndexes = np.where((subKey == str(Inumber)))
    InumIndexes = where((subKey == str(Inumber)))
    if len(InumIndexes[0]) == 0:
            InumIndexes = where((key == str(Inumber)))
            usekey = True
    i = 0
    while i < len(InumIndexes[0]):  
        if usekey:
            iNumber = key[InumIndexes[0][i], 0]
            acctName = key[InumIndexes[0][i], 1]
            orderNumber = key[InumIndexes[0][i], 2]  
        else:
            iNumber = subKey[InumIndexes[0][i], 0]
            acctName = subKey[InumIndexes[0][i], 1]
            orderNumber = subKey[InumIndexes[0][i], 2]    
        name = f'BioI-{iNumber}_{acctName}_{orderNumber}'
        if str(name) not in folderList:
            folderList.append(name)
        i+=1
    return folderList

#function returns a list of txt files from the last number of days
def get_files_within_last_days(directories, num_days):
    try:
        current_timestamp = datetime.now().timestamp()
        # Change to use calendar days instead of seconds to handle year boundaries
        cutoff_date = datetime.now() - timedelta(days=num_days)
        cutoff_timestamp = cutoff_date.timestamp()

        file_info_list = []
        for directory in directories:
            with scandir(directory) as entries:
                for entry in entries:
                    if entry.is_file():
                        last_modified_timestamp = entry.stat().st_mtime
                        # Compare using datetime objects to handle year boundary correctly
                        last_modified_date = datetime.fromtimestamp(last_modified_timestamp)
                        if last_modified_date >= cutoff_date and entry.name.endswith('.txt'):
                            file_info_list.append((entry.name, last_modified_timestamp))

        sorted_file_info_list = sorted(file_info_list, key=lambda x: x[1], reverse=True)
        newest_file_names = [file_info[0] for file_info in sorted_file_info_list]
        return newest_file_names

    except Exception as e:
        return f"Error: {str(e)}"

    
#function returns a list of txt files 
def GettxtFilesFromLast12Hrs(directories):
    try:
        current_date = datetime.now()
        cutoff_date = current_date - timedelta(hours=12)

        file_info_list = []
        for directory in directories:
            with scandir(directory) as entries:
                for entry in entries:
                    if entry.is_file():
                        last_modified_date = datetime.fromtimestamp(entry.stat().st_mtime)
                        if last_modified_date >= cutoff_date and entry.name.endswith('.txt'):
                            file_info_list.append((entry.name, entry.stat().st_mtime))

        sorted_file_info_list = sorted(file_info_list, key=lambda x: x[1], reverse=True)
        newest_file_names = [file_info[0] for file_info in sorted_file_info_list]
        return newest_file_names

    except Exception as e:
        return f"Error: {str(e)}"

#function renames folders from the ABI format to our clean format: 2024-03-22_BioI-20644
def UpdateFolderName(path):
    base = OsPath.basename(path)
    newBase = f'BioI-{GetINumberFromFolderName(base)}'
    parentDir = path.replace(base, newBase)
    return parentDir

def GetNumbers(files):
    numbers = []
    for file in files:
        numbers.append(GetINumberFromFolderName(file))
    return numbers
#function will only be called when there is something in the list
def AllINumbersAreSame(list):
    allSame = True
    if len(list) >=1:
        match = list[0]
    for number in list: 
        if match != number:
            allSame = False
    return allSame
def AllDifferentINumbers(list):
    return len(list) == len(set(list))
def GetSpecifiedInum(fileName):
    iNumFromFileName = ''
    if search('{i.\\d+}', fileName.lower()):
        iNumFromFileName = search('{i.\\d+}', fileName.lower()).group(0)
        iNumFromFileName = search('\\d+', iNumFromFileName).group(0)
    return iNumFromFileName
def GetFolder(path, number, orderFolderNeeded):
    foldername = ''
    for entry in scandir(path):
        if entry.is_dir():
            if search(f'bioi-{number}', entry.name.lower()):
                if search(f'bioi-\\d+_.+_\\d+', entry.name.lower()) and entry.name.lower() != str(orderFolderNeeded):
                    pass
                else:
                    #makes a preference for I number folders already completed. THis is in cases like BioI-20000a and BioI-20000b
                    #if foldername is like f'bioi-\\d+_.+_\\d+' but entry.name is like f'bioi-{number}' and not like f'bioi-\\d+_.+_\\d+':
                    '''if search(f'bioi-\\d+_.+_\\d+', foldername) and (search( f'bioi-{number}', entry.name.lower()) and not search(f'bioi-\\d+_.+_\\d+', entry.name.lower())):
                        foldername = entry.name
                    else:'''
                        
                    foldername = entry.name
    return foldername

def GetPCRFolder(path, pcrnumber):
    foldername = ''
    for entry in scandir(path):
        if entry.is_dir():
            if search(f'FB-{pcrnumber}', entry.name):
                foldername = entry.name
                break
    else:
        if foldername == '':
            foldername = f'FB-{pcrnumber}'
        
    return foldername

def GetSpecifiedOrderNum(fileName):
    orderNumFromFileName = ''
    if search('{\\d+}', fileName):
        orderNumFromFileName = search('{\\d+}', fileName.lower()).group(0)
        orderNumFromFileName = search('\\d+', orderNumFromFileName).group(0) #removes the brackets.
    return orderNumFromFileName


#function determines if a file can be sorted, and if so returns information about the sample and where to sort it to.
def DetermineSortDestination(keyToUse, fileToSort, cleanName, currentInum, recurse):
    recursionLimit = 1
    inum = ''
    accountName = ''
    orderNumber = ''
    orderFolder = ''

    #keyIndexes = np.where(keyToUse == cleanName)[0]
    '''print(keyToUse)
    for x in keyToUse:
        print(x)'''
    keyIndexes = where(keyToUse == cleanName)[0]
    found = False
    inumIndexList = []
    for index in keyIndexes:
        inum = keyToUse[index][0]
        inumIndexList.append(inum)
    
    if len(keyIndexes) > 1: #multiple matches
        if AllDifferentINumbers(inumIndexList):
            inum = GetSpecifiedInum(fileToSort)
            if inum != '' : #if the I number has been specified. This means the lab tech who set up the reaction concatenated the Inumber to the name upon setup. Then match to specified Inumber
                for index in keyIndexes:
                    if keyToUse[index][0] ==inum: 
                        #we already have the I number
                        accountName = keyToUse[index][1]
                        orderNumber = keyToUse[index][2]
                        
                        orderF = f'BioI-{inum}_{accountName}_{orderNumber}'

                        parentFolder = GetFolder(dataFolderPath, inum, orderF)
                        if parentFolder ==''or parentFolder == orderF:
                            orderFolder = f'{dataFolderPath}\\{orderF}'
                        else:
                            orderFolder = f'{dataFolderPath}\\{parentFolder}\\{orderF}'
                        found = True
                        break
                #else: #if we never find the inumber that was specified. Because it was typed wrong, or maybe it was an exoI-number. 
                    #don't sort the file. Let user figure it out.
            else: #no I number has been specified. Match to currentInumber
                for index in keyIndexes:
                    if keyToUse[index][0] ==currentInum: 
                        #we already have the I number
                        accountName = keyToUse[index][1]
                        orderNumber = keyToUse[index][2]
                        orderF = f'BioI-{currentInum}_{accountName}_{orderNumber}'

                        parentFolder = GetFolder(dataFolderPath, currentInum, orderF)
                        if parentFolder ==''or parentFolder == {orderF}:
                            orderFolder = f'{dataFolderPath}\\{orderF}'
                        else:
                            orderFolder = f'{dataFolderPath}\\{parentFolder}\\{orderF}'
                        found = True
        else: #there are one or more duplicate Inumbers where the reaction was found in the key
            orderNumber = GetSpecifiedOrderNum(fileToSort)
            if orderNumber != '': #if the order number has been specified. This means the lab tech who set up the reaction concatenated the order number to the name upon setup.
                for index in keyIndexes:
                    if keyToUse[index][2]==orderNumber:
                        inum = keyToUse[index][0]
                        accountName = keyToUse[index][1]
                        
                        #there shouldn't ever be a case where get folder returns nothing because the subkey in which current file was 
                        #located was generated from the BioI numbers in the dataFolderPath. 
                        orderF = f'BioI-{inum}_{accountName}_{orderNumber}'

                        parentFolder = GetFolder(dataFolderPath, inum, orderF)
                        if parentFolder ==''or parentFolder == orderF:
                            orderFolder = f'{dataFolderPath}\\{orderF}'
                        else:
                            orderFolder = f'{dataFolderPath}\\{parentFolder}\\{orderF}'
                        found = True
            else: #order number was not specified. and the reaction is found to have a duplicate on the same I number.
                #can't sort these
                pass
    elif len(keyIndexes)==1:
        inum = keyToUse[keyIndexes][0][0]
        accountName = keyToUse[keyIndexes][0][1]
        orderNumber = keyToUse[keyIndexes][0][2]
        orderF = f'BioI-{inum}_{accountName}_{orderNumber}'

        parentFolder = GetFolder(dataFolderPath, inum, orderF)
        if parentFolder =='' or parentFolder == orderF:
            orderFolder = f'{dataFolderPath}\\{orderF}'
        else:
            orderFolder = f'{dataFolderPath}\\{parentFolder}\\{orderF}'
        found = True
    else: #len(keyindexes) < 1: 
        #print('recurse')
        #recurse and search the entire key.
        if recurse < recursionLimit:
            AdjustFullKeyToABIChars()
            found, inum, accountName, orderNumber, orderFolder = DetermineSortDestination(key, fileToSort, cleanName, currentInum, recurse + 1)

    return found, inum, accountName, orderNumber, orderFolder

class RetryReinjects():
    def __init__(self):
        self.root = Tk()
        #self.mainframe = tk.Frame(self.root)
        self.root.title('No Reinjects Found!')
        self.label = Label(self.root, text="You can leave this window open, go make the Reinject txt file or update a partial plate BioI txt file then click proceed \n OR click proceed without making a txt file \n OR click Quit to end the entire script without anything being sorted.")
        proceedButton = Button(self.root, text='Proceed', command=self.Proceed)
        quitButton = Button(self.root, text='Quit', command=self.Quit)
        self.label.pack()
        proceedButton.pack()
        quitButton.pack()
        self.root.mainloop()
        return
    def Proceed(self):
        self.root.destroy()
    def Quit(self):
        SysExit()
#End functions used for sorting
###############################################################################################################



###############################################################################################################
#Begin Main Code:

if __name__ == "__main__":
    batPath = "P:\\generate-data-sorting-key-file.bat"
    try: 
        #this is the batch file that updates our order key in the P drive. 
        SubprocessRun(batPath, shell=True, check=True)
        dataFolderPath = filedialog.askdirectory(title="Select today's data folder.")
        keyFilePathName = 'P:\\order_key.txt'
        #keyFilePathName = 'P:\\order_key - testing.txt'
        reinjectListFileName = f'Reinject List_{GetFormattedDate()}.xlsx'
        reinjectListFilePath = f'P:\\Data\\Reinjects\\{reinjectListFileName}'
        #reinjectListFilePath = 'P:\\Data\\Reinjects\\Reinject List_7-1-2024.xlsx'
        #reinjectListFilePath = f'H:\\{reinjectListFileName}'
        completeReinjectList = []
        reinjectPrepList = []

        completeRawReinjectList = []

        #key = np.loadtxt(keyFilePathName, dtype=str, delimiter='\t')
        key = loadtxt(keyFilePathName, dtype=str, delimiter='\t')
        controls = [
            '_pGEM_T7Promoter', 
            '_pGEM_SP6Promoter',
            '_pGEM_M13F-20',
            '_pGEM_M13R-27',
            '_Water_T7Promoter',
            '_Water_SP6Promoter',
            '_Water_M13F-20',
            '_Water_M13R-27'
        ]
        
        #step 1) get all the BioI folders for today
        #todaysINumbers is a list of only I number folders that need to be sorted, NOT all the I numbers in the folder, ironically, it's not actually all today's I numbers.
        iNumbersToSort, bioIFolders  = GetTodaysInumbersFromFolder(dataFolderPath)
        #print(iNumbersToSort)

        
        indUptoABIPath = 'G:\\Lab\\Spreadsheets\\Individual Uploaded to ABI'
        spreadsheetsPath = 'G:\\Lab\\Spreadsheets'
        paths = [indUptoABIPath, spreadsheetsPath]
        recent_files = get_files_within_last_days(paths, 5)
        #print(recent_files)
        dataIndPath = 'P:\\Data\\Individuals'
        lowerINum = GetMostRecentINumberFromDataInd(dataIndPath)
        #print('lower I number from Data Individuals.')
        #print(lowerINum)
        finalInumsFilesNames = GetINumbersGreaterThan(recent_files, lowerINum)
        #print(finalInumsFilesNames)
        finalInums = GetNumbers(finalInumsFilesNames)
        #print(finalInums)
        #this will force the inclusion of all txt files made within 12 hrs regardless of if the I number was excluded by the GetInumbersGreaterThan Function 
        anyAdditionalInums = GetNumbers(GettxtFilesFromLast12Hrs(paths))
        for num in anyAdditionalInums:
            if not num in finalInums:
                finalInums.append(num)
        #print(finalInums)
        
        #print(finalInums)
        #finalInums = ['21024', '21025', '21026','21027']
        #subKey = np.array(GetSubKey(finalInums))
        subKey = array(GetSubKey(finalInums))
        #print(subKey)
        AdjustKeyToABIChars()
        #need a list of all the I numbers in the folder so that we can use all of them to find any and all reinject txt files, not limited to only the folders that are being sorted. 
        rawReinjectList = GetReinjectList(finalInums)
        #print(reinjectList)


        reinjectList = CleanBracesFormat(rawReinjectList)
        #print(reinjectList)
        if OsPath.exists(reinjectListFilePath):
            reinjectListDB = readxl(reinjectListFilePath, ws='Sheet1')
            #reinjectListDB = xl.readxl(f'P:\\Data\\Reinjects\\testing.xlsx', ws='Sheet1')
            sheet = reinjectListDB.ws('Sheet1')
            for row in sheet.rows:
                reinjectPrepList.append(row[0] + row[1])
        
        #reinjectPrepList = np.array(reinjectPrepList)
        reinjectPrepList = array(reinjectPrepList)  #make a numpy array
        
        #print(reinjectPrepList)
        completeReinjectList.extend(reinjectList)
        completeRawReinjectList.extend(rawReinjectList)
        #print(completeReinjectList)

        rawReinjectsFromPotentialPartialPlates = GetReinjectsFromPartialPlates(finalInums)
        #print(reinjectsFromPotentialPartialPlates)
        completeRawReinjectList.extend(rawReinjectsFromPotentialPartialPlates)
        reinjectsFromPotentialPartialPlates = CleanBracesFormat(rawReinjectsFromPotentialPartialPlates)
        #print(reinjectsFromPotentialPartialPlates)
        completeReinjectList.extend(reinjectsFromPotentialPartialPlates)
        
        
        #finally check if we got a reinject, if not give user another shot.
        if len(completeReinjectList)==0 and len(reinjectPrepList) != 0: 
            RetryReinjects()

            #rerun all the Reinject code.
            rawReinjectList = GetReinjectList(finalInums)
            reinjectList = CleanBracesFormat(rawReinjectList)
            completeRawReinjectList.extend(rawReinjectList)
            completeReinjectList.extend(reinjectList)
            rawReinjectsFromPotentialPartialPlates = GetReinjectsFromPartialPlates(finalInums)
            completeRawReinjectList.extend(rawReinjectsFromPotentialPartialPlates)
            reinjectsFromPotentialPartialPlates = CleanBracesFormat(rawReinjectsFromPotentialPartialPlates)
            completeReinjectList.extend(reinjectsFromPotentialPartialPlates)

        #print(completeReinjectList)
        j = 1 
        orderFileList = []
        pcrFileList = []
        #step 2) go through the BioI folders and sort out the files in them
            #folder path includes drive lettter to base directory base    
        for folder in bioIFolders:
            
            #prepare all the order folders first. Rare cases where the files are not recognize or sorted at all, 
                    #we still want to have an order folder available for manual sorting
            currentInumber = GetINumberFromFolderName(folder)
            orderFolderNames = GetOrderFolderNames(currentInumber)
            for name in orderFolderNames: #loop to make the folders.
                if OsPath.exists(f'{folder}\\{name}')==False:
                    mkdir(f'{folder}\\{name}')

            #loop to sort each file in the current folder. 
            for file in listdir(folder):
                if OsPath.isfile(f'{folder}\\{file}'):
                    pcrNumber = GetPCRNumber(file)                                      #will return '' if there is no {PCRxxxx} in the file name
                    cleanFileName = AdjustABIChars(file)[:-4]                           #the file name with characters replaced just like VBA excel txt file maker splice -4 for ".ab1"
                    cleanFileName = NeutralizeSuffixes(cleanFileName)                   #removes suffixes, premixed and rti
                    cleanFileName = sub('{.*?}', '', cleanFileName)                  #remove all brace contents
                    #print(cleanFileName)
                    #print(subKey)
                    #keyIndexes = np.where(subKey == cleanFileName)[0]                   #this is list of indices where the file name is found in the keyFile

                    #if (cleanFileName in fileList)==False: #only do something with the file if it's a file we haven't seen before; this means, do nothing with preemptive reinjects.
                    if(cleanFileName in controls):                                     #controls ab1 file
                        inum = GetINumberFromFolderName(folder)
                        destinationPath = f'{folder}\\Controls'
                        if OsPath.exists(f'{destinationPath}')==False:
                            mkdir(f'{destinationPath}')
                        replace(f'{folder}\\{file}', f'{destinationPath}\\{file}')
                        #print(f'{j} {folder}\\{file} placed to {destinationPath}\\{file}')
                        j+=1

                    elif cleanFileName == '':                                           #blank ab1 file  
                        inum = GetINumberFromFolderName(folder)
                        destinationPath = f'{folder}\\Blank'
                        if OsPath.exists(f'{destinationPath}')==False:
                            mkdir(f'{destinationPath}')
                        replace(f'{folder}\\{file}', f'{destinationPath}\\{file}')
                        #print(f'{j} {folder}\\{file} placed to {destinationPath}\\{file}')
                        j+=1
                        
                    elif pcrNumber != '':                                               #PCR ab1 file
                        PCRDataFolder = f'{dataFolderPath}\\{GetPCRFolder(dataFolderPath, pcrNumber)}'
                        #PCRDataFolder = f'{dataFolderPath}\\FB-{pcrNumber}'
                        if ([PCRDataFolder, cleanFileName] in pcrFileList)==False:
                            cleanBraceFileName = sub('{.*?}', '', file)
                            if OsPath.exists(PCRDataFolder)==False:
                                mkdir(PCRDataFolder)
                            if OsPath.exists(f'{PCRDataFolder}\\{cleanBraceFileName}'):
                                if OsPath.exists(f'{PCRDataFolder}\\Alternate Injections')==False:    
                                    mkdir(f'{PCRDataFolder}\\Alternate Injections')     
                                destinationPath = f'{PCRDataFolder}\\Alternate Injections\\{file}'
                                replace(f'{folder}\\{file}', destinationPath) 

                            if (cleanFileName in completeReinjectList): #if was reinjected, put file in alternate injections folder
                                if OsPath.exists(f'{PCRDataFolder}\\Alternate Injections')==False: #make sure the folder exists before putting it there.
                                    mkdir(f'{PCRDataFolder}\\Alternate Injections')
                                indexInList = completeReinjectList.index(cleanFileName)
                                rawName = completeRawReinjectList[indexInList]
                                if '{!P}' in rawName: #but actually don't put it in the injection folder if labeled as a preemptive reinject.
                                    destinationPath = f'{PCRDataFolder}\\{cleanBraceFileName}'
                                else:
                                    destinationPath = f'{PCRDataFolder}\\Alternate Injections\\{file}'
                                replace(f'{folder}\\{file}', destinationPath)
                                #print(f'{j} {folder}\\{file} placed to {PCRDataFolder}\\Alternate Injections\\{file}')
                                j+=1
                            else: #if was not reinject, put file regular PCR folder.
                                destinationPath = f'{PCRDataFolder}\\{cleanBraceFileName}'
                                replace(f'{folder}\\{file}', destinationPath)
                                #print(f'{j} {folder}\\{file} placed to {PCRDataFolder}\\{file}')
                                j+=1
                            pcrFileList.append([PCRDataFolder, cleanFileName])
                    else:                                                              #a customer file
                        sort, inum, accountName, orderNumber, orderFolder = DetermineSortDestination(subKey, file, cleanFileName, currentInumber, 0)
                        if sort:        #if sort was returned as true, then do sorting things
                            if ([orderFolder, cleanFileName] in orderFileList)==False:            
                                cleanBraceFileName = sub('{.*?}', '', file)
                                if OsPath.exists(orderFolder)==False: 
                                    mkdir(orderFolder)

                                #try:
                                #if the file exists in the folder already
                                if OsPath.exists(f'{orderFolder}\\{cleanBraceFileName}'):
                                    if OsPath.exists(f'{orderFolder}\\Alternate Injections')==False:    
                                        mkdir(f'{orderFolder}\\Alternate Injections')     
                                    destinationPath = f'{orderFolder}\\Alternate Injections\\{file}'
                                    replace(f'{folder}\\{file}', destinationPath)                           #put the file into that alt inj folder without changing the name
                                
                                #if file is in the reinject list we'll want to put this in alt injections folder
                                elif (cleanFileName in completeReinjectList) == True:
                                    if OsPath.exists(f'{orderFolder}\\Alternate Injections')==False:    
                                        mkdir(f'{orderFolder}\\Alternate Injections')                                                     #make that folder   
                                    indexInList = completeReinjectList.index(cleanFileName)
                                    rawName = completeRawReinjectList[indexInList]
                                    if '{!P}' in rawName: #but actually don't put it in the injection folder if labeled as a preemptive reinject.
                                        destinationPath = f'{orderFolder}\\{cleanBraceFileName}'
                                    else:
                                        destinationPath = f'{orderFolder}\\Alternate Injections\\{file}'
                                    replace(f'{folder}\\{file}', destinationPath)                           #put the file into that alt inj folder without changing the name
                                    #print(f'{j} {folder}\\{file} placed to {destinationPath}\\{file}')
                                    j+=1
                                else:
                                #except ValueError:
                                    destinationPath = f'{orderFolder}\\{cleanBraceFileName}'
                                    replace(f'{folder}\\{file}', destinationPath)          #put the file into the main order folder, plus rename it excluding all brace contents
                                    #print(f'{j} {folder}\\{file} placed to {orderfolder}\\{cleanBraceFileName}')
                                    j+=1                    
                                #fileList.append([orderfolder, cleanFileName])
                                orderFileList.append([orderFolder, cleanFileName])
                        else: #sort returned false
                            print(f'{j} file {file} not found.')
                            j+=1
            #rename the folder
            newFolderName = UpdateFolderName(folder)
            if OsPath.exists(newFolderName)==False:
                replace(folder, newFolderName)
    except CalledProcessError:
        print(f'Batch file, {batPath} didn\'t work.')
    except FileNotFoundError:
        print('File not found.')