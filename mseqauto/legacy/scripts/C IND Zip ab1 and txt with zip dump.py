#This script allows user to select from a dialog a folder that contains bioI folder which contain order folders. 
#The script will search the order folders for ab1 files and the 5 mseq-created .txt files
#If the order folder has these things, it'll compress them into a .zip file, saved into the order folder
#special formatting is incorporated for Dmitri Andreev


from tkinter import filedialog
from os import listdir, path as OsPath, scandir, mkdir, rmdir
from re import search, sub
from zipfile import ZipFile, ZIP_DEFLATED
from shutil import copyfile as ShutilCopy
##############################################################################################################
#Begin Functions

#function gets full paths of the BioI folders which contain order folders which ultimately contain ab1 and txt file to compress
#returns list of full paths
def GetInumberFolders(path):
    INumberPaths = []
    for item in listdir(path):                              #for all folders in the user-selected daily data folder                                           
        if OsPath.isdir(f'{path}\\{item}'):                    #if item is a folder
            if search('bioi-\\d+', item.lower()):       #if item looks like a BioI- folder. Only want to get bioI folders that have been sorted - looking like "BioI-20000"
                INumberPaths.append(f'{path}\\{item}')
    return  INumberPaths

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

    if search('bioi-\\d+_.+_\\d', plateFolder.lower()): #if "plateFolder" itself is an order
        orderFolderPaths.append(f'{plateFolder}')

    for item in listdir(plateFolder):
        if OsPath.isdir(f'{plateFolder}\\{item}'):
            if search('bioi-\\d+_.+_\\d', item.lower()):            #if item looks like an order folder BioI-20000_YoMama_123456
                orderFolderPaths.append(f'{plateFolder}\\{item}')   # add to list
    return orderFolderPaths

#function iterates items in given folder(order folder) 
# if a .zip is found, return false, else true
def CheckForZip(folder):
    zipExists = False
    for item in listdir(folder):
            if OsPath.isfile(f'{folder}\\{item}'):
                if item.endswith('.zip'):
                    zipExists = True
    return zipExists

#function finds all the .ab1 files in a given orderfolder
#returns list of full paths to those .ab1 files
def GetAllab1Files(folder):
    ab1FilePaths = []
    for item in listdir(folder):
        if OsPath.isfile(f'{folder}\\{item}'):
            #print(item)
            if item.endswith('.ab1'):
                ab1FilePaths.append(f'{folder}\\{item}')
    return ab1FilePaths

#function finds the 5 txt files in a given orderfolder
#returns list: first element is boolean if all 5 .txt files are found
#              second element is list of full paths to those .txt files
def Get5txtFiles(folder):
    fivetxts = 0 #must count 5 txt files. 
    txtFilePaths =[]
    


    if not 'andreev' in folder.lower(): #if cutomers other than Dmitri 
        for item in listdir(folder):
            if OsPath.isfile(f'{folder}\\{item}'):
                #print(item)
                if item.endswith('.raw.qual.txt'):
                    txtFilePaths.append(f'{folder}\\{item}')
                    fivetxts +=1
                elif item.endswith('.raw.seq.txt'):
                    txtFilePaths.append(f'{folder}\\{item}')
                    fivetxts +=1
                elif item.endswith('.seq.info.txt'):
                    txtFilePaths.append(f'{folder}\\{item}')
                    fivetxts +=1
                elif item.endswith('.seq.qual.txt'):
                    txtFilePaths.append(f'{folder}\\{item}')
                    fivetxts +=1
                elif item.endswith('.seq.txt'):
                    txtFilePaths.append(f'{folder}\\{item}')
                    fivetxts +=1
        all5 = True if (fivetxts == 5) else False
        txtFileData = [all5, txtFilePaths]

        return txtFileData
    else:
        all5=True #pretend we have all txt files
        #txtFilePaths remains empty cuz we don't want to include any txt files in zip archive for Dmitri :)
        txtFileData = [all5, txtFilePaths]
        return txtFileData
    
#function finds and returns just the order number of the BioI folder name
def GetOrderNumberFromFolderName(folder):
    orderNum = ''
    if search('.+_\\d+', folder):
        orderNum = search('_\\d+', folder).group(0)
        orderNum = search('\\d+', orderNum).group(0)
        return orderNum
    else:
        return ''


#function finds and returns just the I number of the BioI folder name
def GetINumberFromFolderName(folderName):
    #BioString = folderName[folderName.lower().find('bio'):]
    
    if search('bioi-\\d+', str(folderName).lower()):
        BioString = search('bioi-\\d+', str(folderName).lower()).group(0)
        Inumber = search('\\d+', BioString).group(0)
        return Inumber
    else:
        return ''
    

#function zips ab1 files and the 5 txt files.
#returns nothing
def ZipOrder(dirPath, zip_filename):
    endingList = ['.ab1', '.raw.qual.txt','.raw.seq.txt','.seq.info.txt','seq.qual.txt','.seq.txt']
    endingList = tuple(endingList)
    with ZipFile(zip_filename, 'w') as myzip:
        for entry in scandir(dirPath):
            if entry.name.endswith(endingList):
                myzip.write(entry.path, arcname=entry.name, compress_type=ZIP_DEFLATED)

#function zips ab1 files only. Used for Dmitiri
def ZipOrderNoTxt(dirPath, zip_filename):
    with ZipFile(zip_filename, 'w') as myzip:
        for entry in scandir(dirPath):
            if entry.name.endswith('.ab1'):
                print(f'{entry.name} added')
                myzip.write(entry.path, arcname=entry.name, compress_type=ZIP_DEFLATED)

def printComplete(name):
    i = 0 
    while i < 5: 
        print('')
        i +=1
    print(f'{OsPath.basename(name)} completed')
    i = 0 
    while i < 5: 
        print('')
        i +=1

#End functions used for sorting
###############################################################################################################



##############################################################################################################
#Begin Main code
if __name__ == "__main__":
    dataFolderPath = filedialog.askdirectory(title="Select today's data folder. To zip .ab1 and .txt files in order folders")
    dataFolderPath = sub(r'/','\\\\', dataFolderPath)
    zipDumpFolder = f'{dataFolderPath}\\zip dump\\'
    if OsPath.exists(zipDumpFolder)==False:
        mkdir(zipDumpFolder)
    bioIFolders  = GetInumberFolders(dataFolderPath)  #get bioIfolders that have been sorted already

    #print(bioIFolders)
    pcrFolders = GetPCRFolders(dataFolderPath)

    #loop through each bioI folder
    for bioIFolder in bioIFolders:
        orderFolders = GetOrderFolders(bioIFolder) #get order folders which will be compressed.
        #print(orderFolders)
        for orderFolder in orderFolders:
            
            zipFileExist = CheckForZip(orderFolder)         #first check if there is a zip file. 
            if not zipFileExist:                            #if not then do zip/compressing things. 
                ab1Files = GetAllab1Files(orderFolder)
                #print(ab1Files)
                txtFilesData = Get5txtFiles(orderFolder)

                ab1Check = True if len(ab1Files)>0 else False       
                
                if txtFilesData[0] == True :
                    txtCheck = True 
                    txtFiles = txtFilesData[1]
                    #print(txtFiles)
                else:
                    txtCheck = False

                if ab1Check and txtCheck:
                    filesToCompress = []

                    filesToCompress.extend(ab1Files)
                    filesToCompress.extend(txtFiles)

                    dirPath = f'{orderFolder}'
                    if not 'andreev' in orderFolder.lower():
                        #zip file will look like "BioI-20000_YoMama_123456.zip"
                        zipNamePath = f'{orderFolder}\\{OsPath.basename(orderFolder)}.zip'
                        ZipOrder(dirPath, zipNamePath)
                        ShutilCopy(zipNamePath, f'{zipDumpFolder}{OsPath.basename(zipNamePath)}')
                        printComplete(orderFolder)
                    else:
                        #Dmitir wants zip to look like "123456_I-20000.zip" ... old school
                        #Mseq should be skipped, but in case it wasn't, and there are txt files, we ignore that and zip ab1 only.
                        orderNumber = GetOrderNumberFromFolderName(orderFolder)
                        Inumber = GetINumberFromFolderName(orderFolder)
                        #print(f'dmitri: {orderNumber}   {Inumber}')
                        zipNamePath = f'{orderFolder}\\{orderNumber}_I-{Inumber}.zip'
                        ZipOrderNoTxt(dirPath, zipNamePath)
                        ShutilCopy(zipNamePath, f'{zipDumpFolder}\\{OsPath.basename(zipNamePath)}')
                        printComplete(orderFolder)
                    
    
    for pcrFolder in pcrFolders:
        zipFileExist = CheckForZip(pcrFolder)         #first check if there is a zip file. 
        if not zipFileExist:                            #if not then do zip/compressing things. 
            ab1Files = GetAllab1Files(pcrFolder)
            #print(ab1Files)
            txtFilesData = Get5txtFiles(pcrFolder)

            ab1Check = True if len(ab1Files)>0 else False
            if txtFilesData[0] == True :
                txtCheck = True 
                txtFiles = txtFilesData[1]
                #print(txtFiles)
            else:
                txtCheck = False

            if ab1Check and txtCheck:
                filesToCompress = []

                filesToCompress.extend(ab1Files)
                filesToCompress.extend(txtFiles)    
                zipNamePath = f'{pcrFolder}\\{OsPath.basename(pcrFolder)}.zip'
                dirPath = pcrFolder
                ZipOrder(dirPath, zipNamePath)
                ShutilCopy(zipNamePath, f'{zipDumpFolder}{OsPath.basename(zipNamePath)}')
                printComplete(pcrFolder)
    #if nothing was put into zipdump folder then delete the folder.
    if not listdir(zipDumpFolder):
        rmdir(zipDumpFolder)
