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

    

#function zips ab1 files and the 5 txt files.
#returns nothing
def ZipOrder(dirPath, zip_filename):
    endingList = ['.ab1', '.raw.qual.txt','.raw.seq.txt','.seq.info.txt','seq.qual.txt','.seq.txt']
    endingList = tuple(endingList)
    with ZipFile(zip_filename, 'w') as myzip:
        for entry in scandir(dirPath):
            if entry.name.endswith(endingList):
                print(f'{entry.name} added')
                myzip.write(entry.path, arcname=entry.name, compress_type=ZIP_DEFLATED)

#function zips fsa files only. 
def Zipfsa(dirPath, zip_filename):
    with ZipFile(zip_filename, 'w') as myzip:
        for entry in scandir(dirPath):
            if entry.name.endswith('.fsa'):
                print(f'{entry.name} added')
                myzip.write(entry.path, arcname=entry.name, compress_type=ZIP_DEFLATED)
#End functions used for sorting
###############################################################################################################



##############################################################################################################
#Begin Main code
if __name__ == "__main__":
    dataFolderPath = filedialog.askdirectory(title="Select today's data folder. To zip .ab1 and .txt files in plate folders")
    dataFolderPath = sub(r'/','\\\\', dataFolderPath)
    zipDumpFolder = f'{dataFolderPath}\\zip dump\\'
    if OsPath.exists(zipDumpFolder)==False:
        mkdir(zipDumpFolder)
    PlateFolders  = GetPlateFolders(dataFolderPath)  #get bioIfolders that have been sorted already
    #print(bioIFolders)

    #loop through each bioI folder
    for folder in PlateFolders:
        zipFileExist = CheckForZip(folder)         #first check if there is a zip file. 
        if not zipFileExist:                            #if not then do zip/compressing things. 
            ab1Files = GetAllab1Files(folder)
            #print(ab1Files)
            txtFilesData = Get5txtFiles(folder)

            ab1Check = True if len(ab1Files)>0 else False       
            
            if txtFilesData[0] == True :
                txtCheck = True 
                txtFiles = txtFilesData[1]
            else:
                txtCheck = False

            if ab1Check and txtCheck:
                filesToCompress = []

                filesToCompress.extend(ab1Files)
                filesToCompress.extend(txtFiles)

                dirPath = f'{folder}'
                zipNamePath = f'{folder}\\{OsPath.basename(folder)}.zip'

                #zip that sheet
                ZipOrder(dirPath, zipNamePath)
                ShutilCopy(zipNamePath, f'{zipDumpFolder}{OsPath.basename(zipNamePath)}')

                i = 0 
                while i < 5: 
                    print('')
                    i +=1
                print(f'{OsPath.basename(folder)} completed')
                i = 0 
                while i < 5: 
                    print('')
                    i +=1
            else:
                fsaFiles = CheckForfsaFiles(folder)
                if fsaFiles: 
                    dirPath = f'{folder}'
                    zipNamePath = f'{folder}\\{OsPath.basename(folder)}.zip'
                    Zipfsa(dirPath, zipNamePath)
                    ShutilCopy(zipNamePath, f'{zipDumpFolder}{OsPath.basename(zipNamePath)}')
                    
                    i = 0 
                    while i < 5: 
                        print('')
                        i +=1
                    print(f'{OsPath.basename(folder)} completed')
                    i = 0 
                    while i < 5: 
                        print('')
                        i +=1
    #if nothing was put into zipdump folder then delete the folder.
    if not listdir(zipDumpFolder):
        rmdir(zipDumpFolder)
                    
