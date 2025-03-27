from tkinter import filedialog
from os import listdir, path as OsPath, mkdir, replace, rename
from re import search, sub

def GetPNumberFolders(path):
    folderPaths = []
    for item in listdir(path):
        if OsPath.isdir(f'{path}\\{item}'):
                if search('p\\d+.+', item.lower()):
                    folderPaths.append(f'{path}\\{item}')
    return folderPaths

if __name__ == "__main__":
    controlsList = [
        'pgem_m13f-20',
        'water_m13f-20'
    ]
    controlFolderName = 'Controls'
    BlankFolderName = 'Blank'

    dataFolderPath = filedialog.askdirectory(title="Select today's data folder.")
    pNumberFolders = GetPNumberFolders(dataFolderPath)
    print(pNumberFolders)
    
    for folder in pNumberFolders:
        for item in listdir(folder):
            itemPath = f'{folder}\\{item}'
            if OsPath.isfile(f'{folder}\\{item}') and item.endswith('.ab1'):
                print(f'{item}    {item.lower()[4:]}')
                
                if item.lower()[4:-4] in controlsList:
                    destination = f'{folder}\\{controlFolderName}'
                    if not OsPath.exists(destination):
                        mkdir(destination)
                    replace(itemPath, f'{destination}\\{item}') #subfolder into control folder
                    
                elif len(item)==9: # 01A__.ab1 
                    destination = f'{folder}\\{BlankFolderName}'
                    if not OsPath.exists(destination):
                        mkdir(destination)
                    replace(itemPath, f'{destination}\\{item}') #subfolder into blank folder
                
                #if item has {} then rename item. 
                if ('{' in item and '}' in item): 
                    cleanBraceFileName = sub(r'{.*?}', '', item)
                    rename(itemPath, f'{folder}\\{cleanBraceFileName}')
