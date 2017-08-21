from openpyxl import Workbook #load_workbook, Workbook # reads .xlsx/.xlsm files
from xlrd import * # Reads .xls
import xlwt
import os
import shutil



# Use stored data
def getXlsData(fileName ,sheetNum ):
    '''This function takes a file name and a sheet number. It opes the file at the sheet that was
    specified. The Sheet numbers start at 0.'''
    mainData = []

    with open_workbook(fileName) as mainOpen:
       sheet4 = mainOpen.sheet_by_index(sheetNum)
       mainData = [[sheet4.cell_value(r,c) for c in range(sheet4.ncols)] for r in range(sheet4.nrows)]
    return mainData

def columnNames( fileInfo , row):
    '''This returns the column names of the file. You ave to specify where the column names are.'''
    return fileInfo[row]
    
def enumeratedDict(columnNameList):
    adict = {}
    for i in range(len(columnNameList)):
        adict[columnNameList[i]] = i
    return adict

def copyTemplate(fileDir,tempName, nameOfOutput):
    current = os.getcwd()
    shutil.copy(fileDir, current)
    copyDir = os.path.join(current,tempName)
    os.rename(copyDir, nameOfOutput)
    copyDir = os.path.join(current,nameOfOutput)
    return copyDir

    

    


if __name__ == '__main__':

    #This is the info that i will need from the use
    #Have a function that subtracts 1 from the user input for proper indexing sheet number and row number
    fileLocal = "/Users/cristiangalindo/Documents/workspace/Krustallos/MainFiles/B&Bitems169-263.xls"
    fileSheet = 3
    fileColNameRow = 9

    templateName = "FlatFileHome.xlsm"
    templatExtension = templateName.split('.')[1]
    template = os.path.join(os.getcwd(),"Templates", templateName)
    #template = "/Users/cristiangalindo/Documents/workspace/Krustallos/Templates/FlatFileHome.xlsm"
    tempSheet = 3
    tempColNameRow = 1

    rename = "Dummy"

    #Here is the program
    #New workbook
    
    #########
    #Store DATA
    fileColDict = {}
    fileData = getXlsData(fileLocal,fileSheet)
    fileColName = columnNames(fileData, fileColNameRow)
    fileDict = enumeratedDict(fileColName)

    
    #Creates a copy of the template
    dirOfCopy = copyTemplate(template,templateName, rename+"."+templatExtension)
    print(dirOfCopy)

    for k in fileDict.keys():
        print(k,"   ",fileDict[k])
        if k == "Sell By Unit":
            print("IT DID GET IT!!!!!")
            print("\n")
    print("DATA NAMES")
    print("\n")
    print(fileColName)
    print("DATA DICT")
    print("\n")
    print(fileDict)
    print("\n")

    templateData = getXlsData(dirOfCopy,tempSheet)
    tempColName = columnNames(templateData, tempColNameRow)
    tempDict = enumeratedDict(tempColName)
    print("Template Names")
    print(tempColName)
    print("\n")
    print("Template DICT")
    print("\n")
    print(tempDict)
    print("\n")

    [("SKU","ItemName")]




    
