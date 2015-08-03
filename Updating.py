from openpyxl import Workbook #creats writable Workbook
from xlrd import * # Reads .xls
from xlutils.copy import copy
from collections import OrderedDict, defaultdict
import shutil
import os
from PopulateAmazonTNoS import columnNames, invEnumerated

'''The user is allowed to specify what type of info he is passing.The user should be able to
use these two options:
-Inventory
-Sold
For the inventory option, the user is passing the file will all the inventory.
(I need to make sure all the numbers are correct. This options gives me how many are avaliable).

For the sold the user will pass on a template of the items that were sold.
(I need to add the information to the sold column and calculate the new ammount avaliable).'''



def inventoryDict(sheetInfo, dictionary, items,start):
    alist = [dictionary[items[0]],dictionary[items[1]]]
    adict = OrderedDict()
    cols = sheetInfo.ncols
    rows = sheetInfo.nrows
    for r in range(start,rows):
        currentRowItem = sheetInfo.cell_value(r , alist[0])
        currentRowAmm = sheetInfo.cell_value(r , alist[1])
        adict[currentRowItem] = currentRowAmm
    return adict
            
                
    

if __name__ == '__main__':


    amazonInfoFileName = "AmazonShort.xls"
    amazonInfoType = "Inventory"
    amazonInfoLocal = os.path.join(os.getcwd(), "ToUpdate", "Shortened", amazonInfoFileName)
    amazonInfoSheet = 1
    aColNameStart = 1
    aInfoStart = 2

    ebayInfoFileName = "EbayShort.xls"
    ebayInfoType = "Inventory"
    ebayInfoLocal = os.path.join(os.getcwd(), "ToUpdate", "Shortened", ebayInfoFileName)
    ebayInfoSheet = 1
    eColNameStart = 1
    eInfoStart = 2

    websiteInfoFileName = "WebsiteShort.xls"
    websiteInfoType = "Inventory"
    websiteInfoLocal = os.path.join(os.getcwd(), "ToUpdate", "Shortened", websiteInfoFileName)
    websiteInfoSheet = 1
    wColNameStart = 1
    wInfoStart = 2

    posInfoFile = "POS.xls"
    posInfoType = "Inventory"
    posInfoLocal = os.path.join(os.getcwd(), "ToUpdate", "Shortened", posInfoFile)
    posInfoSheet = 1
    pColNameStart = 1
    pInfoStart = 2
    
    inventoryFileName = "AllInventory.xls"
    inventoryInfoLocal = os.path.join(os.getcwd(), "MainFiles", "InventoryFile", inventoryFileName)
    invSheetNum = 4
    invColNameStart = 10
    invInfoStart = 11

    
    #Here is where i change the indexes to match
    amazonInfoSheet = amazonInfoSheet - 1
    aColNameStart = aColNameStart - 1
    aInfoStart = aInfoStart - 1

    ebayInfoSheet = ebayInfoSheet - 1
    eColNameStart = eColNameStart - 1
    eInfoStart = eInfoStart - 1
    
    websiteInfoSheet = websiteInfoSheet - 1
    wColNameStart = wColNameStart - 1
    wInfoStart = wInfoStart - 1

    posInfoSheet = posInfoSheet - 1
    pColNameStart = pColNameStart - 1
    pInfoStart = pInfoStart - 1

    invSheetNum = invSheetNum - 1
    invColNameStart = invColNameStart - 1
    invInfoStart = invInfoStart - 1
    #end of index change



    aInfoCols = ["sku","quantity"]
    amazonColNames = columnNames(amazonInfoLocal, aColNameStart, amazonInfoSheet)
    amazonInfoOpen = open_workbook(amazonInfoLocal)
    amazonSheet = amazonInfoOpen.sheet_by_index(amazonInfoSheet)
    amaRows = amazonSheet.nrows
    amaCols = amazonSheet.ncols
    amazonInv = inventoryDict(amazonSheet, amazonColNames, aInfoCols, aInfoStart)
##    print("Amazon INV")
##    print(amazonInv)
##    print()


    eInfoCols = ["Custom Label", "Quantity Available"]
    ebayColNames = columnNames(ebayInfoLocal, eColNameStart, ebayInfoSheet)
    ebayInfoOpen = open_workbook(ebayInfoLocal)
    ebaySheet = ebayInfoOpen.sheet_by_index(ebayInfoSheet)
    ebayRows = ebaySheet.nrows
    ebayCols = ebaySheet.ncols
    ebayInv = inventoryDict(ebaySheet, ebayColNames, eInfoCols, eInfoStart)
##    print("Ebay INV")
##    print(ebayInv)
##    print()

    

    webInfoCols = ["Product ID", "Stock"]
    webColNames = columnNames(websiteInfoLocal, wColNameStart, websiteInfoSheet)
    websiteInfoOpen = open_workbook(websiteInfoLocal)
    websiteSheet = websiteInfoOpen.sheet_by_index(websiteInfoSheet)
    webRows = websiteSheet.nrows
    webCols = websiteSheet.ncols
    webInv = inventoryDict(websiteSheet, webColNames, webInfoCols, wInfoStart)
##    print("Website INV")
##    print(webInv)  
##    print()

    posInfoCols = ["Item Name", "Qty 1"]
    posColNames = columnNames(posInfoLocal, pColNameStart,posInfoSheet)
    posInfoOpen = open_workbook(posInfoLocal)
    posSysSheet = posInfoOpen.sheet_by_index(posInfoSheet)
    posRows = posSysSheet.nrows
    posCols = posSysSheet.ncols
    posInv = inventoryDict(posSysSheet, posColNames, posInfoCols, pInfoStart)
##    print("POS Inv")
##    print(posInv)
##    print()
    
    newWorkbook = Workbook()
    wSheet = newWorkbook.active

################################################################################
####HERE IS WHERE IT GETS ALL THE INVENTORY AND CREATES A NEW TEMPLATE####


    incorrectDict = defaultdict()
    
    with open_workbook(inventoryInfoLocal) as openInventory:
        inventorySheet = openInventory.sheet_by_index(invSheetNum)
        rows = inventorySheet.nrows
        cols = inventorySheet.ncols
        
        columnNames = columnNames(inventoryInfoLocal, invColNameStart, invSheetNum)
        revColNames = invEnumerated(columnNames)

        ### MAKE A DICT OF NAMES THAT DON'T EXIST IN THE STORE FILES BUT WAS PUT ON THE MAIN FILE
        
        for k,v in columnNames.items():
            
            currentCell = wSheet.cell(row = 1, column = v + 1)
            currentCell.value = k






        parentDict = OrderedDict()
        children = 0
        print(columnNames)
        parentColNum = 0

        soldPos = 0
        soldAma = 0
        soldWeb = 0
        soldEbay = 0
        totalSold = 0
        finalAmount = 0

        relationCol = int(columnNames["Parents/Child/None"])
        posCol = int(columnNames["POS Item Name"])
        amazonCol = int(columnNames["Amazon Sku"])
        websiteCol = int(columnNames["Website Item Name"])
        ebayCol = int(columnNames["Ebay Custom Label"])
        searsCol = int(columnNames["Sears Name"])
        typeCol = int(columnNames["Type"])
        descripCol = int(columnNames["Short Description"])
        
        startQty = int(columnNames["Starting Quantity"])
        soldStoreCol = int(columnNames["Sold in Store"])
        soldAmaCol = int(columnNames["Sold in Amazon"])
        soldEbayCol = int(columnNames["Sold in Ebay"])
        soldWebCol = int(columnNames["Sold in Website"])
        soldSearsCol = int(columnNames["Sold in Sears"])
        totalCol = int(columnNames["Total Sold"])
        finalQtyCol = int(columnNames["Qty 1"])

        
        for r in range(invInfoStart,rows):

            parOrChild = inventorySheet.cell_value(r , relationCol )
            currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = relationCol +1)
            currentCell.value = parOrChild
                    
            if parOrChild == "Parent":
                

                
                if children != 0:
                    
                    itemName = inventorySheet.cell_value(parentColNum , posCol)
                    startingQty = inventorySheet.cell_value(parentColNum , startQty)
                    
                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2 ,column = totalCol +1)
                    currentCell.value = totalSold
                    
                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2  ,column = soldStoreCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldStoreCol)
                    currentCell.value = soldBefore + soldPos

                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2  ,column = soldAmaCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldAmaCol)
                    currentCell.value = soldBefore + soldAma

                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2  ,column = soldEbayCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldEbayCol)
                    currentCell.value = soldBefore + soldEbay

                    currentCell = wSheet.cell(row = (parentColNum - invInfoStart + 2) ,column = finalQtyCol + 1)
                    currentCell.value = startingQty - (soldPos + soldAma + soldEbay)
                    parentDict[itemName] = startingQty - (soldPos + soldAma + soldEbay)
                    

                    soldHis = 0
                    soldPos = 0
                    soldAma = 0
                    soldEbay = 0
                    children = 0
                
                if children == 0:
                    parentColNum = r
                    #Calculates the amount already sold of that item
                    startingQty = inventorySheet.cell_value(r , startQty)
                    itemName = inventorySheet.cell_value(parentColNum , posCol)
                    p = inventorySheet.cell_value(r , soldStoreCol)
                    a = inventorySheet.cell_value(r , soldAmaCol)
                    w = inventorySheet.cell_value(r , soldWebCol)
                    e = inventorySheet.cell_value(r , soldEbayCol)
                    s = inventorySheet.cell_value(r , soldSearsCol)

                    
                    soldHis = p+a+w+e+s

                    #Fills in the item name and starting quantity for the parent
                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = posCol +1)
                    currentCell.value = itemName

                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = startQty +1)
                    currentCell.value = inventorySheet.cell_value(r,startQty)

                    
                    

                
            if parOrChild == "Child":
                children += 1
                #Fills in type and description for children
                currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = typeCol +1)
                currentCell.value = inventorySheet.cell_value(r,typeCol)

                currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = descripCol +1)
                currentCell.value = inventorySheet.cell_value(r,descripCol)

                currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = startQty +1)
                currentCell.value = inventorySheet.cell_value(r,startQty)
                ## Calculates the inv for the pos system
                try:
                    
                    item = inventorySheet.cell_value(r,posCol)
                    if item not in ["", " "]:
                        posInventory = posInv[item]
                        
                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = posCol +1)
                        currentCell.value = item
                        
                        startAmount = inventorySheet.cell_value(r , startQty)
                        alreadySold = inventorySheet.cell_value(r , soldStoreCol)
                        soldInPos = (startAmount - posInventory) + alreadySold

                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldStoreCol +1)
                        currentCell.value = soldInPos
                        soldPos += soldInPos
                    else:
                        raise KeyError


                except KeyError:
                    if 'pos' in incorrectDict.keys():
                        incorrectDict['pos'].append(inventorySheet.cell_value(r,posCol))
                    else:
                        incorrectDict['pos'] = [inventorySheet.cell_value(r,posCol)]

                ## calculates the children inv for amazon
                try:
                    
                    item = inventorySheet.cell_value(r,amazonCol)
                    
                    if item not in ["", " "]:
                        amazonInventory = amazonInv[item]
                        
                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = amazonCol +1)
                        currentCell.value = item

                        startAmount = inventorySheet.cell_value(r , startQty)
                        alreadySold = inventorySheet.cell_value(r , soldAmaCol)
                        soldInAma = (startAmount - amazonInventory) + alreadySold

                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldAmaCol +1)
                        currentCell.value = soldInAma
                        soldAma += soldInAma
                    else:
                        raise KeyError

                    
                except KeyError:
                    if 'amazon' in incorrectDict.keys():
                        incorrectDict['amazon'].append(inventorySheet.cell_value(r,amazonCol))
                    else:
                        incorrectDict['amazon'] = [inventorySheet.cell_value(r,amazonCol)]

                
                ## calculates the children inv for Website
                ##This one is special
                try:
                    item = inventorySheet.cell_value(r,websiteCol)
                    websiteInventory = webInv[item]
                    parent = inventorySheet.cell_value(r,websiteCol)
                    
                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = websiteCol +1)
                    currentCell.value = item
                    
                except KeyError:
                    if 'website' in incorrectDict.keys():
                        incorrectDict['website'].append(inventorySheet.cell_value(r,nameCol))
                    else:
                        incorrectDict['website'] = [inventorySheet.cell_value(r,nameCol)]

                        
                ## calculates the children inv for Ebay
                try:
                    item = inventorySheet.cell_value(r,ebayCol)
                    if item not in ["", " "]:
                        ebayInventory = ebayInv[item]
                        
                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = ebayCol +1)
                        currentCell.value = item


                        startAmount = inventorySheet.cell_value(r , startQty)
                        alreadySold = inventorySheet.cell_value(r , soldEbayCol)
                        soldInEbay = (startAmount - ebayInventory) + alreadySold

                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldEbayCol +1)
                        currentCell.value = soldInEbay
                        soldEbay += soldInEbay
                    else:
                        raise KeyError

                    
                except KeyError:
                    if 'ebay' in incorrectDict.keys():
                        incorrectDict['ebay'].append(inventorySheet.cell_value(r,ebayCol))
                    else:
                            incorrectDict['ebay'] = [inventorySheet.cell_value(r,ebayCol)]

            if parOrChild != "Child" and parOrChild != "Parent":
                #This would be the last steps of a parent before a new item with no parent
                if children != 0:
                    itemName = inventorySheet.cell_value(parentColNum , posCol)
                    startingQty = inventorySheet.cell_value(parentColNum , startQty)
                    
                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2 ,column = totalCol +1)
                    currentCell.value = totalSold
                    
                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2  ,column = soldStoreCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldStoreCol)
                    currentCell.value = soldBefore + soldPos

                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2  ,column = soldAmaCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldAmaCol)
                    currentCell.value = soldBefore + soldAma

                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2  ,column = soldEbayCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldEbayCol)
                    currentCell.value = soldBefore + soldEbay

                    currentCell = wSheet.cell(row = (parentColNum - invInfoStart + 2) ,column = finalQtyCol + 1)
                    currentCell.value = startingQty - (soldPos + soldAma)

                    soldHis = 0
                    soldPos = 0
                    soldAma = 0
                    soldEbay = 0
                    children = 0
                

    ## saves the error log
                    
    errorLog = "MasterFileNamesErrorLog.txt"
    with open(errorLog, 'w') as errors:
        for k,v in incorrectDict.items():
            errors.write("Wrong Id in the  store " + k + "\n")
            for item in v:
                errors.write(item + "\n")
                
    os.rename(os.path.join(os.getcwd(),errorLog),os.path.join(os.getcwd(), "errors", errorLog))
    newWorkbook.save("Inventory.xls")
    print("Saved!")
                    
##                    try:
##                        nameCol = 
##                        searsName = posInv[inventorySheet.cell_value(r,columnNames["Sears Name"])]
##                    except KeyError:
##                        pass
                        
                    
                    
                


##                    
##                if c == columnNames["POS Item Name"]:
##                    try:
##                        avaliablePos = posInv[inventorySheet.cell_value(r,c)]
##                        currentCell.value = inventorySheet.cell_value(r,c)
##                    except KeyError:
##                        if 'pos' in incorrectDict.keys():
##                            incorrectDict['pos'].append(inventorySheet.cell_value(r,c))
##                        else:
##                            incorrectDict['pos'] = [inventorySheet.cell_value(r,c)]
##                            
##                    
##                if c == columnNames["Amazon Sku"]:
##                    try:
##                        avaliableAma = amazonInv[inventorySheet.cell_value(r,c)]
##                        currentCell.value = inventorySheet.cell_value(r,c)
##                    except KeyError:
##                        if 'amazon' in incorrectDict.keys():
##                            incorrectDict['amazon'].append(inventorySheet.cell_value(r,c))
##                        else:
##                            incorrectDict['amazon'] = [inventorySheet.cell_value(r,c)]
##                    
##                if c == columnNames["Website Item Name"]:
##                    try:
##                        avaliableWeb = webInv[inventorySheet.cell_value(r,c)]
##                        currentCell.value = inventorySheet.cell_value(r,c)
##                    except KeyError:
##                        if 'website' in incorrectDict.keys():
##                            incorrectDict['website'].append(inventorySheet.cell_value(r,c))
##                        else:
##                            incorrectDict['website'] = [inventorySheet.cell_value(r,c)]
##                    
##                if c == columnNames["Ebay Custom Label"]:
##                    try:
##                        avaliableEbay = ebayInv[inventorySheet.cell_value(r,c)]
##                        currentCell.value = inventorySheet.cell_value(r,c)
##                    except KeyError:
##                        if 'ebay' in incorrectDict.keys():
##                            incorrectDict['ebay'].append(inventorySheet.cell_value(r,c))
##                        else:
##                            incorrectDict['ebay'] = [inventorySheet.cell_value(r,c)]
##                    
##                if c == columnNames["Short Description"]:
##                    currentCell.value = inventorySheet.cell_value(r ,c)
##                    
##                if c == columnNames["Starting Quantity"]:
##                    startAmount = inventorySheet.cell_value(r ,c)
##                    currentCell.value = startAmount
##
##                if c == columnNames["Sold in Store"]:
##                    soldPos = startAmount - avaliablePos
##                    if soldPos == startAmount:
##                        currentCell.value = 0
##                        soldPos = 0
##                    if soldPos != startAmount:
##                        if inventorySheet.cell_value(r ,c) == 0:
##                            currentCell.value = soldPos
##                        if inventorySheet.cell_value(r ,c) != 0:
##                            soldPos = inventorySheet.cell_value(r ,c) + soldPos
##                            currentCell.value = soldPos
##                    
##                if c == columnNames["Sold in Amazon"]:
##                    soldAma = startAmount - avaliableAma
##                    if soldAma == startAmount:
##                        currentCell.value = 0
##                        soldAma = 0
##                    if soldAma != startAmount:
##                        if inventorySheet.cell_value(r ,c) == 0:
##                            currentCell.value = soldAma
##                        if inventorySheet.cell_value(r ,c) != 0:
##                            soldAma = inventorySheet.cell_value(r ,c) + soldAma
##                            currentCell.value = soldAma
##                        
##                if c == columnNames["Sold in Ebay"]:
##                    soldEbay = startAmount - avaliableEbay
##                    if soldEbay == startAmount:
##                        currentCell.value = 0
##                        soldEbay = 0
##                    if soldEbay != startAmount:
##                        if inventorySheet.cell_value(r ,c) == 0:
##                            currentCell.value = soldEbay
##                        if inventorySheet.cell_value(r ,c) != 0:
##                            soldEbay = inventorySheet.cell_value(r ,c) + soldEbay
##                            currentCell.value = soldEbay
##                    
##                if c == columnNames["Sold in Website"]:
##                    soldWeb = startAmount - avaliableWeb
##                    if soldWeb == startAmount:
##                        currentCell.value = 0
##                        soldWeb = 0
##                    if soldWeb != startAmount:
##                        if inventorySheet.cell_value(r ,c) == 0:
##                            currentCell.value = soldWeb
##                        if inventorySheet.cell_value(r ,c) != 0:
##                            soldWeb = inventorySheet.cell_value(r ,c) + soldWeb
##                            currentCell.value = soldWeb
##                        
##                
##                if c == columnNames["Total Sold"]: 
##                    totalSold = soldPos + soldAma + soldEbay + soldWeb
##                    if inventorySheet.cell_value(r ,c) == 0:
##                        currentCell.value = totalSold 
##                    if inventorySheet.cell_value(r ,c) != 0:
##                        currentCell.value = inventorySheet.cell_value(r ,c) + totalSold
##                        
##                        
##                if c == columnNames["Final Amount Avaliable"]:
##                    currentCell.value = startAmount - totalSold
##
##                
##
##
##    ## saves the error log
##                    
##    errorLog = "MasterFileNamesErrorLog.txt"
##    with open(errorLog, 'w') as errors:
##        for k,v in incorrectDict.items():
##            errors.write("Wrong Id in the  store " + k + "\n")
##            for item in v:
##                errors.write(item + "\n")
##                
##    os.rename(os.path.join(os.getcwd(),errorLog),os.path.join(os.getcwd(), "errors", errorLog))
##    newWorkbook.save("Inventory.xls")
##    print("Saved!")
##
##################################################################################
##### HERE IS WHERE IT SHOULD CREATE THE TEMPLATES TO REUPLOAD###
##    
##               
##    amaInvUpdateWb = Workbook()
##    amaSheet = amaInvUpdateWb.active
##
##    ebayInvUpdateWb = Workbook()
##    ebaySheet = ebayInvUpdateWb.active
##
##    webInvUpdateWb = Workbook()
##    webSheet = webInvUpdateWb.active
##
##    posInvUpdateWb = Workbook()
##    posSheet = posInvUpdateWb.active
##
##
##################################################################################
##### HERE IS WHERE IT SHOULD repopulate THE TEMPLATES TO REUPLOAD###
##
##    
##    newPosInv = OrderedDict()
##    newAmaInv = OrderedDict()
##    newEbayInv = OrderedDict()
##    newWebInv = OrderedDict()
##    
##    productNotOnMaster = OrderedDict()
##        
##    newInvFileLocal = os.path.join(os.getcwd(),"Inventory.xls")
##    newInvSheetNum = 0
##    newInvStart = 1
##    
##
##    
##    #This part goes through the newly populated template and creates dictionaries that
##    #Will be used to create the new inventory for the new templates
##    with open_workbook(newInvFileLocal) as openInventory:
##        inventorySheet = openInventory.sheet_by_index(newInvSheetNum)
##        rows = inventorySheet.nrows
##        cols = inventorySheet.ncols
##        for r in range(newInvStart,rows):
##            posName = ""
##            amazonName = ""
##            ebayName = ""
##            websiteName = ""
##            
##            for c in range(cols):
##                currentCell = wSheet.cell(row = r + 1 ,column = c+1)
##                #print(inventorySheet.cell_value(r,c))
##
##
##                if c == columnNames["POS Item Name"]:
##                    posName = inventorySheet.cell_value(r ,c)
##                    
##                if c == columnNames["Amazon Sku"]:
##                    amazonName = inventorySheet.cell_value(r ,c)
##                    
##                if c == columnNames["Ebay Custom Label"]:
##                    ebayName = inventorySheet.cell_value(r ,c)
##                    
##                if c == columnNames["Website Item Name"]:
##                    websiteName = inventorySheet.cell_value(r ,c)
##                    
##                if c == columnNames["Final Amount Avaliable"]:
##                    finalAmount = inventorySheet.cell_value(r ,c)
##                    newPosInv[posName] = finalAmount
##                    newAmaInv[amazonName] = finalAmount
##                    newEbayInv[ebayName] = finalAmount
##                    newWebInv[websiteName] = finalAmount
##
##
####################################################################################
#####Repupulates the new amazon file with the new inventory###
##
##    ###MAKE A DICT OF NAMES THAT EXIST ON THE FILES BUT NOT ON THE MAIN FILE
##                    
##    amaCount = 1
##    amaReuploadCols = OrderedDict()
##    for k in amazonColNames.keys():
##        if k.lower() != "asin":
##            currCell = amaSheet.cell(row = 1, column = amaCount)
##            currCell.value = k.lower()
##            amaReuploadCols[k] = amaCount
##            amaCount += 1
##
##
##    prow = 1
##    for r in range(1,amaRows):
##        quantity = 0
##        for c in range(amaCols):
##            cellValue = amazonSheet.cell_value(r,c)
##            try:
##                if c == amazonColNames['sku']:
##                    currCell = amaSheet.cell(row = prow + 1, column = amaReuploadCols['sku']) 
##                    quantity = newAmaInv[cellValue]
##                    currCell.value = cellValue
##                if c == amazonColNames['price']:
##                    currCell = amaSheet.cell(row = prow + 1, column = amaReuploadCols['price'])
##                    currCell.value = cellValue
##                if c == amazonColNames['quantity']:
##                    currCell = amaSheet.cell(row = prow + 1, column = amaReuploadCols['quantity'])
##                    currCell.value = quantity
##            except KeyError:
##                prow -= 1
##                ## Create files that will tell us what product ids are on the file that are not on the main file
##                print("Caught " , cellValue)
##                if 'amazon' in productNotOnMaster.keys():
##                    productNotOnMaster['amazon'].append(cellValue)
##                else:
##                    productNotOnMaster['amazon'] = [cellValue]
##                
##        prow += 1
##                
####################################################################################
#####Repupulates the new ebay file with the new inventory##l
##    for k,v in ebayColNames.items():
##        currCell = ebaySheet.cell(row = 1, column = v + 1)
##        currCell.value = k
##
##
##
##
####################################################################################
#####Repupulates the new website file with the new inventory###
##    for k,v in webColNames.items():
##        currCell = webSheet.cell(row = 1, column = v + 1)
##        currCell.value = k
##
##    prow = 1 
##    for r in range(1,webRows):
##        quantity = 0
##        for c in range(webCols):
##            cellValue = websiteSheet.cell_value(r,c)
##            currCell = webSheet.cell(row = prow + 1, column = c + 1)
##            try:
##                if c == webColNames['Product ID']:
##                    quantity = newWebInv[cellValue]
##                    currCell.value = cellValue
##                    
##                if c == webColNames['Stock']:
##                    currCell.value = quantity
##                    
##                if c != webColNames['Stock'] and c != webColNames['Product ID']:
##                    currCell.value = cellValue
##            except KeyError:
##                prow -= 1
##                print(cellValue)
##                if 'website' in productNotOnMaster.keys():
##                    productNotOnMaster['website'].append(cellValue)
##                else:
##                    productNotOnMaster['website'] = [cellValue]
##                
##        prow += 1
##                    
##                    
##
####################################################################################
#####Repupulates the new pos file with the new inventory###
##        
##    for k,v in posColNames.items():
##        currCell = posSheet.cell(row = 1, column = v + 1)
##        currCell.value = k
##    
##    prow = 1
##    for r in range(1,posRows):
##        try:
##            itemName = posSysSheet.cell_value(r,posColNames['Item Name'])
##            quantity = newPosInv[itemName]
##            for c in range(posCols):
##                cellValue = posSysSheet.cell_value(r,c)
##                currCell = posSheet.cell(row = prow + 1, column = c + 1)
##
##                ##if a product is on the pos sys but not in the master file it will
##                ##write the first column number, change that
##
##                if c == posColNames['Qty 1']:
##                    currCell.value = quantity
##
##                if c != posColNames['Qty 1']:
##                    currCell.value = cellValue
##        except KeyError:
##            prow -= 1
##            print(itemName)
##            if 'pos' in productNotOnMaster.keys():
##                productNotOnMaster['pos'].append(itemName)
##            else:
##                productNotOnMaster['pos'] = [itemName]
##                
##        prow += 1
##                
##            #print(cellValue)
##        
##
##
##    
##    
##                    
##
##
##    missingLog = "missingInMaster.txt"
##
##    with open(missingLog, 'w') as missing:
##        for k,v in productNotOnMaster.items():
##            missing.write("Missinf Id in the  store " + k + "\n")
##            for item in v:
##                missing.write(item + "\n")
##    os.rename(os.path.join(os.getcwd(),missingLog),os.path.join(os.getcwd(), "errors", missingLog))
##
##
####################################################################################
#####HERE IS THE NAMING OF THE FILES###
##
##
##    print(productNotOnMaster)
##    amaUpdateFileName = "AmazonInvUpdate.xls"
##    ebayUpdateFileName = "EbayInvUpdate.xls"
##    webUpdateFileName = "WebInvUpdate.xls"
##    posUpdateFileName = "POSInvUpdate.xls"
##    
##    amaInvUpdateWb.save(amaUpdateFileName)
##    ebayInvUpdateWb.save(ebayUpdateFileName)
##    webInvUpdateWb.save(webUpdateFileName)
##    posInvUpdateWb.save(posUpdateFileName)
##
##
####################################################################################
#####HERE IS WHERE THE FILES ARE MOVED###
##    os.rename(os.path.join(os.getcwd(),amaUpdateFileName),os.path.join(os.getcwd(), "Updates", amaUpdateFileName))
##    os.rename(os.path.join(os.getcwd(),ebayUpdateFileName),os.path.join(os.getcwd(), "Updates", ebayUpdateFileName))
##    os.rename(os.path.join(os.getcwd(),webUpdateFileName),os.path.join(os.getcwd(), "Updates", webUpdateFileName))
##    os.rename(os.path.join(os.getcwd(),posUpdateFileName),os.path.join(os.getcwd(), "Updates", posUpdateFileName))
##    
##
##    print("Saved updates!")
##
##    
##
##
##    
##    
##
