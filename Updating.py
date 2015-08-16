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
    aRowNameStart = 1
    aInfoStart = 2

    ebayInfoFileName = "EbayShort.xls"
    ebayInfoType = "Inventory"
    ebayInfoLocal = os.path.join(os.getcwd(), "ToUpdate", "Shortened", ebayInfoFileName)
    ebayInfoSheet = 1
    eRowNameStart = 1
    eInfoStart = 2

    websiteInfoFileName = "WebsiteShort.xls"
    websiteInfoType = "Inventory"
    websiteInfoLocal = os.path.join(os.getcwd(), "ToUpdate", "Shortened", websiteInfoFileName)
    websiteInfoSheet = 1
    wRowNameStart = 1
    wInfoStart = 2

    posInfoFile = "POS.xls"
    posInfoType = "Inventory"
    posInfoLocal = os.path.join(os.getcwd(), "ToUpdate", "Shortened", posInfoFile)
    posInfoSheet = 1
    pRowNameStart = 1
    pInfoStart = 2

    searsInfoFileName = "SearsShort.xls"
    searInfoLocal = os.path.join(os.getcwd(), "ToUpdate", "Shortened", searsInfoFileName)
    searsSheetNum = 3
    sRowNameStart = 1
    sInfoStart = 2
    
    inventoryFileName = "AllInventory.xls"
    inventoryInfoLocal = os.path.join(os.getcwd(), "MainFiles", "InventoryFile", inventoryFileName)
    invSheetNum = 4
    invRowNameStart = 10
    invInfoStart = 11



    

    
    #Here is where i change the indexes to match
    amazonInfoSheet = amazonInfoSheet - 1
    aRowNameStart = aRowNameStart - 1
    aInfoStart = aInfoStart - 1

    ebayInfoSheet = ebayInfoSheet - 1
    eRowNameStart = eRowNameStart - 1
    eInfoStart = eInfoStart - 1
    
    websiteInfoSheet = websiteInfoSheet - 1
    wRowNameStart = wRowNameStart - 1
    wInfoStart = wInfoStart - 1

    posInfoSheet = posInfoSheet - 1
    pRowNameStart = pRowNameStart - 1
    pInfoStart = pInfoStart - 1

    invSheetNum = invSheetNum - 1
    invRowNameStart = invRowNameStart - 1
    invInfoStart = invInfoStart - 1

    searsSheetNum = searsSheetNum - 1
    sRowNameStart = sRowNameStart - 1
    sInfoStart = sInfoStart - 1
    
    
    #end of index change



    aInfoCols = ["sku","quantity"]
    amazonColNames = columnNames(amazonInfoLocal, aRowNameStart, amazonInfoSheet)
    amazonInfoOpen = open_workbook(amazonInfoLocal)
    amazonSheet = amazonInfoOpen.sheet_by_index(amazonInfoSheet)
    amaRows = amazonSheet.nrows
    amaCols = amazonSheet.ncols
    amazonInv = inventoryDict(amazonSheet, amazonColNames, aInfoCols, aInfoStart)
##    print("Amazon INV")
##    print(amazonInv)
##    print()


    eInfoCols = ["Custom Label", "Quantity Available"]
    ebayColNames = columnNames(ebayInfoLocal, eRowNameStart, ebayInfoSheet)
    ebayInfoOpen = open_workbook(ebayInfoLocal)
    ebaySheet = ebayInfoOpen.sheet_by_index(ebayInfoSheet)
    ebayRows = ebaySheet.nrows
    ebayCols = ebaySheet.ncols
    ebayInv = inventoryDict(ebaySheet, ebayColNames, eInfoCols, eInfoStart)
##    print("Ebay INV")
##    print(ebayInv)
##    print()

    

    webInfoCols = ["Product ID", "Stock"]
    webColNames = columnNames(websiteInfoLocal, wRowNameStart, websiteInfoSheet)
    websiteInfoOpen = open_workbook(websiteInfoLocal)
    websiteSheet = websiteInfoOpen.sheet_by_index(websiteInfoSheet)
    webRows = websiteSheet.nrows
    webCols = websiteSheet.ncols
    webInv = inventoryDict(websiteSheet, webColNames, webInfoCols, wInfoStart)
##    print("Website INV")
##    print(webInv)  
##    print()

    posInfoCols = ["Item Name", "Qty 1"]
    posColNames = columnNames(posInfoLocal, pRowNameStart,posInfoSheet)
    posInfoOpen = open_workbook(posInfoLocal)
    posSysSheet = posInfoOpen.sheet_by_index(posInfoSheet)
    posRows = posSysSheet.nrows
    posCols = posSysSheet.ncols
    posInv = inventoryDict(posSysSheet, posColNames, posInfoCols, pInfoStart)
##    print("POS Inv")
##    print(posInv)
##    print()


    sInfoCols = ["Item Id", "Existing Available Quantity"]
    searsColNames = columnNames(searInfoLocal, sRowNameStart , searsSheetNum)
    searsInfoOpen = open_workbook(searInfoLocal)
    searsInfoSheet = searsInfoOpen.sheet_by_index(searsSheetNum)
    sRows = searsInfoSheet.nrows
    sCols = searsInfoSheet.ncols

    searsInv = inventoryDict(searsInfoSheet, searsColNames, sInfoCols , sInfoStart)

    #print("Sears Col Names", searsColNames)

    
    newWorkbook = Workbook()
    wSheet = newWorkbook.active

################################################################################
####HERE IS WHERE IT GETS ALL THE INVENTORY AND CREATES A NEW TEMPLATE####


    incorrectDict = defaultdict()
    
    with open_workbook(inventoryInfoLocal) as openInventory:
        inventorySheet = openInventory.sheet_by_index(invSheetNum)
        rows = inventorySheet.nrows
        cols = inventorySheet.ncols
        
        columnNames = columnNames(inventoryInfoLocal, invRowNameStart, invSheetNum)
        revColNames = invEnumerated(columnNames)

        ### MAKE A DICT OF NAMES THAT DON'T EXIST IN THE STORE FILES BUT WAS PUT ON THE MAIN FILE
        
        for k,v in columnNames.items():
            
            currentCell = wSheet.cell(row = 1, column = v + 1)
            currentCell.value = k






        modifyItems = []
        children = 0
        parentRowNum = 0

        soldPos = 0
        soldAma = 0
        soldWeb = 0
        soldEbay = 0
        soldSears = 0
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
            
            startAmount = 0
            avaliablePos = 0
            avaliableWeb = 0
            avaliableAma = 0
            avaliableEbay = 0
            avaliableSears = 0
            totalSold = 0
                    
            if parOrChild == "Parent":
                
                if children == 0:
                    parentRowNum = r
                    #Calculates the amount already sold of that item
                    startingQty = inventorySheet.cell_value(r , startQty)
                    itemName = inventorySheet.cell_value(parentRowNum , posCol)
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
                        ###Change so that a 0 is placed if an item is not there
                        ###ATM it leaves a blank space
                        soldInPos = 0
                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldStoreCol +1)
                        currentCell.value = soldInPos
                        soldPos += soldInPos

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
                        ###Change so that a 0 is placedif an item is not there
                        ###ATM it leaves a blank space
                        soldInAma = 0
                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldAmaCol +1)
                        currentCell.value = soldInAma
                        soldAma += soldInAma
                    
                except KeyError:
                    if 'amazon' in incorrectDict.keys():
                        incorrectDict['amazon'].append(inventorySheet.cell_value(r,amazonCol))
                    else:
                        incorrectDict['amazon'] = [inventorySheet.cell_value(r,amazonCol)]

                
                ## calculates the children inv for Website
                ##This one is special
                try:
                    item = inventorySheet.cell_value(parentRowNum,posCol)
                    websiteInventory = webInv[item]
                    
                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = websiteCol +1)
                    currentCell.value = item
                    
                    if websiteInventory == inventorySheet.cell_value(parentRowNum,finalQtyCol):
                        soldInWeb = inventorySheet.cell_value(r,soldWebCol)
                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldWebCol +1)
                        currentCell.value = soldInWeb

                    if websiteInventory != inventorySheet.cell_value(parentRowNum,finalQtyCol):
                        startAmount = inventorySheet.cell_value(parentRowNum , startQty)
                        alreadySold = inventorySheet.cell_value(parentRowNum , soldWebCol)
                        soldInWeb = inventorySheet.cell_value(r , soldWebCol)
                        if item not in modifyItems:
                            modifyItems.append(item)

                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldWebCol +1)
                        currentCell.value = inventorySheet.cell_value(r,soldWebCol)
                        soldWeb = (startAmount - websiteInventory) + alreadySold

                        
                    

                    
                except KeyError:
                    print("error")
                    if 'website' in incorrectDict.keys():
                        incorrectDict['website'].append(inventorySheet.cell_value(r,websiteCol))
                    else:
                        incorrectDict['website'] = [inventorySheet.cell_value(r,websiteCol)]

                        
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
                        #puts a 0 if there is no name for this item in this store in the main file
                        soldInEbay = 0
                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldEbayCol +1)
                        currentCell.value = soldInEbay
                        soldEbay += soldInEbay
                        

                    
                except KeyError:
                    if 'ebay' in incorrectDict.keys():
                        incorrectDict['ebay'].append(inventorySheet.cell_value(r,ebayCol))
                    else:
                        incorrectDict['ebay'] = [inventorySheet.cell_value(r,ebayCol)]



                try:
                    item = inventorySheet.cell_value(r,searsCol)
                    if item not in ["", " "]:
                        searsInventory = searsInv[item]

                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = searsCol +1)
                        currentCell.value = item

                        startAmount = inventorySheet.cell_value(r , startQty)
                        alreadySold = inventorySheet.cell_value(r , soldSearsCol)
                        soldInSears = (startAmount - searsInventory) + alreadySold

                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldSearsCol +1)
                        currentCell.value = soldInSears
                        soldSears += soldInSears
                    else:
                        soldInSears = 0
                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldSearsCol +1)
                        currentCell.value = soldInEbay
                        soldEbay += soldInEbay
                        
                    searsName = posInv[inventorySheet.cell_value(r,columnNames["Sears Name"])]
                except KeyError:
                    if 'sears' in incorrectDict.keys():
                        incorrectDict['sears'].append(inventorySheet.cell_value(r,searsCol))
                    else:
                        incorrectDict['sears'] = [inventorySheet.cell_value(r,searsCol)]

                        

                #Calculates the total sold for each child
                
                currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = totalCol +1)
                soldBefore = inventorySheet.cell_value(r , totalCol)
                currentCell.value = soldBefore + (soldInPos + soldInAma + soldInEbay + soldInWeb + soldInSears)
                
                #calculates the final amount for each child
                currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = finalQtyCol +1)
                currentCell.value = startAmount - (soldInAma + soldInEbay + soldInPos + soldInSears + soldBefore)

                try:
                    relation = inventorySheet.cell_value(r + 1,relationCol)
                    if relation == "Child":
                        pass
                    if relation in ["Parent", ""," "]:
                        raise IndexError
                except IndexError:
                    itemName = inventorySheet.cell_value(parentRowNum , posCol)
                    startingQty = inventorySheet.cell_value(parentRowNum , startQty)
                    

                    
                    currentCell = wSheet.cell(row = parentRowNum - invInfoStart + 2  ,column = soldStoreCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldStoreCol)
                    currentCell.value = soldBefore + soldPos

                    currentCell = wSheet.cell(row = parentRowNum - invInfoStart + 2  ,column = soldAmaCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldAmaCol)
                    currentCell.value = soldBefore + soldAma

                    currentCell = wSheet.cell(row = parentRowNum - invInfoStart + 2  ,column = soldEbayCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldEbayCol)
                    currentCell.value = soldBefore + soldEbay

                    currentCell = wSheet.cell(row = parentRowNum - invInfoStart + 2  ,column = soldWebCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldWebCol)
                    currentCell.value = soldBefore + soldWeb

                    currentCell = wSheet.cell(row = parentRowNum - invInfoStart + 2  ,column = soldSearsCol + 1)
                    soldBefore = inventorySheet.cell_value(r , soldSearsCol)
                    currentCell.value = soldBefore + soldSears

                    currentCell = wSheet.cell(row = parentRowNum - invInfoStart + 2 ,column = totalCol +1)
                    soldBefore = inventorySheet.cell_value(parentRowNum , totalCol)
                    currentCell.value = soldBefore + (soldPos + soldAma + soldEbay + soldWeb + soldSears)

                    currentCell = wSheet.cell(row = (parentRowNum - invInfoStart + 2) ,column = finalQtyCol + 1)
                    currentCell.value = startingQty - (soldPos + soldAma + soldEbay + soldWeb + soldSears)
                    

                    soldHis = 0
                    soldPos = 0
                    soldAma = 0
                    soldEbay = 0
                    soldWeb = 0
                    soldSears = 0
                    children = 0
                    

            if parOrChild != "Child" and parOrChild != "Parent":
                #This would happen if the item isn't a parent or a child
                    
                if children == 0:
                    soldPos = 0
                    soldAma = 0
                    soldWeb = 0
                    soldEbay = 0
                    soldSears = 0
                    totalSold = 0
                    finalAmount = 0
                    for c in range(cols):
                        #print(r,c)
                        currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = c+1)
                        if c == columnNames["POS Item Name"]:
                            try:
                                avaliablePos = posInv[inventorySheet.cell_value(r,c)]
                                currentCell.value = inventorySheet.cell_value(r,c)
                            except KeyError:
                                if 'pos' in incorrectDict.keys():
                                    incorrectDict['pos'].append(inventorySheet.cell_value(r,c))
                                else:
                                    incorrectDict['pos'] = [inventorySheet.cell_value(r,c)]
                                    
                            
                        if c == columnNames["Amazon Sku"]:
                            try:
                                avaliableAma = amazonInv[inventorySheet.cell_value(r,c)]
                                currentCell.value = inventorySheet.cell_value(r,c)
                            except KeyError:
                                if 'amazon' in incorrectDict.keys():
                                    incorrectDict['amazon'].append(inventorySheet.cell_value(r,c))
                                else:
                                    incorrectDict['amazon'] = [inventorySheet.cell_value(r,c)]
                            
                        if c == columnNames["Website Item Name"]:
                            try:
                                avaliableWeb = webInv[inventorySheet.cell_value(r,c)]
                                currentCell.value = inventorySheet.cell_value(r,c)
                            except KeyError:
                                if 'website' in incorrectDict.keys():
                                    incorrectDict['website'].append(inventorySheet.cell_value(r,c))
                                else:
                                    incorrectDict['website'] = [inventorySheet.cell_value(r,c)]
                            
                        if c == columnNames["Ebay Custom Label"]:
                            try:
                                avaliableEbay = ebayInv[inventorySheet.cell_value(r,c)]
                                currentCell.value = inventorySheet.cell_value(r,c)
                            except KeyError:
                                if 'ebay' in incorrectDict.keys():
                                    incorrectDict['ebay'].append(inventorySheet.cell_value(r,c))
                                else:
                                    incorrectDict['ebay'] = [inventorySheet.cell_value(r,c)]

                        if c == columnNames["Sears Name"]:
                            try:
                                avaliableSears = searsInv[inventorySheet.cell_value(r,c)]
                                currentCell.value = inventorySheet.cell_value(r,c)
                            except KeyError:
                                if 'sears' in incorrectDict.keys():
                                    incorrectDict['sears'].append(inventorySheet.cell_value(r,c))
                                else:
                                    incorrectDict['sears'] = [inventorySheet.cell_value(r,c)]

                                    
                            
                        if c == columnNames["Short Description"]:
                            currentCell.value = inventorySheet.cell_value(r ,c)
                            
                        if c == columnNames["Starting Quantity"]:
                            startAmount = inventorySheet.cell_value(r ,c)
                            currentCell.value = startAmount

                        if c == columnNames["Sold in Store"]:
                            soldPos = startAmount - avaliablePos
                            if soldPos == startAmount:
                                currentCell.value = 0
                                soldPos = 0
                            if soldPos != startAmount:
                                if inventorySheet.cell_value(r ,c) == 0:
                                    currentCell.value = soldPos
                                if inventorySheet.cell_value(r ,c) != 0:
                                    soldPos = inventorySheet.cell_value(r ,c) + soldPos
                                    currentCell.value = soldPos
                            
                        if c == columnNames["Sold in Amazon"]:
                            soldAma = startAmount - avaliableAma
                            if soldAma == startAmount:
                                currentCell.value = 0
                                soldAma = 0
                            if soldAma != startAmount:
                                if inventorySheet.cell_value(r ,c) == 0:
                                    currentCell.value = soldAma
                                if inventorySheet.cell_value(r ,c) != 0:
                                    soldAma = inventorySheet.cell_value(r ,c) + soldAma
                                    currentCell.value = soldAma
                                
                        if c == columnNames["Sold in Website"]:
                            soldWeb = startAmount - avaliableWeb
                            if soldWeb == startAmount:
                                currentCell.value = 0
                                soldWeb = 0
                            if soldWeb != startAmount:
                                if inventorySheet.cell_value(r ,c) == 0:
                                    currentCell.value = soldWeb
                                if inventorySheet.cell_value(r ,c) != 0:
                                    soldWeb = inventorySheet.cell_value(r ,c) + soldWeb
                                    currentCell.value = soldWeb
                                    
                        if c == columnNames["Sold in Ebay"]:
                            soldEbay = startAmount - avaliableEbay
                            if soldEbay == startAmount:
                                currentCell.value = 0
                                soldEbay = 0
                            if soldEbay != startAmount:
                                if inventorySheet.cell_value(r ,c) == 0:
                                    currentCell.value = soldEbay
                                if inventorySheet.cell_value(r ,c) != 0:
                                    soldEbay = inventorySheet.cell_value(r ,c) + soldEbay
                                    currentCell.value = soldEbay

                    
                        if c == columnNames["Sold in Sears"]:
                            soldSears = startAmount - avaliableSears
                            if soldSears == startAmount:
                                currentCell.value = 0
                                soldSears = 0
                            if soldSears != startAmount:
                                if inventorySheet.cell_value(r ,c) == 0:
                                    currentCell.value = soldSears
                                if inventorySheet.cell_value(r ,c) != 0:
                                    soldSears = inventorySheet.cell_value(r ,c) + soldSears
                                    currentCell.value = soldSears

                           

                                
                        
                        if c == columnNames["Total Sold"]: 
                            totalSold = soldPos + soldAma + soldEbay + soldWeb
                            if inventorySheet.cell_value(r ,c) == 0:
                                currentCell.value = totalSold 
                            if inventorySheet.cell_value(r ,c) != 0:
                                currentCell.value = inventorySheet.cell_value(r ,c) + totalSold
                                
                                
                        if c == columnNames["Qty 1"]:
                            currentCell.value = startAmount - totalSold
                    soldPos = 0
                    soldAma = 0
                    soldWeb = 0
                    soldEbay = 0
                    soldSears = 0
                    totalSold = 0
                    finalAmount = 0


                    

    ## saves the error log


    modify = "checkAndModify.txt"
    with open(modify, 'w') as mod:
        if modifyItems != []:
            mod.write("Some Rings were sold in the website.\nAdd the amount of each child sold:" + "\n")
            for i in modifyItems:
                mod.write(i + "\n")
                
 
                
    os.rename(os.path.join(os.getcwd(),modify),os.path.join(os.getcwd(), "errors", modify))

    
    errorLog = "MasterFileNamesErrorLog.txt"
    with open(errorLog, 'w') as errors:
        for k,v in incorrectDict.items():
            errors.write("Wrong Id in the  store " + k + "\n")
            for item in v:
                errors.write(item + "\n")
                
    os.rename(os.path.join(os.getcwd(),errorLog),os.path.join(os.getcwd(), "errors", errorLog))
    newWorkbook.save("Inventory.xls")
    print("Saved!")
                    

                    
                


################################################################################
### HERE IS WHERE IT SHOULD CREATE THE TEMPLATES TO REUPLOAD###
    
               
    amaInvUpdateWb = Workbook()
    amaSheet = amaInvUpdateWb.active

    ebayInvUpdateWb = Workbook()
    ebayNewSheet = ebayInvUpdateWb.active

    webInvUpdateWb = Workbook()
    webSheet = webInvUpdateWb.active

    posInvUpdateWb = Workbook()
    posSheet = posInvUpdateWb.active

    searsInvUpdateWb = Workbook()
    searsSheet = searsInvUpdateWb.active


################################################################################
### HERE IS WHERE IT SHOULD repopulate THE TEMPLATES TO REUPLOAD###

    
    newPosInv = OrderedDict()
    newAmaInv = OrderedDict()
    newEbayInv = OrderedDict()
    newWebInv = OrderedDict()
    newSearsInv = OrderedDict()
    
    productNotOnMaster = OrderedDict()
        
    newInvFileLocal = os.path.join(os.getcwd(),"Inventory.xls")
    newInvSheetNum = 0
    newInvStart = 1

    #This part goes through the newly populated template and creates dictionaries that
    #Will be used to create the new inventory for the new templates
    with open_workbook(newInvFileLocal) as openInventory:
        inventorySheet = openInventory.sheet_by_index(newInvSheetNum)
        rows = inventorySheet.nrows
        cols = inventorySheet.ncols
        for r in range(newInvStart,rows):
            posName = ""
            amazonName = ""
            ebayName = ""
            websiteName = ""
            searsName = ""
            
            for c in range(cols):
                currentCell = wSheet.cell(row = r + 1 ,column = c+1)
                #print(inventorySheet.cell_value(r,c))

                if inventorySheet.cell_value(r, relationCol) == "Parent":
                    newWebInv[inventorySheet.cell_value(r ,c)] = inventorySheet.cell_value(r ,finalQtyCol)
                    

                if inventorySheet.cell_value(r, relationCol) != "Parent":
                    if inventorySheet.cell_value(r, relationCol) != "Child":
                        if c == columnNames["Website Item Name"]:
                            newWebInv[inventorySheet.cell_value(r ,c)] = inventorySheet.cell_value(r ,finalQtyCol)
                                      
                    if c == columnNames["POS Item Name"]:
                        posName = inventorySheet.cell_value(r ,c)
                        
                    if c == columnNames["Amazon Sku"]:
                        amazonName = inventorySheet.cell_value(r ,c)
                        
                    if c == columnNames["Ebay Custom Label"]:
                        ebayName = inventorySheet.cell_value(r ,c)

                    if c == columnNames["Sears Name"]:
                        searsName = inventorySheet.cell_value(r ,c)
                        
                        
                    if c == columnNames["Qty 1"]:
                        finalAmount = inventorySheet.cell_value(r ,c)
                        newPosInv[posName] = finalAmount
                        newAmaInv[amazonName] = finalAmount
                        newEbayInv[ebayName] = finalAmount
                        newSearsInv[searsName] = finalAmount



##################################################################################
###Repupulates the new amazon file with the new inventory###

    ###MAKE A DICT OF NAMES THAT EXIST ON THE FILES BUT NOT ON THE MAIN FILE
                    
    amaCount = 1
    amaReuploadCols = OrderedDict()
    for k in amazonColNames.keys():
        if k.lower() != "asin":
            currCell = amaSheet.cell(row = 1, column = amaCount)
            currCell.value = k.lower()
            amaReuploadCols[k] = amaCount
            amaCount += 1


    prow = 1
    for r in range(1,amaRows):
        quantity = 0
        for c in range(amaCols):
            cellValue = amazonSheet.cell_value(r,c)
            try:
                if c == amazonColNames['sku']:
                    currCell = amaSheet.cell(row = prow + 1, column = amaReuploadCols['sku']) 
                    quantity = newAmaInv[cellValue]
                    currCell.value = cellValue
                if c == amazonColNames['price']:
                    currCell = amaSheet.cell(row = prow + 1, column = amaReuploadCols['price'])
                    currCell.value = cellValue
                if c == amazonColNames['quantity']:
                    currCell = amaSheet.cell(row = prow + 1, column = amaReuploadCols['quantity'])
                    currCell.value = quantity
            except KeyError:
                prow -= 1
                ## Create files that will tell us what product ids are on the file that are not on the main file
                print(cellValue, "Missing in Amazon Col")
                if 'amazon' in productNotOnMaster.keys():
                    productNotOnMaster['amazon'].append(cellValue)
                else:
                    productNotOnMaster['amazon'] = [cellValue]
                
        prow += 1
                
##################################################################################
###Repupulates the new ebay file with the new inventory##l
    for k,v in ebayColNames.items():
        currCell = ebayNewSheet.cell(row = 1, column = v + 1)
        currCell.value = k
        
    prow = 1 
    for r in range(1,ebayRows):
        quantity = 0
        for c in range(ebayCols):
            #change
            cellValue = ebaySheet.cell_value(r,c)
            currCell = ebayNewSheet.cell(row = prow + 1, column = c + 1)
            try:
                if c == ebayColNames["Custom Label"]:
                    quantity = newEbayInv[cellValue]
                    currCell.value = cellValue
                    
                if c == ebayColNames["Quantity Available"]:
                    currCell.value = quantity
                    
                if c != ebayColNames["Custom Label"] and c != ebayColNames["Quantity Available"]:
                    currCell.value = cellValue
                    
            except KeyError:
                prow -= 1
                print(cellValue, "Missing in Ebay Col")
                if 'ebay' in productNotOnMaster.keys():
                    productNotOnMaster['ebay'].append(cellValue)
                else:
                    productNotOnMaster['ebay'] = [cellValue]
                    
        prow += 1




##################################################################################
###Repupulates the new website file with the new inventory###
    for k,v in webColNames.items():
        currCell = webSheet.cell(row = 1, column = v + 1)
        currCell.value = k

    prow = 1 
    for r in range(1,webRows):
        quantity = 0
        for c in range(webCols):
            cellValue = websiteSheet.cell_value(r,c)
            currCell = webSheet.cell(row = prow + 1, column = c + 1)
            try:
                if c == webColNames['Product ID']:
                    quantity = newWebInv[cellValue]
                    currCell.value = cellValue
                    
                if c == webColNames['Stock']:
                    currCell.value = quantity
                    
                if c != webColNames['Stock'] and c != webColNames['Product ID']:
                    currCell.value = cellValue
            except KeyError:
                prow -= 1
                print(cellValue, "Missing in Web Col")
                if 'website' in productNotOnMaster.keys():
                    productNotOnMaster['website'].append(cellValue)
                else:
                    productNotOnMaster['website'] = [cellValue]
                
        prow += 1
                    
                    

##################################################################################
###Repupulates the new pos file with the new inventory###
        
    for k,v in posColNames.items():
        currCell = posSheet.cell(row = 1, column = v + 1)
        currCell.value = k
    
    prow = 1
    for r in range(1,posRows):
        itemName = posSysSheet.cell_value(r,posColNames['Item Name'])
        try:
            quantity = newPosInv[itemName]
            for c in range(posCols):
                cellValue = posSysSheet.cell_value(r,c)
                currCell = posSheet.cell(row = prow + 1, column = c + 1)

                ##if a product is on the pos sys but not in the master file it will
                ##write the first column number, change that

                if c == posColNames['Qty 1']:
                    currCell.value = quantity

                if c != posColNames['Qty 1']:
                    currCell.value = cellValue
        except KeyError:
            prow -= 1
            print(itemName, "Missing in POS Col")
            if 'pos' in productNotOnMaster.keys():
                productNotOnMaster['pos'].append(itemName)
            else:
                productNotOnMaster['pos'] = [itemName]
                
        prow += 1
                
            #print(cellValue)
        
##################################################################################
###Repupulates the new Sears file with the new inventory###

    for k,v in searsColNames.items():
        currCell = searsSheet.cell(row = 1, column = v + 1)
        currCell.value = k

    prow = 1
    for r in range(1,sRows):
        itemName = searsInfoSheet.cell_value(r,searsColNames['Item Id'])
        try:
            quantity = newSearsInv[itemName]
            for c in range(sCols):
                cellValue = searsInfoSheet.cell_value(r,c)
                currCell = searsSheet.cell(row = prow + 1, column = c+ 1)

                if c == searsColNames['Updated Available Quantity']:
                    currCell.value = quantity

                if c != searsColNames['Updated Available Quantity']:
                    currCell.value = cellValue

        except KeyError:
            prow -= 1
            print(itemName, "Missing in Sears Col")
            if 'sears' in productNotOnMaster.keys():
                productNotOnMaster['sears'].append(itemName)
            else:
                productNotOnMaster['sears'] = [itemName]
                
        prow += 1
            
                    

                
    
    
                    


    missingLog = "missingInMaster.txt"

    with open(missingLog, 'w') as missing:
        for k,v in productNotOnMaster.items():
            missing.write("Missinf Id in the  store " + k + "\n")
            for item in v:
                missing.write(item + "\n")
    os.rename(os.path.join(os.getcwd(),missingLog),os.path.join(os.getcwd(), "errors", missingLog))


##################################################################################
###HERE IS THE NAMING OF THE FILES###


    amaUpdateFileName = "AmazonInvUpdate.xls"
    ebayUpdateFileName = "EbayInvUpdate.xls"
    webUpdateFileName = "WebInvUpdate.xls"
    posUpdateFileName = "POSInvUpdate.xls"
    searsUpdateFileName = "SearsInvUpdate.xls"
    
    amaInvUpdateWb.save(amaUpdateFileName)
    ebayInvUpdateWb.save(ebayUpdateFileName)
    webInvUpdateWb.save(webUpdateFileName)
    posInvUpdateWb.save(posUpdateFileName)
    searsInvUpdateWb.save(searsUpdateFileName)


##################################################################################
###HERE IS WHERE THE FILES ARE MOVED###
    os.rename(os.path.join(os.getcwd(),amaUpdateFileName),os.path.join(os.getcwd(), "Updates", amaUpdateFileName))
    os.rename(os.path.join(os.getcwd(),ebayUpdateFileName),os.path.join(os.getcwd(), "Updates", ebayUpdateFileName))
    os.rename(os.path.join(os.getcwd(),webUpdateFileName),os.path.join(os.getcwd(), "Updates", webUpdateFileName))
    os.rename(os.path.join(os.getcwd(),posUpdateFileName),os.path.join(os.getcwd(), "Updates", posUpdateFileName))
    os.rename(os.path.join(os.getcwd(),searsUpdateFileName),os.path.join(os.getcwd(), "Updates", searsUpdateFileName))

    print("Saved updates!")

    


    
    

