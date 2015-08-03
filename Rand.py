            parOrChild = inventorySheet.cell_value(r , columnNames["Parents/Child/None"])
            currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = columnNames["Parents/Child/None"] +1)
            currentCell.value = parOrChild
                    
            if parOrChild == "Parent":
                

                
                if children != 0:
                    qty = columnNames["Qty 1"]
                    total = columnNames["Total Sold"]
                    itemName = inventorySheet.cell_value(parentColNum , columnNames["POS Item Name"])
                    print(itemName)
                    startingQty = inventorySheet.cell_value(parentColNum , columnNames["Starting Quantity"])

                    totalSold = t + p + a + w + e + s
                    print(totalSold, "SOLD")
                    print(t, "t",p,a,w,e,s,)
                    
                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2 ,column = total +1)
                    currentCell.value = totalSold

                    currentCell = wSheet.cell(row = (parentColNum - invInfoStart + 2) ,column = qty +1)
                    currentCell.value = startingQty - (p + a + w + e + s)

                    soldHis = 0
                    children = 0
                
                if children == 0:
                    startingQty = inventorySheet.cell_value(r , columnNames["Starting Quantity"])
                    itemName = inventorySheet.cell_value(parentColNum , columnNames["POS Item Name"])
                    p = inventorySheet.cell_value(r , columnNames["Sold in Store"])
                    a = inventorySheet.cell_value(r , columnNames["Sold in Amazon"])
                    w = inventorySheet.cell_value(r , columnNames["Sold in Website"])
                    e = inventorySheet.cell_value(r , columnNames["Sold in Ebay "])
                    s = inventorySheet.cell_value(r , columnNames["Sold in Sears"])
                    
                    parentDict[itemName] = startingQty - (p+a+w+e+s)
                    parentColNum = r
                    soldCol = columnNames[""]
                    parentSold = inventorySheet.cell_value(parentColNum , soldCol)
                    soldHis = parentSold
                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = columnNames["POS Item Name"] +1)
                    currentCell.value = itemName
                    

                
            if parOrChild == "Child":
                children += 1
                ## Calculates the inv for the pos system
                try:
                    nameCol = columnNames["POS Item Name"]
                    item = inventorySheet.cell_value(r,nameCol)
                    parent = inventorySheet.cell_value(r,columnNames["Website Item Name"])
                    posInventory = posInv[item]
                    
                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = nameCol +1)
                    currentCell.value = item
                    
                    soldCol = columnNames["Sold in Store"]
                    startQty = inventorySheet.cell_value(r , columnNames["Starting Quantity"])
                    alreadySold = inventorySheet.cell_value(r , soldCol)
                    soldInPos = (startQty - posInventory) + alreadySold

                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = soldCol +1)
                    currentCell.value = soldInPos
                    soldPos = soldInPos
                    

                    currentCell = wSheet.cell(row = parentColNum - invInfoStart + 2  ,column = soldCol +1)
                    soldHis += soldPos
                    print(soldHis, " Added up")
                    currentCell.value = soldHis 

                except KeyError:
                    if 'pos' in incorrectDict.keys():
                        incorrectDict['pos'].append(inventorySheet.cell_value(r,nameCol))
                    else:
                        incorrectDict['pos'] = [inventorySheet.cell_value(r,nameCol)]
                        
                try:
                    nameCol = columnNames["Amazon Sku"]
                    item = inventorySheet.cell_value(r,nameCol)                    
                    parent = inventorySheet.cell_value(r,columnNames["Website Item Name"])
                    amazonInventory = amazonInv[item]
                    
                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = nameCol +1)
                    currentCell.value = item

                    
                except KeyError:
                    if 'amazon' in incorrectDict.keys():
                        incorrectDict['amazon'].append(inventorySheet.cell_value(r,nameCol))
                    else:
                        incorrectDict['amazon'] = [inventorySheet.cell_value(r,nameCol)]
                        
                try:
                    nameCol = columnNames["Website Item Name"]
                    item = inventorySheet.cell_value(r,nameCol)
                    parent = inventorySheet.cell_value(r,columnNames["Website Item Name"])
                    
                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = nameCol +1)
                    websiteInventory = webInv[item]
                    currentCell.value = item
                except KeyError:
                    if 'website' in incorrectDict.keys():
                        incorrectDict['website'].append(inventorySheet.cell_value(r,nameCol))
                    else:
                        incorrectDict['website'] = [inventorySheet.cell_value(r,nameCol)]
                
                try:
                    nameCol = columnNames["Ebay Custom Label"]
                    item = inventorySheet.cell_value(r,nameCol)
                    parent = inventorySheet.cell_value(r,columnNames["Website Item Name"])
                    
                    currentCell = wSheet.cell(row = r - invInfoStart + 2 ,column = nameCol +1)
                    ebayInventory = ebayInv[item]
                    currentCell.value = item
                except KeyError:
                    if 'ebay' in incorrectDict.keys():
                        incorrectDict['ebay'].append(inventorySheet.cell_value(r,nameCol))
                    else:
                            incorrectDict['ebay'] = [inventorySheet.cell_value(r,nameCol)]

