import openpyxl as opyxl
import re

fileName = 'BetaCalculator_Update 3.0.3.xlsx'
testFile = 'testexcel.xlsx'
defSheetName = 'Data Aggregation'
sheetName = 'GreenUser'
sheetNames = ['Start', 'Household Input', 'Style Change', 'GreenUser', 'Data Aggregation', 'Report Preferences',
              'Food Footprint Report', 'Diet Calculator', 'Database Food', 'total footprint report',
              'export alternative lifestyles', 'Diet Input', 'household calculator', 'back end data',
              'GreenStyleUser', 'explore alternative diets', 'Sub_Data', 'Transport Info', 'Dropdown',
              'Electricity mix footprint', 'old data sheet', 'useful links']
WB = opyxl.load_workbook(fileName)
WS = WB.get_sheet_by_name(defSheetName)

val = WS['G2'].value
WS['G2'] = 61
WB.save(fileName)

WBTest = opyxl.load_workbook(fileName)
WSTest = WBTest.get_sheet_by_name(defSheetName)
val = WS['B34'].value

if(type(val) is str):
    valList = re.findall(r"[\w']+", val) ## Removes all sort of expressions.
    print(valList)

    sheetValListIndices = []
    cellValListIndices = []
    index = 0
    for eachVal in valList:
        sheetFlag = False
        for sheet in sheetNames:
            if(eachVal == sheet):
                print("Its a sheet")
                sheetFlag = True
                sheetValListIndices.append(index)
        if(sheetFlag is False):
            print("Its a cell")
            cellValListIndices.append(index)
        index+= 1
    print(sheetValListIndices)
    print(cellValListIndices)
    ## Assumption: Each cell contains only one refering sheet and the corresponding cell.


###### Feeding GU relations, Frequency relations formula wise to python. ###########
## Gulighting -- 1. (Green user,M33) -->  2. (52.14 * Green user, H33) --> 3. (Green user, J33 * Backend data, B63)
## Todo: Find backend data b63
    ## It has vlookup and plenty of other sheets.












