# Use xlrd for the time being. Then use pandas when you know.
import sys
from xlrd import open_workbook
def marginalErrorReading():
    for sheet in margErrTable.sheets():
        if sheet.name == 'Output':
            global margRows, margCols
            margRows = sheet.nrows
            margCols = sheet.ncols
            for col in range(sheet.ncols):
                values = [] ## Localised values for a row
                if(col != 0):
                    for row in range(sheet.nrows):
                        data = sheet.cell(row, col).value
                        if row == 0:
                            headers.append(data)
                        else:
                            values.append(data)
                    if(row > 0):
                        margErrRed.append(values)

def pickLeastValues():
    leastValues = [] # Pick least values in each params.
    for value in margErrRed:
        leastValues.append(min(value))
    return leastValues

def pickLeastParams(minValuesList):
    minValue = min(minValuesList)
    #print(minValue)
    ##Find the position for it.
    global minParam
    minParam = minValuesList.index(minValue)
    minParam = minParam + 1 #Since the index starts from 1.

    ##Find the recommendation percentage
    paramTempPos = 1
    for value in margErrRed:
        if(paramTempPos == minParam):
            minParamValuePos = value.index(minValue) + 1 ## SInce the position starts from 0
            paramTempPos = paramTempPos + 1
        else:
            paramTempPos = paramTempPos + 1
    recoPercentage = ((minParamValuePos * 10) - 10) #Since the reco percentage starts from 0
    return recoPercentage

def overWriting(colToBeAltered):
    ## Alteration takes place in the column.
    ## Step 1: Find out the position of margErrRed.

    for i in range(margRows - 1):
        margErrRed[colToBeAltered - 1][i] = sys.maxsize

margErrTable = open_workbook('Demo Mapper.xlsx')
global margErrRed
margErrRed = []
global headers
headers = []
global changePercent
changePercent = []
marginalErrorReading()
totalParameters = 4 ## Hardcode of parameters based on the no.of parameters chosen.
for i in range(totalParameters):
    minValues = pickLeastValues()
    #print(minValues)
    recoPercentage = pickLeastParams(minValues)
    parameter = headers[minParam - 1]
    print ("The recommendation percentage is ", recoPercentage, "in Parameter : ", parameter )
    if(i != totalParameters - 1): ## Since i iterates from 0.
        print("Or")
    overWriting(minParam)






