# Use xlrd for the time being. Then use pandas when you know.
from xlrd import open_workbook


margErrTable = open_workbook('LEDsafari Analysis.xlsx')
global margErrRed
margErrRed = []
global headers
headers = []


def marginalErrorReading():
    for sheet in margErrTable.sheets():
        if sheet.name == 'Sheet2':
            for col in range(sheet.ncols):
                values = [] ## Localised values for a row
                for row in range(sheet.nrows):
                    data = sheet.cell(row, col).value
                    if row == 0:
                        headers.append(data)
                    else:
                        values.append(data)
                if(row > 0):
                    margErrRed.append(values)
    print(margErrRed)

def pickLeastValues():
    leastValues = [] # Pick least values in each params.
    for value in margErrRed:
        leastValues.append(min(value))
    return leastValues

def pickLeastParams(minValuesList):
    minValue = min(minValuesList)
    print(minValue)
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
            #TODO: Over writing the values once used
            #for eachValue in value:
             #   margErrRed.
        else:
            paramTempPos = paramTempPos + 1
    recoPercentage = (minParamValuePos * 10)
    return recoPercentage

marginalErrorReading()
minValues = pickLeastValues()
print(minValues)
recoPercentage = pickLeastParams(minValues)
print ("The recommendation percentage is ", recoPercentage, "in Parameter : ", minParam )






