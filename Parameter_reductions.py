import openpyxl as opyxl

def ReadValues(cellStr, cellStart, cellEnd):
        values = []
        for i in range(cellStart, cellEnd + 1):
            cellIndex = cellStr
            cellIndex+= str(i)
            values.append(WS[cellIndex].value)
        return values
def DictMapping(keyParam, valueFreq, dict):
    if(dict == 'Freq'):
        paramFreqDict[keyParam] = valueFreq
    if(dict == 'Const'):
        constDict[keyParam] = valueFreq
#################### Formula relations - G vs Param vs Constants #############
def CalcGValues(gname, parList, constList):
    ### GU lighting ###
    if(gname == 'G1'):
        P16 = parList[15] ## Since the index starts from zero.
        K18 = constList[17]
        K02 = constList[1]

        g = (52.14 * (4 * 7 * P16 * K18) * K02)

    ###GU TV,PC.... ####
    if(gname == 'G2'):
        K02 = constList[1]
        K14 = constList[13]
        P12 = parList[11]
        K15 = constList[14]
        K16 = constList[15]
        P13 = parList[12]
        P14 = parList[13]
        K17 = constList[16]
        P15 = parList[14]

        g = (K02 * 52.14 * ((K14 * P12) + (7 * K15 * P13) + (7 * K16 * P14) + (7 * K17 * P15)))

    ### GU Kitchen ###
    if(gname == 'G3'):
        P01 = parList[0]
        K03 = constList[2]
        K01 = constList[0]
        P02 = parList[1]
        K04 = constList[3]
        P09 = parList[8]
        K11 = constList[10]
        K02 = constList[1]
        P10 = parList[9]
        K12 = constList[11]
        P11 = parList[10]
        K13 = constList[12]

        g = (52.14 * ((P01 * K03 * K01) + (P02 * K04 * K01) + (7 * P09 * K11 * K02) +
                     (P10 * K12 * 0.5 * K01) + (P11 * K13 * K01)))

    ### GU laundry ###
    if(gname == 'G4'):
        P07 = parList[6]
        K19 = constList[18]
        P08 = parList[7]
        K10 = constList[9]
        K02 = constList[1]

        g = 52.14 * ((P07 * K19) + (P08 * K10 * K02))
    return g
def GetGUTotal():
    return (guLighting + guTV + guKitchen + guLaundry)
def MakeGuList():
    list = [guLighting, guTV, guKitchen, guLaundry, guFood]
    return list
def RecommendationChecker():
    if(guTotal > 500):
        return True
    else:
        return False
def PickMaxGU():
    return guList.index(max(guList))
def GetGUParams(gu):
    if(gu == 0):
        return paramForGuLighting
    if(gu ==1):
        return paramForGUTV
    if(gu == 2):
        return paramForGUKitchen
    if(gu == 3):
        return paramForGULaundry
def ImpactAnalysis(new, old):
        return (old - new)
def PickMaxImpact(current, new):
    if(new >= current):
        return True
    else:
        return False
def FindParamImpact(index1, redRate, paramList):
    originalGU = guList[index1]
    defaultImpact = 0.0
    impactParam = []
    tentativePar = []
    redVal = 0.0
    reduction = 0.0
    for val in paramList:
        tentativePar = freqList[:] ## Python learning : Copies the values instead of reference.
        reduction = (tentativePar[val] * redRate)
        tentativePar[val] = reduction
        gnameStr = 'G'
        gnameStr += str(index1 + 1) #Note: G doesn't follow the index concept.
        newGU = CalcGValues(gnameStr, tentativePar, constValList)
        newImpact = ImpactAnalysis(newGU, originalGU)
        impactChecker = PickMaxImpact(defaultImpact, newImpact)
        if(impactChecker is True):
            impactParam.append(val)
            defaultImpact = newImpact
            redVal = reduction
    return impactParam, redVal, tentativePar
def StoreParListTemp(parlist, impactParams, presentGu, presentRate, paramChange):
    global guNow
    guNow = presentGu
    global tempParList
    tempParList = parlist[:]
    global reductionRate
    reductionRate = (presentRate - 0.2)
    global paramChangeFlag
    paramChangeFlag = paramChange

    if(paramChange is True):
        ##Note: Removal of parameter which has been reduced to 50%
        for val in impactParams:
            tempParList = tempParList.remove(val)

def GetParListTemp():
    return tempParList
def GetPresentGU():
    return guNow
def StoreFirstTimeFlag(str):
    global flag
    flag = str
def GetFirstTimeFlag():
    return flag
def GetRate():
    return reductionRate
def GetParamFlag():
    return paramChangeFlag
def ReduceEmissions(redRate):
    global freqList
    global guTotal
    global guList
    global parListTemp
    rate = redRate

    maxGUIndex = PickMaxGU()
    firstTimeFlag = GetFirstTimeFlag()
    if(firstTimeFlag is True):
        parListTemp = GetGUParams(maxGUIndex)
        StoreFirstTimeFlag('False')
    else:
        parListTemp = GetParListTemp()

    maxParImp, reducedValue, newFreqList = FindParamImpact(maxGUIndex, rate, parListTemp)
    freqList = newFreqList[:]
    guTotalNow = -1.0

    for val in maxParImp:
        freqList[val] = reducedValue
        gnameStr = 'G'
        gnameStr += str(maxGUIndex + 1)
        print(freqList)
        gu = CalcGValues(gnameStr, freqList, constValList)
        guTotalNow = ((guTotal - guList[maxGUIndex]) + gu)
        print(guTotalNow)

    if((guTotalNow > 500) & (rate > 0.5)):
        StoreParListTemp(parListTemp , maxParImp, gu, rate, False)
        guList[maxGUIndex] = gu
        guTotal = GetGUTotal()
        return True ## Note: Please keep in mind that while returning it has an updated freq list.


    elif((guTotalNow > 500) & (rate < 0.5)):
        StoreParListTemp(parListTemp , maxParImp, gu, rate, True)
        guList[maxGUIndex] = gu
        guTotal = GetGUTotal()
        return True
    return False

#################### Hardcoded values - File details ################################
file = 'BetaCalculator_Update 3.0.5.xlsx'
defaultSheet = 'Data Aggregation'
WB = opyxl.load_workbook(file, data_only= True)
WS = WB[defaultSheet]

################### Hardcoded values - Excel cell details ############################
freqCellIndex = 'G'
freqCellStart = 2
freqCellEnd = 17

paramCellIndex = 'F'
paramCellStart = 2
paramCellEnd = 17

constCellIndex = 'F'
constCellStart = 21
constCellEnd = 39

constValCellIndex = 'G'
constValCellStart = 21
constVallCellEnd = 39

#################### Hard coded values - GU-Param relations #########################
##Note: Index starts from zero.
paramForGuLighting = [15]
paramForGUTV = [11, 12, 13, 14]
paramForGUKitchen = [0, 1, 2, 3, 4, 5, 8, 9, 10]
paramForGULaundry = [6, 7]

################# Common variables for other functions ##############################
tempParList = []
flag = True
parListTemp = []
paramChangeFlag = False
################### Iteration function calls #########################################
frequency = ReadValues(freqCellIndex, freqCellStart, freqCellEnd)
freqList = list(map(float, frequency))
parameter = ReadValues(paramCellIndex, paramCellStart, paramCellEnd)

index = 0
paramFreqDict = {'Parameter' : 'Frequency' }
for val in parameter:
   # paramFreqDict = FreqParamMapping(val, frequency[index])
    DictMapping(val, frequency[index], 'Freq')
    index += 1

constants = ReadValues(constCellIndex, constCellStart, constCellEnd)
constVal = ReadValues(constValCellIndex, constValCellStart, constVallCellEnd)
constValList = list(map(float, constVal))
index = 0
constDict = {'Constants' : 'Values'}
for val in constants:
    DictMapping(val, constVal[index], 'Const')
guLighting = CalcGValues('G1', freqList, constValList)
guTV = CalcGValues('G2', freqList, constValList)
guKitchen = CalcGValues('G3', freqList, constValList)
guLaundry = CalcGValues('G4',  freqList, constValList)
guTotal = GetGUTotal()
guFood = 0.0 ## Hardcoded value
guList = MakeGuList()
recoFlag = RecommendationChecker()
reductionRate = 0.9
while (recoFlag is True):
    paramChangeFlag = GetParamFlag()
    if(paramChangeFlag == False):
        recoFlag = ReduceEmissions(reductionRate)
    else:
        reductionRate = 0.9
        recoFlag = ReduceEmissions(reductionRate)











