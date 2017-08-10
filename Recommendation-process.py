import openpyxl as opyxl
import matplotlib.pyplot as plt
import numpy
def GetUserDetails():
    name = WS['B1'].value
    return name

def GetOriginalFootPrint():
    footPrint = WS['B27'].value
    return footPrint

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
        K20 = constList[19]

        g = (52.14 * ((P01 * K03 * K01) + (P02 * K04 * K01) + (7 * P09 * K11 * K02) +
                     (P10 * K12 * 0.5 * K20) + (P11 * K13 * K20)))

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
       # print("old", old)
       # print("new", new)
        return (old - new)

def PickMaxImpact(current, new):
   # print("current", current)
   # print("new", new)

    if(new > current):
        return True
    else:
        return False

def FindParamImpact(index1, redRate, paramList):
    originalGU = guList[index1]
    defaultImpact = 0.0
    impactParam = [-1]
    tentativePar = []
    redVal = 0.0
    reduction = 0.0
    if(paramList is not None):
        for val in paramList:
            tentativePar = presentFreq[:] ## Python learning : Copies the values instead of reference.
            reduction = (tentativePar[val] * redRate) ## Reduction from original value.
            if(redRate == 0.7):
                reduction = (tentativePar[val] * 0.78) ## Since we need to reduce the value from 0.9% (Not from the original values)
            if(redRate == 0.5):
                reduction = (tentativePar[val] * 0.71)## Since we need to reduce the value from 0.7% (Not from the original values)
            tentativePar[val] = reduction 
            gnameStr = 'G'
            gnameStr += str(index1 + 1) #Note: G doesn't follow the index concept.
           # print("tentative par",tentativePar)
            newGU = CalcGValues(gnameStr, tentativePar, constValList)
            newImpact = ImpactAnalysis(newGU, originalGU)
            impactChecker = PickMaxImpact(defaultImpact, newImpact)
         #   print(impactChecker)

            if(impactChecker is True):
                ## Scope of enhancement in future. Now it is coded in a way that impact param list always has 1 element.
                ## In future, impact param can be changed to a float variable instead of list. But it needs good (..)
                ## (..)changes in the code.
                impactParam[0] = val
                defaultImpact = newImpact
             #   print("The reduction is ", reduction)
                redVal = reduction
           # print(impactParam)
   # print("tentative par", tentativePar)
    return impactParam, redVal, tentativePar

def StoreParListTemp(parlist, impactParams, presentGu, presentRate, paramChange):
    global parIndex
    global guNow
    guNow = presentGu
    global tempParList
    tempParList = parlist[:]
    global reductionRate
    if(reductionRate == 0.7):
        reductionRate = 0.5 ## Reason not subtracting 0.2 from 0.7 : To avoid 0.49999999999999994 bug
    else:
        reductionRate = (presentRate - 0.2)
    global paramChangeFlag
    paramChangeFlag = paramChange

    if(paramChange is True):
        ##Note: Removal of parameter which has been reduced to 50%
         # parIndex += 1
        if(impactParams[0] > -1):
            removePar = impactParams[0]
          #  print("Temp par list is ", tempParList)
          #  print("Remove par is", removePar)
            tempParList.remove(removePar)
          #  print("Impact params value is", impactParams[0])
        elif(impactParams[0] == -1):
            print("GU is saturated. Look for new one")
            tempParList = []
            

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

def SetPresentFrequency(freqlist):
    global presentFeq
    presentFreq = freqlist[:]
   # print("Present freq in set", presentFreq )

def GetPresentFrequency():
    global presentFreq
  #  print("Present freq in get", presentFreq)
    return presentFreq
    

def ReduceEmissions(redRate, originalFreqFlag, target):
    global originalFrequency
    global freqList
    global guTotal
    global guList
    global parListTemp
    global guTotalNow
    global presentFreq

    rate = redRate
    if(originalFreqFlag == True):
        freqList = originalFrequency
    else:
        freqList[:] = presentFreq
    maxGUIndex = PickMaxGU()
    firstTimeFlag = GetFirstTimeFlag()
    if(firstTimeFlag is True):
        parListTemp = GetGUParams(maxGUIndex)
        StoreFirstTimeFlag(False)
    else:
        parListTemp = GetParListTemp()
        #if(parIndex > 0):
           # parListTemp = parListTemp[parIndex:]
    #    print("par list temp", parListTemp)
    if(len(parListTemp) > 0):
        maxParImp, reducedValue, newFreqList = FindParamImpact(maxGUIndex, rate, parListTemp)
        freqList = newFreqList[:]
        presentFreq = freqList[:]
        gu = 0.0
        for val in maxParImp:
            freqList[val] = reducedValue
            gnameStr = 'G'
            gnameStr += str(maxGUIndex + 1)
            gu = CalcGValues(gnameStr, freqList, constValList)
            guTotalNow = ((guTotalNow - guList[maxGUIndex]) + gu)

        reducedParamList.append(maxParImp[0])

       # print(rate)
       # print(guTotalNow)
        if((guTotalNow > target) & (rate >0.5)):
            StoreParListTemp(parListTemp , maxParImp, gu, rate, False)
            guList[maxGUIndex] = gu
            guTotal = GetGUTotal()
            return True ## Note: Please keep in mind that while returning it has an updated freq list.

        elif((guTotalNow > target) & (rate == 0.5)):
            StoreParListTemp(parListTemp , maxParImp, gu, rate, True)
            guList[maxGUIndex] = gu
            guTotal = GetGUTotal()
            return True
        return False

def StoreOriginalFrequency(list):
    global originalFrequency
    originalFrequency = list[:]

def GetOriginalFrequency():
    return originalFrequency

## High and low target sets.
def SetTarget():
    if(guTotal >= 1000):
        targ = 1000
    else:
        targ = 500
    return targ

def WriteOutputFile():
    file = open("Recommendation-summary.txt", "w")
    file.write("****************************************************************************************************\n")
    file.write("                                     LEDSafari CO2 Calculator                                       \n")
    file.write("****************************************************************************************************\n")

    file.write("\n")
    file.write("Hi ")
    #file.write(userName)
   # ",", "\n")
    file.write(userName)
    file.write(",")
    file.write("\n")
    file.write("\n")
    file.write("Thank you once again for using LEDSafari's Carbon emission calculator.")
    file.write("\n")
    file.write("\n")
    file.write("Your recommendation details are given below :")
    file.write("\n")
    file.write("\n")
    file.write("The parameters and the corresponding frequencies are ")
    file.write("\n")
    file.write("\n")
    file.write("****************************************************************************************************\n")
    file.write("                                    Your current lifestyle                                          \n")
    file.write("****************************************************************************************************\n")
    ind = 0
    for i in parameter:
        file.write("-->")
        file.write(" ")
        file.write(i)
        file.write(" - ")
        file.write(str(originalFrequency[ind]))
        file.write("  ")
        if (i == 'Shower'):
            file.write('min/week')
        if(i == 'Bath'):
            file.write('times/week')
        if(i == 'Toilet'):
            file.write('times/day')
        if(i == 'Brushing teeth'):
            file.write('Washing hands')
        if(i == 'Shaving'):
            file.write('times/week')
        if(i == 'Washing'):
            file.write('kg/week')
        if(i == 'Drying'):
            file.write('kg/week')
        if(i == 'Boiling water (tea/coffee)'):
            file.write('L/day')
        if(i == 'Doing the dishes manually'):
            file.write('times/week')
        if(i == 'Dishwasher'):
            file.write('cycles/week')
        if(i == 'Television'):
            file.write('hours/week')
        if(i == 'Personal computer'):
            file.write('hours/day')
        if(i == 'Mobile phone'):
            file.write('hours/day')
        if(i == 'Tablet'):
            file.write('hours/day')
        if(i == 'Lighting technology around the house'):
            file.write('hours/day')

        file.write("\n")
        file.write("\n")
        ind +=1

    file.write("\n")

    file.write("****************************************************************************************************\n")
    file.write("                                  Suggested lifestyle                                               \n")
    file.write("****************************************************************************************************\n")
    file.write("\n")
    file.write(" Please note : In case of decimal values in some parameters and if decimal values are not feasible, you are kindly requested to round off to a lower possible minimum number.")
    file.write("\n")
    file.write("i.e In case of 12.6 in Dishwasher, you are requested to consider as 12 cycles/week.")
    file.write("\n")
    file.write("\n")
    ind1 = 0
    for i in parameter:
        file.write("-->")
        file.write(" ")
        file.write(i)
        file.write(" - ")
        file.write(str(freqList[ind1]))
        file.write("  ")
        if (i == 'Shower'):
            file.write('min/week')
        if(i == 'Bath'):
            file.write('times/week')
        if(i == 'Toilet'):
            file.write('times/day')
        if(i == 'Brushing teeth'):
            file.write('Washing hands')
        if(i == 'Shaving'):
            file.write('times/week')
        if(i == 'Washing'):
            file.write('kg/week')
        if(i == 'Drying'):
            file.write('kg/week')
        if(i == 'Boiling water (tea/coffee)'):
            file.write('L/day')
        if(i == 'Doing the dishes manually'):
            file.write('times/week')
        if(i == 'Dishwasher'):
            file.write('cycles/week')
        if(i == 'Television'):
            file.write('hours/week')
        if(i == 'Personal computer'):
            file.write('hours/day')
        if(i == 'Mobile phone'):
            file.write('hours/day')
        if(i == 'Tablet'):
            file.write('hours/day')
        if(i == 'Lighting technology around the house'):
            file.write('hours/day')
        file.write("\n")
        file.write("\n")
        ind1 +=1
    file.write("\n")
    file.write("\n")

    file.write("****************************************************************************************************\n ")
    file.write("                              Benefits of suggested lifestyle                                       \n")
    file.write("****************************************************************************************************\n")
    if(set_target == 500):
        target = 'High target'
    if(set_target == 1000):
        target = 'Low target'
    file.write("Congratulations ! With this suggested lifestyle, you have achieved : ")
    file.write(target)
    file.write("\n")
    file.write("\n")
    file.write("The Carbon emission as per your current lifestyle is ")
    file.write(str(guTotal))
    file.write(" ")
    file.write("kgCO2e ")
    file.write("\n")
    file.write("\n")
    file.write("The Carbon emission as per suggested lifestyle is ")
    file.write(str(guTotalNow))
    file.write(" ")
    file.write("kgCO2e ")
    file.write("\n")
    file.write("\n")
    file.close()


def displayMainOutput(flag):
    if(flag is True):
        print("\n")
        print("Hi ", userName, ",", "\n")
        print("Thank you for using LEDSafari's Carbon emission calculator. Please refer the text file named 'Recommendation-summary.txt'")
    else:
        print("Hi ", userName, ",", "\n")
        print("Thank you for using LEDSafari's Carbon emission calculator. Your current lifestyle is itself under a very good state. Cheers !")


def EquivalencyCalc(footPrint):
    ##**** Note: Based on carbon emission equivalency calculations ***##
    car = (footPrint / 2.98)
    plane = (footPrint / 4.007)
    tree = (footPrint / 2)

    return car, plane, tree

def GetColValues(workbook, col, start, end):
    colStr = col
    columnValList = []
    for val in range(start, end):
        colVal = (col + str(val))
        columnVal = workbook[colVal].value
        columnValList.append(columnVal)
    return columnValList

def GetStyleChanges():
    ## Getting Style change sheet here instead of data aggregation sheet.
    sheet = 'Style Change'
    WS_styleChange = WB[sheet]

    Col_A = GetColValues(WS_styleChange, 'A', 4, 17) ## Hardcoded value based on Style change sheet format.
    Col_C = GetColValues(WS_styleChange, 'C', 4, 17)

    changeList = []
    for i in range(len(Col_A)):
        if(Col_A[i] != Col_C[i]):
            changeList.append(i)

    return changeList, Col_A, Col_C


def CalcReductionPercent(option):
    if(option == 'Style'):
        numerator = (initialFootPrint - guTotal)
        denominator = initialFootPrint
        percent = (numerator / denominator) * 100
        return percent

def CalcFinalReduction(par):
    originalParVal = originalFrequency[par]
    reducedParVal = presentFreq[par]

    numerator = (originalParVal - reducedParVal)
    denominator = originalParVal
    percent = (numerator / denominator) * 100

    return percent

def WriteEnhancedOutputFile():
    file =  open("E-Recommendation-summary.txt", 'w')
    file.write("Hello ")
    file.write(userName)
    file.write(",")

    file.write("\n")
    file.write("\n")

    file.write("Thank you for using LEDSafari Carbon Calculator! We appreciate your interest towards greening your lifestyle and we hope you find our recommendations useful.")

    file.write("\n")
    file.write("\n")

    file.write("Based on your input, your current carbon footprint is ")
    footprint = round(initialFootPrint, 2)
    file.write(str(footprint))
    file.write(" kgCO2e")
    file.write(". This is equivalent to:")

    file.write("\n")
    file.write("\n")

    mediumCar, plane, trees = EquivalencyCalc(initialFootPrint)

    mediumCar = round(mediumCar, 2)
    plane = round(plane, 2)
    trees = round(trees, 2)

    file.write(str(mediumCar))
    file.write(" kilometeres in a medium car")

    file.write("\n")
    file.write("\n")

    file.write(str(plane))
    file.write(" kilometeres in a plane")

    file.write("\n")
    file.write("\n")

    file.write(str(trees))
    file.write(" trees absorbing CO2")

    file.write("\n")
    file.write("\n")

    file.write("The first step to switching to a green life is to make basic changes in the way you live. Based on the preferences, given by you in the Style change sheet, you are willing to make the following changes:")

    file.write("\n")
    file.write("\n")

    styleChangeIndices, A, C = GetStyleChanges()

    if(len(styleChangeIndices) > 0) :
        for val in styleChangeIndices:
            file.write("From")
            file.write("\n")
            file.write("\n")
            a_val = A[val]
            file.write(str(a_val))
            file.write("To")
            c_val = C[val]
            file.write(str(c_val))
            file.write("\n")
            file.write("\n")

        file.write("\n")
        file.write("\n")
    else:
        file.write("None")
        file.write("\n")
        file.write("\n")

    percentReduction = CalcReductionPercent('Style')
    percentReduction = round(percentReduction, 2)
    file.write("On making the changes, you have achieved a ")
    file.write(str(percentReduction))
    file.write(" % reduction in your emissions. Great work! Now your consumption levels are analysed for the most significant parameters. Based on the numbers, we make the following consumption reduction suggestions for you:")

    file.write("\n")
    file.write("\n")

    paramReductionList = list(set(reducedParamList))

    for val in paramReductionList:
        file.write("Reduce ")
        file.write("'")
        parVal = parameter[val]
        file.write(str(parVal))
        file.write("'")
        file.write(" to ")
        paramFinalRed = CalcFinalReduction(val)
        paramFinalRed = round(paramFinalRed, 2)
        file.write(str(paramFinalRed))
        file.write(" %.")

    file.write("\n")
    file.write("\n")

    file.write("On completion of the style and the consumption changes your final emissions are ")
    file.write(str(guTotalNow))
    file.write(" kgCO2e. By making these changes, your lifestyle has now reached the  ")
    if(set_target == 500):
        file.write('High ')
    else:
        file.write('Low ')
    file.write(" target of ")
    file.write(str(set_target))
    file.write(" kgCO2e per year. The results of the emissions saved is equivalent to:")

    file.write("\n")
    file.write("\n")

    mediumCar, plane, trees = EquivalencyCalc(guTotalNow)

    mediumCar = round(mediumCar, 2)
    plane = round(plane, 2)
    trees = round(trees, 2)

    file.write(str(mediumCar))
    file.write(" kilometeres in a medium car")

    file.write("\n")
    file.write("\n")

    file.write(str(plane))
    file.write(" kilometeres in a plane")

    file.write("\n")
    file.write("\n")

    file.write(str(trees))
    file.write(" trees absorbing CO2")

    file.write("\n")
    file.write("\n")

def PlotGraph():
    X = [1, 2, 3]
    labels = ['Current', 'Green User', 'Suggested']
    Y = [initialFootPrint, guTotal, guTotalNow]
    plt.title("Dip in Carbon emission after recommendation")
    plt.xlabel('Various stages')
    plt.ylabel('Carbon emission level (in KgCO2e units)')
    plt.xticks(X, labels)
    plt.plot(X, Y)
    plt.show()


##############################################################################################
#-------------------------- Main program starts here----------------------------------------#
##############################################################################################

#################### Step 1: Hardcoded values - Excel format details ################################
file = 'Data/source.xlsx'
defaultSheet = 'Data Aggregation'
WB = opyxl.load_workbook(file, data_only= True)
WS = WB[defaultSheet]

freqCellIndex = 'G'
freqCellStart = 2
freqCellEnd = 17

paramCellIndex = 'F'
paramCellStart = 2
paramCellEnd = 17

constCellIndex = 'F'
constCellStart = 21
constCellEnd = 40

constValCellIndex = 'G'
constValCellStart = 21
constVallCellEnd = 40

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
originalFrequency = []
#parIndex = 0
presentFreq = []
reducedParamList = []

####################################################################################################
####################### Step 2: Reads the values from Excel #########################################
#####################################################################################################

userName = GetUserDetails()
initialFootPrint = GetOriginalFootPrint()

frequency = ReadValues(freqCellIndex, freqCellStart, freqCellEnd)
freqList = list(map(float, frequency))
presentFreq = freqList[:]
StoreOriginalFrequency(freqList)
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
guTotalNow = guTotal
guFood = 0.0 ## Hardcoded value
guList = MakeGuList()

##########################################################################################
################### Step 3: Main functionalities ##########################################
#########################################################################################
recoFlag = RecommendationChecker()
displayMainOutput(recoFlag)
if(recoFlag is True):
    reductionRate = 0.9
    iterationCount = 0
    #print("gu list is ", guList)
    set_target = SetTarget()

    while (recoFlag is True):
        paramChangeFlag = GetParamFlag()
        if(paramChangeFlag == False):
            if(iterationCount == 0):
                recoFlag = ReduceEmissions(reductionRate, True, set_target)
            else:
                recoFlag = ReduceEmissions(reductionRate, False, set_target)
        else:
            guLighting = CalcGValues('G1', freqList, constValList)
            guTV = CalcGValues('G2', freqList, constValList)
            guKitchen = CalcGValues('G3', freqList, constValList)
            guLaundry = CalcGValues('G4',  freqList, constValList)
            guTotal = GetGUTotal()
            guTotalNow = guTotal
            guFood = 0.0 ## Hardcoded value
            guList = MakeGuList()
            #print("gu list now is ", guList)
            reductionRate = 0.9
            recoFlag = ReduceEmissions(reductionRate, False, set_target)
            #parIndex = 0
            iterationCount = 1
        presentFreq = freqList[:]

    print("\n")
    WriteOutputFile()
    WriteEnhancedOutputFile()
    PlotGraph()









