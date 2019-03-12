import xlwings as xw
import math
import os
# data table root path
rootPath = r'E:\chenxiang\data\testdata'
# time table path
timePath = r'E:\chenxiang\data\shou_shu_time.xlsx'
outputPath = r'E:\chenxiang\data\output.xlsx'

fileNames = os.listdir(rootPath)
filePaths = []
for fileName in fileNames:
    filePaths.append(os.path.join(rootPath, fileName))


app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = False

timeTable = app.books.open(timePath)
timeList = []
try:
    timeSheet = timeTable.sheets['Sheet1']
    shape = timeSheet.range(1, 1).expand().shape
    timeList = timeSheet.range('A2:H' + str(shape[0])).value
finally:
    timeTable.save()
    timeTable.close()
    app.quit()

# print(timeList)
def findUserById(userid):
    for d in timeList:
        v = str(d[1]).strip()
        try:
            if int(math.floor(float(v))) == int(userid):
                return d 
        except:
            if str(v) == str(userid):
                return d

def getUserTime(userid):
    u = findUserById(userid)
    # print(u)
    if u:
        return u[2:8]
    else:
        print('NOT FOUND THE USER', userid)
    
timeDict = {}
for fileName in fileNames:
    t = getUserTime(fileName[:-4])
    if t:
        timeDict[fileName[:-4]] = t
    else:
        print('ERROR not found the time of %s' %(str(fileName)))

# print(timeDict)

# print(t1, t2)
def strListToIntList(strList):
    l = []
    for i in strList:
        l.append(int(i))
    return l

# if a >= b then return true
def isLargerTime(a, b):
    i = 0
    while i < len(a):
        if a[i] < b[i]:
            break
        elif a[i] == b[i]:
            i = i + 1
        else:
            return True
    return i == len(a)

def isInTimeRange(timeStart, timeEnd, t):
    # print(timeStart, timeEnd, t)
    tsList = strListToIntList(str(timeStart).split(':'))
    teList = strListToIntList(str(timeEnd).split(':'))
    tList = strListToIntList(str(t).split(':'))
    # print(tsList)
    # print(teList)
    # print(tList)
    if isLargerTime(tList, tsList) and isLargerTime(teList, tList):
        return True
    return False

def cleanData(l):
    r = []
    for v in l:
        if isValidVal(v):
            r.append(int(v))
        else:
            r.append('')
    return r
  
def validLen(l, validtype=[int, float]):
    count = 0
    for i in l:
        valid = False
        try:
            int(i)
            valid = True
        except:
            try:
                float(i)
                valid = True
            except:
                pass
        if valid:
            count = count + 1
    return count

def average(l, validtype=[int, float]):
    summ = 0
    for num in l:
        try:
            num = int(num)
        except:
            try:
                num = float(num)
            except:
                return ''
        summ = summ + num
    
    vlen = validLen(l,validtype=validtype)
    if vlen > 0:
        return summ / vlen
    return ''

def isValidVal(val):
    try:
        if int(val) >= 30 and int(val) <= 250:
            return True
    except:
        return False
    return False

def minKey(v):
    try:
        return int(v)
    except:
        pass
    try:
        return float(v)
    except:
        return 9999
    
def maxKey(v):
    try:
        return int(v)
    except:
        pass
    try:
        return float(v)
    except:
        return 0

def formatNumVal(l):
    r = []
    for num in l:
        val = num
        try:
            val = int(val)
        except:
            try:
                val = float(val)
            except:
                # print('error formatNumVal ', val)
                pass

        r.append(val)
    return r

def MNValidLen(l):
    m = formatNumVal(l)
    r = False
    for e in m:
        if e != 0:
            r = True
    return r

# filepath = r'C:\Users\Arenas\Desktop\chenxiang\137097.xls'
# dataTable = app.books.open(filepath)
outputTable = app.books.open(outputPath)
outputSheet = outputTable.sheets['Sheet1']

outputIndex = 0
for idx in range(len(filePaths)):
    filepath = filePaths[idx]
    dataTable = app.books.open(filepath)
    userName = fileNames[idx][:-4]
    # print('-----------------------------')
    # print('start handle %s' %(userName))
    try:
        sheet = dataTable.sheets[userName]
        shape = sheet.range(1, 1).expand().shape
        # extractData = []
        allData = sheet.range('A1:A'+str(shape[0])).value
        # allData = sheet.range('A1:A'+str(shape[0])).value
        tList = []
        GList = []
        MList = []
        NList = []
        X1List = []
        EList = []
        FList = []
        for d in allData:
            listd = d.split(' ')
            tList.append(listd[0])
            EList.append(listd[4])
            FList.append(listd[5])
            GList.append(listd[6])
            MList.append(listd[12])
            NList.append(listd[13])

        validEList = []
        validFList = []
        validX1List = []

        validEListT1 = []
        validFListT1 = []
        validMListT1 = []
        validNListT1 = []
        validGListT1 = []
        validX2ListT1 = []

        validEListT2 = []
        validFListT2 = []
        validMListT2 = []
        validNListT2 = []
        validGListT2 = []
        validX2ListT2 = []

        timeOfUser = timeDict[userName]
        for i in range(len(tList)):
            if isInTimeRange(timeOfUser[4], timeOfUser[5], tList[i]):
                validEList.append(EList[i])
                validFList.append(FList[i])
                
            if isInTimeRange(timeOfUser[0], timeOfUser[1], tList[i]):
                validMListT1.append(MList[i])
                validNListT1.append(NList[i])
                validGListT1.append(GList[i])
                validEListT1.append(EList[i])
            
            if isInTimeRange(timeOfUser[2], timeOfUser[3], tList[i]):
                validMListT2.append(MList[i])
                validNListT2.append(NList[i])
                validGListT2.append(GList[i])
                validFListT1.append(FList[i])

        validEList = cleanData(validEList)
        validFList = cleanData(validFList)
        validMListT1 = cleanData(validMListT1)
        validNListT1 = cleanData(validNListT1)

        validGListT1 = formatNumVal(validGListT1)
        validGListT1 = cleanData(validGListT1)
        validEListT1 = formatNumVal(validEListT1)
        validEListT1 = cleanData(validEListT1)
        validFListT1 = formatNumVal(validFListT1)
        validFListT1 = cleanData(validFListT1)

        for i in range(len(validEList)):
            if isValidVal(validEList[i]) and isValidVal(validFList[i]):
                X1List.append(validFList[i] + (validEList[i] - validFList[i])/3)
            else:
                X1List.append('')
        
        if not MNValidLen(validMListT1) or not MNValidLen(validNListT1):
            for i in range(len(validEListT1)):
                if isValidVal(validEListT1[i]) and isValidVal(validFListT1[i]):
                    validX2ListT1.append(validFListT1[i] + (validEListT1[i] - validFListT1[i])/3)
                else:
                    validX2ListT1.append('')
        else:
            for i in range(len(validMListT1)):
                if isValidVal(validMListT1[i]) and isValidVal(validNListT1[i]):
                    validX2ListT1.append(validNListT1[i] + (validMListT1[i] - validNListT1[i])/3)
                else:
                    validX2ListT1.append('')

        if not MNValidLen(validMListT2) or not MNValidLen(validNListT2):
            for i in range(len(validEListT2)):
                if isValidVal(validEListT2[i]) and isValidVal(validFListT2[i]):
                    validX2ListT2.append(validFListT2[i] + (validEListT2[i] - validFListT2[i])/3)
                else:
                    validX2ListT2.append('')
        else:
            for i in range(len(validMListT2)):
                if isValidVal(validMListT2[i]) and isValidVal(validNListT2[i]):
                    validX2ListT2.append(validNListT2[i] + (validMListT2[i] - validNListT2[i])/3)
                else:
                    validX2ListT2.append('')


        print('validMListT1 len == ', len(validMListT1))
        print('validMListT1 min == ', min(validMListT1, key = minKey))
        print('validMListT1 max == ', max(validMListT1, key = maxKey))
        print('validMListT1 avg == ', average(validMListT1))

        print('validNListT1 len == ', len(validNListT1))
        print('validNListT1 min == ', min(validNListT1, key = minKey))
        print('validNListT1 max == ', max(validNListT1, key = maxKey))
        print('validNListT1 avg == ', average(validNListT1))

        print('validGListT1 len == ', len(validGListT1))
        print('validGListT1 min == ', min(validGListT1, key = minKey))
        print('validGListT1 max == ', max(validGListT1, key = maxKey))
        print('validGListT1 avg == ', average(validGListT1))

        print('validX2ListT1 len == ', len(validX2ListT1))
        print('validX2ListT1 min == ', min(validX2ListT1, key = minKey))
        print('validX2ListT1 max == ', max(validX2ListT1, key = maxKey))
        print('validX2ListT1 avg == ', average(validX2ListT1))
        # print('X1List len == ', len(X1List))
        # print('X1 min == ', min(X1List, key = minKey))
        # print('X1 max == ', max(X1List, key = maxKey))
        # print('X1 avg == ', average(X1List))

        # print('validElist len == ', len(validEList))
        # print('E min == ', min(validEList, key = minKey))
        # print('E max == ', max(validEList, key = maxKey))
        # print('E avg == ', average(validEList))

        # print('validFlist len == ', len(validFList))
        # print('F min == ', min(validFList, key = minKey))
        # print('F max == ', max(validFList, key = maxKey))
        # print('F avg == ', average(validFList))
        # extractData.append(tList)
        # extractData.append(GList)
        # extractData.append(MList)
        # extractData.append(NList)
        # for i in range(10):
        #     extractData.append(sheet[i, 0].value)
        # print(extractData)
        # index = 0
        # # print(tList)
        # validTList = []
        # validGList = []
        # validMList = []
        # validNList = []
        # # print(timeOfUser)
        # for i in range(len(tList)):
        #     if isInTimeRange(timeOfUser[0], timeOfUser[1], tList[i]) or isInTimeRange(timeOfUser[2], timeOfUser[3], tList[i]):
        #         validTList.append(tList[i])
        #         validGList.append(GList[i])
        #         validMList.append(MList[i])
        #         validNList.append(NList[i])

        # # print('extract valid time range data:', len(validTList))
        # validGList = cleanData(validGList)
        # validMList = cleanData(validMList)
        # validNList = cleanData(validNList)
        # print('cleaned GList len = ', len(validGList))
        # print('cleaned MList len = ', len(validMList))
        # print('cleaned NList len = ', len(validNList))
        # print('GList max = %d, min = %d, avg = %d' % (max(validGList), min(validGList), average(validGList)))
        # print('MList max = %d, min = %d, avg = %d' % (max(validMList), min(validMList), average(validMList)))
        # print('NList max = %d, min = %d, avg = %d' % (max(validNList), min(validNList), average(validNList)))

        # outputSheet.range('A' + str(outputIndex + 2)).value = userName
        # outputSheet.range('B' + str(outputIndex + 2)).value = min(validGList)
        # outputSheet.range('C' + str(outputIndex + 2)).value = max(validGList)
        # outputSheet.range('D' + str(outputIndex + 2)).value = average(validGList)
        # outputSheet.range('E' + str(outputIndex + 2)).value = min(validMList)
        # outputSheet.range('F' + str(outputIndex + 2)).value = max(validMList)
        # outputSheet.range('G' + str(outputIndex + 2)).value = average(validMList)
        # outputSheet.range('H' + str(outputIndex + 2)).value = min(validNList)
        # outputSheet.range('I' + str(outputIndex + 2)).value = max(validNList)
        # outputSheet.range('J' + str(outputIndex + 2)).value = average(validNList)
        # outputIndex = outputIndex + 1
    finally:
        dataTable.save()
        dataTable.close()
        # app.quit()
    
    # print('end of handle %s' %(userName))
    # print('-----------------------------')

outputTable.save()
outputTable.close()
app.quit()
# filepath = r'E:\my\testexcel\test.xls'
# dataTable = app.books.open(filepath)
# a = wb.sheets['137097'].range('A1:A20').value
# print(a)
# for i in range(20):
#     a = wb.sheets['Sheet1'][i, 0].value
#     print(a)
# wb.save()
# wb.close()
# app.quit()