import xlwings as xw
import math
import os
import logging
import datetime
starttime = datetime.datetime.now()
logging.basicConfig(filename='logger.log', level=logging.INFO)

TYPE_ALL_M = 1
TYPE_ALL_E = 2
TYPE_HALF_E = 3
TYPE_ALL_N = 4
TYPE_ALL_F = 5
TYPE_HALF_F = 6

# rootPath = r'D:\code\chenxiang\data\errorData'
rootPath = r'D:\code\chenxiang\data\testdata'
# rootPath = r'D:\code\chenxiang\data\data'
timePath = r'D:\code\chenxiang\data\shou_shu_time.xlsx'
# timePath = r'D:\code\chenxiang\data\testtime.xlsx'
outputPath = r'D:\code\chenxiang\data\output_HALF_MN.xlsx'
# outputPath = r'D:\code\chenxiang\data\testoutput.xlsx'

curHandleUserId = 0

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

def findUserById(userid):
    for d in timeList:
        v = str(d[1]).strip()
        # v = str(d[1])
        try:
            if int(math.floor(float(v))) == int(userid):
                return d 
        except:
            if str(v) == str(userid).strip():
                return d

def getUserName(userid):
    u = findUserById(userid)
    if u:
        return u[0]
    else:
        print('ERROR can not find user', userid)

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

def strListToIntList(strList):
    l = []
    for i in strList:
        l.append(int(i))
    return l

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
    if isLargerTime(timeEnd, timeStart):
        if isLargerTime(tList, tsList) and isLargerTime(teList, tList):
            return True
    else:
        if isLargerTime(tList, tsList) or isLargerTime(teList, tList):
            return True
    return False

def isValidVal(val):
    try:
        if int(val) >= 30 and int(val) <= 250:
            return True
    except:
        return False
    return False

def cleanData(l, defaultVal = ''):
    r = []
    for v in l:
        if isValidVal(v):
            r.append(int(v))
        else:
            r.append(defaultVal)
    return r

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
                num = 0
        summ = summ + num
    
    vlen = validLen(l,validtype=validtype)
    # print('vlen = ', vlen)
    if vlen > 0:
        return round(summ / vlen, 3)
    return ''

def maxM(m, e):
    m = max(m, key = maxKey)
    if m == '':
        return max(e, key = maxKey)
    return m

def maxN(n, f):
    return maxM(n, f)

def minM(m, e):
    m = min(m, key = minKey)
    if m == '':
        return min(e, key = minKey)
    return m

def minN(n, f):
    return minM(n, f)

def averageM(m, e):
    avg = average(m)
    if avg == '':
        return average(e)
    return avg

def averageN(n, f):
    return averageM(n, f)

def countLarger(l, m, equ=True):
    c = 0
    retList = []
    for i in l:
        if i != '':
            if equ:
                if float(i) >= m:
                    c = c + 1
                    retList.append(round(i, 3))
            else:
                if float(i) > m:
                    c = c + 1
                    retList.append(round(i, 3))
    return c, retList

def countLower(l, m, equ=True):
    c = 0
    retList = []
    for i in l:
        if i != '':
            if equ:
                if float(i) <= m:
                    c = c + 1
                    retList.append(round(i, 3))
            else:
                if float(i) < m:
                    c = c + 1
                    retList.append(round(i, 3))
    return c, retList

def isEEmpty(e):
    for i in e:
        if i != '' and i != '0000':
            return False
    return True

def isFEmpty(f):
    return isEEmpty(f)

def longerMin(userTime, mini):
    l = userTime.split(':')
    fen = (int(l[1]) + mini) % 60
    shi = (int(l[0]) + math.floor((int(l[1]) + mini)/60))%24
    miao = l[2]
    r = str(shi) + ':' + str(fen)+ ':' + miao
    return r

def longerListTimeRange(l, tList, startTime, endTime, longer=0):
    if isEEmpty(l):
        newList = []
        for i in range(len(tList)):
            if isInTimeRange(startTime, longerMin(endTime, longer), tList[i]):
                newList.append(EList[i])
        return newList
    else:
        return l

def handleEmpty(l, tList, startTime, endTime, dataName):
    if isEEmpty(l):
        newList = longerListTimeRange(l, tList, startTime, endTime, longer=2)
        if isEEmpty(newList):
            newList = longerListTimeRange(l, tList, startTime, endTime, longer=5)
        if isEEmpty(newList):
            logging.info('ERROR %s data %s is empty' % (curHandleUserId, dataName))
        return newList
    else:
        return l

def isAllRangeValid(l):
    invalidCount = 0
    for i, v in enumerate(l):
        if not isValidVal(v):
            invalidCount = invalidCount + 1
        else:
            invalidCount = 0

        if invalidCount >= 3:
            return False, i - 2
    return True, None

outputTable = app.books.open(outputPath)
outputSheet = outputTable.sheets['Sheet1']

rowIdx = 2
try:
    for idx in range(len(filePaths)):
        dataTable = None
        filepath = filePaths[idx]
        userid = fileNames[idx][:-4]
        curHandleUserId = userid
        try:
            userName = getUserName(userid)
            print('handle ', userid)
            dataTable = app.books.open(filepath)
            sheet = dataTable.sheets[0]
            shape = sheet.range(1, 1).expand().shape
            allData = sheet.range('A1:A'+str(shape[0])).value

            tList, MList, NList, EList, HList, FList = [], [], [], [], [], []

            for d in allData:
                listd = d.split(' ')
                tList.append(listd[0])
                EList.append(listd[4])
                FList.append(listd[5])
                HList.append(listd[7])
                MList.append(listd[12])
                NList.append(listd[13])

            timeOfUser = timeDict[userid]
            validMListT2 = []
            validNListT2 = []
            validEListT2 = []
            validFListT2 = []
            validHListT2 = []

            for i in range(len(tList)):
                if isInTimeRange(timeOfUser[2], timeOfUser[3], tList[i]):
                    validMListT2.append(MList[i])
                    validNListT2.append(NList[i])
                    validEListT2.append(EList[i])
                    validFListT2.append(FList[i])
                    validHListT2.append(HList[i])

            validEListT2 = handleEmpty(validEListT2, tList, timeOfUser[2], timeOfUser[3], 'validEListT2')
            validFListT2 = handleEmpty(validFListT2, tList, timeOfUser[2], timeOfUser[3], 'validFListT2')

            validMListT2 = cleanData(validMListT2)
            validNListT2 = cleanData(validNListT2)
            validEListT2 = cleanData(validEListT2)
            validFListT2 = cleanData(validFListT2)
            validHListT2 = cleanData(validHListT2)

            dataType = None
            finalMT2 = validMListT2
            finalNT2 = validNListT2

            if validLen(finalMT2) == 0:
                finalMT2 = validEListT2
                finalNT2 = validFListT2
                dataType = TYPE_ALL_E
            else:
                iaAllValid, invalidIdx = isAllRangeValid(finalMT2)
                if iaAllValid:
                    dataType = TYPE_ALL_M
                else:
                    dataType = TYPE_HALF_E
                    finalMT2 = []
                    for i in range(invalidIdx, len(validEListT2)):
                        finalMT2.append(validEListT2[i])
                    
                    finalNT2 = []
                    for i in range(invalidIdx, len(validFListT2)):
                        finalNT2.append(validFListT2[i])

            outputSheet.range('A'+str(rowIdx)).value = userName
            outputSheet.range('B'+str(rowIdx)).value = userid

            if dataType == TYPE_HALF_E:
                outputSheet.range('A'+str(rowIdx)).color = (151, 255, 255)  # blue
            elif dataType == TYPE_ALL_M:
                outputSheet.range('A'+str(rowIdx)).color = (255, 222, 173)  # yellow
            else:
                outputSheet.range('A'+str(rowIdx)).color = (250, 128, 114)  # red

            outputSheet.range('C'+str(rowIdx)).value = max(finalMT2, key = maxKey)
            outputSheet.range('D'+str(rowIdx)).value = min(finalMT2, key = minKey)
            outputSheet.range('E'+str(rowIdx)).value = average(finalMT2)

            outputSheet.range('F'+str(rowIdx)).value = max(finalNT2, key = maxKey)
            outputSheet.range('G'+str(rowIdx)).value = min(finalNT2, key = minKey)
            outputSheet.range('H'+str(rowIdx)).value = average(finalNT2)

            outputSheet.range('I'+str(rowIdx)).value = max(validHListT2, key = maxKey)
            outputSheet.range('J'+str(rowIdx)).value = min(validHListT2, key = minKey)
            outputSheet.range('K'+str(rowIdx)).value = average(validHListT2)

            rowIdx = rowIdx + 1
        except:
            logging.info('ERROR on handle %s' % (userid))
        finally:
            if dataTable:
                try:
                    dataTable.save()
                    dataTable.close()
                except:
                    logging.info('ERROR on save and close table %s' % (userid))
            
finally:
    outputTable.save()
    outputTable.close()
    app.quit()
    endtime = datetime.datetime.now()
    logging.info('Run %s seconds' % ((endtime - starttime).seconds))