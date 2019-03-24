import xlwings as xw
import math
import os
import logging
import traceback
import datetime
starttime = datetime.datetime.now()
logging.basicConfig(filename='logger.log', level=logging.INFO)

# rootPath = r'D:\code\chenxiang\data\errorData'
# rootPath = r'D:\code\chenxiang\data\testdata'
rootPath = r'D:\code\chenxiang\data\data'
timePath = r'D:\code\chenxiang\data\baguan10.xlsx'
# timePath = r'D:\code\chenxiang\data\testtime.xlsx'
outputPath = r'D:\code\chenxiang\data\output_baguan.xlsx'
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
    timeList = timeSheet.range('A2:C' + str(shape[0])).value
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
        return u[2]
    else:
        logging.info('NOT FOUND USER %s' % (userid))

timeDict = {}
for fileName in fileNames:
    t = getUserTime(fileName[:-4])
    if t:
        timeDict[fileName[:-4]] = t
    else:
        logging.info('NOT FOUND THE TIME OF  %s' % (str(fileName)))

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
                    retList.append(i)
            else:
                if float(i) > m:
                    c = c + 1
                    retList.append(i)
    return c, retList

def countLower(l, m, equ=True):
    c = 0
    retList = []
    for i in l:
        if i != '':
            if equ:
                if float(i) <= m:
                    c = c + 1
                    retList.append(i)
            else:
                if float(i) < m:
                    c = c + 1
                    retList.append(i)
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

def getStartIdxOfCow(time, tList, cowData):
    startTimeIdx = None
    for i in range(len(tList)):
        if isLargerTime(strListToIntList(str(tList[i]).split(':')), strListToIntList(str(time).split(':'))):
            if cowData[i] and isValidVal(cowData[i]):
                startTimeIdx = i
                return i, tList[i], cowData[i]

def getRangeIdx(startTime, endTime, tList, data):
    startTime = strListToIntList(str(startTime).split(':'))
    endTime = strListToIntList(str(endTime).split(':'))
    for i in range(len(tList)):
        tmpTime = strListToIntList(str(tList[i]).split(':'))
        if isLargerTime(tmpTime, startTime) and isLargerTime(endTime, tmpTime):
            if data[i] and isValidVal(data[i]):
                return i, int(data[i])

    return None, None

def longerSec(timePoint, sec):
    l = timePoint.split(':')
    miao = (int(l[2]) + sec) % 60
    fen = (int(l[1]) + math.floor((int(l[2]) + sec) / 60)) % 60
    shi = (int(l[0]) + math.floor((int(l[1]) + int(l[2]) + sec)/120))%24
    r = str(shi) + ':' + str(fen)+ ':' + str(miao)
    return r

def earlierSec(timePoint, sec):
    l = timePoint.split(':')
    subMiao = int(l[2]) - sec
    miao = subMiao % 60
    fen, shi = int(l[1]), int(l[0])
    if subMiao < 0:
        subFen = fen - 1
        fen = subFen % 60
        if subFen < 0:
            shi = (shi - 1) % 24
    r = str(shi) + ':' + str(fen)+ ':' + str(miao)
    return r

def parseData(timePoint, tList, data1, data2):
    res = []
    for i in range(1, 22):
        timePoint = longerMin(timePoint, (i-1) * 1)
        longerTenSec = longerSec(timePoint, 10)
        earlierTenSec = earlierSec(timePoint, 10)
        idx, data = getRangeIdx(timePoint, longerTenSec, tList, data1)

        if not idx:
            idx, data = getRangeIdx(earlierTenSec, timePoint, tList, data1)
        
        if not idx and data2:
            idx, data = getRangeIdx(timePoint, longerTenSec, tList, data2)

        if not idx and data2:
            idx, data = getRangeIdx(earlierTenSec, timePoint, tList, data2)

        if not idx:
            data = 0

        res.append(data)

    return res

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

            tList, GList, MList, NList, EList, FList, XList= [], [], [], [], [], [], []

            for d in allData:
                listd = d.split(' ')
                tList.append(listd[0])
                EList.append(listd[4])
                FList.append(listd[5])
                GList.append(listd[6])
                MList.append(listd[12])
                NList.append(listd[13])

            timeOfUser = timeDict[userid]
            # get M
            newMList = parseData(timeOfUser, tList, MList, EList)
            # print(newMList)
            newNList = parseData(timeOfUser, tList, NList, FList)
            # print(newNList)
            newGList = parseData(timeOfUser, tList, GList, None)
            # print(newGList)
            for i in range(len(newMList)):
                if newMList[i] > 0:
                    XList.append(round(newNList[i] + (newMList[i] - newNList[i])/3, 3))
                else:
                    XList.append(0)
            # print(XList)

            outputSheet.range('A'+str(rowIdx)).value = userName
            outputSheet.range('B'+str(rowIdx)).value = userid
            outputSheet.range('C'+str(rowIdx)).value = newMList
            outputSheet.range('X'+str(rowIdx)).value = newNList
            outputSheet.range('AS'+str(rowIdx)).value = newGList
            outputSheet.range('BN'+str(rowIdx)).value = XList
            rowIdx = rowIdx + 1
        except:
            traceback.print_exc() 
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