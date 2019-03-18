import xlwings as xw
import math
import os
import logging
logging.basicConfig(filename='logger.log', level=logging.INFO)

# rootPath = r'D:\code\chenxiang\data\errorData'
# rootPath = r'D:\code\chenxiang\data\testdata'
rootPath = r'D:\code\chenxiang\data\data'
timePath = r'D:\code\chenxiang\data\shou_shu_time.xlsx'
# timePath = r'D:\code\chenxiang\data\testtime.xlsx'
outputPath = r'D:\code\chenxiang\data\output.xlsx'
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
    shi = int(l[0]) + math.floor(fen/60)
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


outputTable = app.books.open(outputPath)
outputSheet = outputTable.sheets['Sheet1']
rowIdx = 3
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

            tList, GList, MList, NList, X1List, EList, FList, HList  = [], [], [], [], [], [], [], []

            for d in allData:
                listd = d.split(' ')
                tList.append(listd[0])
                EList.append(listd[4])
                FList.append(listd[5])
                GList.append(listd[6])
                HList.append(listd[7])
                MList.append(listd[12])
                NList.append(listd[13])

            validEListT3, validFListT3, validGListT3 = [], [], []

            timeOfUser = timeDict[userid]
            for i in range(len(tList)):
                if isInTimeRange(timeOfUser[4], timeOfUser[5], tList[i]):
                    validEListT3.append(EList[i])
                    validFListT3.append(FList[i])
                    validGListT3.append(GList[i])

            validEListT3 = handleEmpty(validEListT3, tList, timeOfUser[4], timeOfUser[5], 'validEListT3')
            validFListT3 = handleEmpty(validFListT3, tList, timeOfUser[4], timeOfUser[5], 'validFListT3')
            validGListT3 = handleEmpty(validGListT3, tList, timeOfUser[4], timeOfUser[5], 'validGListT3')

            validEListT3 = cleanData(validEListT3)
            validFListT3 = cleanData(validFListT3)
            validGListT3 = cleanData(validGListT3)

            validEListT3Before = []
            validFListT3Before = []
            for i in range(len(tList)):
                if isInTimeRange('00:00:00', timeOfUser[5], tList[i]):
                    validEListT3Before.append(EList[i])
                    validFListT3Before.append(FList[i])
            
            validEListT3Before = handleEmpty(validEListT3Before, tList, '00:00:00', timeOfUser[5], 'validEListT3Before')
            validFListT3Before = handleEmpty(validFListT3Before, tList, '00:00:00', timeOfUser[5], 'validFListT3Before')

            validEListT3Before = cleanData(validEListT3Before)
            validFListT3Before = cleanData(validFListT3Before)

            for i in range(len(validEListT3Before)):
                try:
                    X1List.append(validFListT3Before[i] + (validEListT3Before[i] - validFListT3Before[i])/3)
                except:
                    # print('ERROR parse X1 i = %d , E = %s, F = %s' % (i, validEListT3[i], validFListT3[i]))
                    pass
                
            # print('X1List = ', X1List)

            validMListT1, validMListT2 = [], []
            validNListT1, validNListT2 = [], []
            validGListT1, validGListT2 = [], []
            validEListT1, validEListT2 = [], []
            validFListT1, validFListT2 = [], []
            validHListT1, validHListT2 = [], []

            for i in range(len(tList)):
                if isInTimeRange(timeOfUser[0], timeOfUser[1], tList[i]):
                    validMListT1.append(MList[i])
                    validNListT1.append(NList[i])
                    validGListT1.append(GList[i])
                    validEListT1.append(EList[i])
                    validFListT1.append(FList[i])
                    validHListT1.append(HList[i])

                if isInTimeRange(timeOfUser[2], timeOfUser[3], tList[i]):
                    validMListT2.append(MList[i])
                    validNListT2.append(NList[i])
                    validGListT2.append(GList[i])
                    validEListT2.append(EList[i])
                    validFListT2.append(FList[i])
                    validHListT2.append(HList[i])

            validEListT1 = handleEmpty(validEListT1, tList, timeOfUser[0], timeOfUser[1], 'validEListT1')
            validFListT1 = handleEmpty(validFListT1, tList, timeOfUser[0], timeOfUser[1], 'validFListT1')
            validGListT1 = handleEmpty(validGListT1, tList, timeOfUser[0], timeOfUser[1], 'validGListT1')

            validEListT2 = handleEmpty(validEListT2, tList, timeOfUser[2], timeOfUser[3], 'validEListT2')
            validFListT2 = handleEmpty(validFListT2, tList, timeOfUser[2], timeOfUser[3], 'validFListT2')
            validGListT2 = handleEmpty(validGListT2, tList, timeOfUser[2], timeOfUser[3], 'validGListT2')

            validMListT1 = cleanData(validMListT1)
            validNListT1 = cleanData(validNListT1)
            validGListT1 = cleanData(validGListT1)
            validEListT1 = cleanData(validEListT1)
            validFListT1 = cleanData(validFListT1)
            validHListT1 = cleanData(validHListT1)

            validMListT2 = cleanData(validMListT2)
            validNListT2 = cleanData(validNListT2)
            validGListT2 = cleanData(validGListT2)
            validEListT2 = cleanData(validEListT2)
            validFListT2 = cleanData(validFListT2)
            validHListT2 = cleanData(validHListT2)

            X2ListT1, X2ListT2 = [], []
            finalMT1 = validMListT1
            finalNT1 = validNListT1

            if validLen(finalMT1) == 0:
                finalMT1 = validEListT1
                finalNT1 = validFListT1
            
            for i in range(len(finalMT1)):
                try:
                    X2ListT1.append(finalNT1[i] + (finalMT1[i] - finalNT1[i])/3)
                except:
                    pass

            finalMT2 = validMListT2
            finalNT2 = validNListT2

            if validLen(finalMT2) == 0:
                finalMT2 = validEListT2
                finalNT2 = validFListT2
            
            for i in range(len(finalMT2)):
                try:
                    X2ListT2.append(finalNT2[i] + (finalMT2[i] - finalNT2[i])/3)
                except:
                    pass


            outputSheet.range('A'+str(rowIdx)).value = userName
            outputSheet.range('B'+str(rowIdx)).value = userid
            outputSheet.range('C'+str(rowIdx)).value = max(validEListT3, key = maxKey)
            outputSheet.range('D'+str(rowIdx)).value = min(validEListT3, key = minKey)
            # print('validEListT3 len == ', len(validEListT3))
            # print('E min == ', min(validEListT3, key = minKey))
            # print('E max == ', max(validEListT3, key = maxKey))
            Y = average(validEListT3)
            outputSheet.range('E'+str(rowIdx)).value = Y
            # print('E avg == ', Y)

            outputSheet.range('F'+str(rowIdx)).value = max(validFListT3, key = maxKey)
            outputSheet.range('G'+str(rowIdx)).value = min(validFListT3, key = minKey)
            outputSheet.range('H'+str(rowIdx)).value = average(validFListT3)
            # print('validFListT3 len == ', len(validFListT3))
            # print('F min == ', min(validFListT3, key = minKey))
            # print('F max == ', max(validFListT3, key = maxKey))
            # print('F avg == ', average(validFListT3))
            
            outputSheet.range('I'+str(rowIdx)).value = max(validGListT3, key = maxKey)
            outputSheet.range('J'+str(rowIdx)).value = min(validGListT3, key = minKey)
            outputSheet.range('K'+str(rowIdx)).value = average(validGListT3)
            # print('validGListT3 len == ', len(validGListT3))
            # print('G min == ', min(validGListT3, key = minKey))
            # print('G max == ', max(validGListT3, key = maxKey))
            # print('G avg == ', average(validGListT3))
            
            outputSheet.range('L'+str(rowIdx)).value = max(X1List, key = maxKey)
            outputSheet.range('M'+str(rowIdx)).value = min(X1List, key = minKey)
            # print('X1List len == ', len(X1List))
            # print('X1 min == ', min(X1List, key = minKey))
            # print('X1 max == ', max(X1List, key = maxKey))
            Y1 = average(X1List)
            outputSheet.range('N'+str(rowIdx)).value = Y1
            # print('X1 avg == ', Y1)

            #------------------------------------------------#
            # print('-----------------------------------------')
            # print('maxM(validMListT1, validEListT1) = ', maxM(validMListT1, validEListT1))
            outputSheet.range('O'+str(rowIdx)).value = maxM(validMListT1, validEListT1)
            outputSheet.range('P'+str(rowIdx)).value = minM(validMListT1, validEListT1)
            outputSheet.range('Q'+str(rowIdx)).value = averageM(validMListT1, validEListT1)
            # print('validMListT1 len == ', len(validMListT1))
            # print('validEListT1 len == ', len(validEListT1))
            # print('M min == ', minM(validMListT1, validEListT1))
            # print('M max == ', maxM(validMListT1, validEListT1))
            # print('M avg == ', averageM(validMListT1, validEListT1))

            outputSheet.range('R'+str(rowIdx)).value = maxN(validNListT1, validFListT1)
            outputSheet.range('S'+str(rowIdx)).value = minN(validNListT1, validFListT1)
            outputSheet.range('T'+str(rowIdx)).value = averageN(validNListT1, validFListT1)
            # print('validNListT1 len == ', len(validNListT1))
            # print('validFListT1 len == ', len(validFListT1))
            # print('N min == ', minN(validNListT1, validFListT1))
            # print('N max == ', maxN(validNListT1, validFListT1))
            # print('N avg == ', averageN(validNListT1, validFListT1))
            
            outputSheet.range('U'+str(rowIdx)).value = max(validGListT1, key = maxKey)
            outputSheet.range('V'+str(rowIdx)).value = min(validGListT1, key = minKey)
            outputSheet.range('W'+str(rowIdx)).value = average(validGListT1)
            # print('validGListT1 len == ', len(validGListT1))
            # print('G min == ', min(validGListT1, key = minKey))
            # print('G max == ', max(validGListT1, key = maxKey))
            # print('G avg == ', average(validGListT1))

            outputSheet.range('AI'+str(rowIdx)).value = maxM(validMListT2, validEListT2)
            outputSheet.range('AJ'+str(rowIdx)).value = minM(validMListT2, validEListT2)
            outputSheet.range('AK'+str(rowIdx)).value = averageM(validMListT2, validEListT2)
            # print('validMListT2 len == ', len(validMListT2))
            # print('validEListT2 len == ', len(validEListT2))
            # print('M min == ', minM(validMListT2, validEListT2))
            # print('M max == ', maxM(validMListT2, validEListT2))
            # print('M avg == ', averageM(validMListT2, validEListT2))

            outputSheet.range('AL'+str(rowIdx)).value = maxN(validNListT2, validFListT2)
            outputSheet.range('AM'+str(rowIdx)).value = minN(validNListT2, validFListT2)
            outputSheet.range('AN'+str(rowIdx)).value = averageN(validNListT2, validFListT2)
            # print('validNListT2 len == ', len(validNListT2))
            # print('validFListT2 len == ', len(validFListT2))
            # print('N min == ', minN(validNListT2, validFListT2))
            # print('N max == ', maxN(validNListT2, validFListT2))
            # print('N avg == ', averageN(validNListT2, validFListT2))

            outputSheet.range('AO'+str(rowIdx)).value = max(validGListT2, key = maxKey)
            outputSheet.range('AP'+str(rowIdx)).value = min(validGListT2, key = minKey)
            outputSheet.range('AQ'+str(rowIdx)).value = average(validGListT2)
            # print('validGListT2 len == ', len(validGListT2))
            # print('G min == ', min(validGListT2, key = minKey))
            # print('G max == ', max(validGListT2, key = maxKey))
            # print('G avg == ', average(validGListT2))
            
            outputSheet.range('Y'+str(rowIdx)).value = max(X2ListT1, key = maxKey)
            outputSheet.range('Z'+str(rowIdx)).value = min(X2ListT1, key = minKey)
            outputSheet.range('AA'+str(rowIdx)).value = average(X2ListT1)
            # print('X2ListT1 len == ', len(X2ListT1))
            # print('X2_T1 min == ', min(X2ListT1, key = minKey))
            # print('X2_T1 max == ', max(X2ListT1, key = maxKey))
            # print('X2_T1 avg == ', average(X2ListT1))
            
            outputSheet.range('AS'+str(rowIdx)).value = max(X2ListT2, key = maxKey)
            outputSheet.range('AT'+str(rowIdx)).value = min(X2ListT2, key = minKey)
            outputSheet.range('AU'+str(rowIdx)).value = average(X2ListT2)
            # print('X2ListT2 len == ', len(X2ListT2))
            # print('X2_T2 min == ', min(X2ListT2, key = minKey))
            # print('X2_T2 max == ', max(X2ListT2, key = maxKey))
            # print('X2_T2 avg == ', average(X2ListT2))

            outputSheet.range('X'+str(rowIdx)).value = countLarger(validGListT1, 110)[0]
            outputSheet.range('AB'+str(rowIdx)).value = countLarger(X2ListT1, 109, equ=False)[0]
            outputSheet.range('AC'+str(rowIdx)).value = countLower(X2ListT1, 60)[0]

            outputSheet.range('AR'+str(rowIdx)).value = countLarger(validGListT2, 110)[0]
            outputSheet.range('AV'+str(rowIdx)).value = countLarger(X2ListT2, 109, equ=False)[0]
            outputSheet.range('AW'+str(rowIdx)).value = countLower(X2ListT2, 60)[0]
            # print('count G_T1 >= 110 ==', countLarger(validGListT1, 110)[0])
            # print('count G_T2 >= 110 ==', countLarger(validGListT2, 110)[0])
            # print('count X2_T1 > 109 ==', countLarger(X2ListT1, 109, equ=False)[0])
            # print('count X2_T2 > 109 ==', countLarger(X2ListT2, 109, equ=False)[0])
            # print('count X2_T1 <= 60 ==', countLower(X2ListT1, 60)[0])
            # print('count X2_T2 <= 60 ==', countLower(X2ListT2, 60)[0])

            mT1LagerYCount, mT1LagerYList = countLarger(validMListT1, Y*1.25)
            # print('count M_T1 >= Y*1.25 ==', mT1LagerYCount)
            # outputSheet.range('AD'+str(rowIdx)).value = mT1LagerYList
            outputSheet.range('AE'+str(rowIdx)).value = mT1LagerYCount

            mT2LagerYCount, mT2LagerYList = countLarger(validMListT2, Y*1.25)
            # print('count M_T2 >= Y*1.25 ==', mT2LagerYCount)
            # outputSheet.range('AX'+str(rowIdx)).value = mT2LagerYList
            outputSheet.range('AY'+str(rowIdx)).value = mT2LagerYCount

            x2T1LagerYCount, x2T1LagerYList = countLarger(X2ListT1, Y1*1.25)
            # print('count X2_T1 >= Y1*1.25 ==', x2T1LagerYCount)
            # outputSheet.range('AF'+str(rowIdx)).value = x2T1LagerYList
            outputSheet.range('AG'+str(rowIdx)).value = x2T1LagerYCount

            x2T2LagerYCount, x2T2LagerYList = countLarger(X2ListT2, Y1*1.25)
            # print('count X2_T2 >= Y1*1.25 ==', x2T2LagerYCount)
            # outputSheet.range('AZ'+str(rowIdx)).value = x2T2LagerYList
            outputSheet.range('BA'+str(rowIdx)).value = x2T2LagerYCount

            # print('count H_T1 <= 90', countLower(validHListT1, 90)[0])
            outputSheet.range('AH'+str(rowIdx)).value = countLower(validHListT1, 90)[0]

            # print('count H_T2 <= 90', countLower(validHListT2, 90)[0])
            outputSheet.range('BB'+str(rowIdx)).value = countLower(validHListT2, 90)[0]
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