import xlwings as xw
import os
# data table root path
rootPath = r'E:\my\chenxiang\table'
# time table path
timePath = r'E:\my\chenxiang\time.xlsx'
outputPath = r'E:\my\chenxiang\output.xlsx'

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
    timeList = timeTable.sheets['Sheet1'].range('A2:E6').value
finally:
    timeTable.save()
    timeTable.close()
    app.quit()

# print(timeList)
def findUserById(userid):
    for d in timeList:
        if int(d[0]) == int(userid):
            return d

def getUserTime(userid):
    u = findUserById(userid)
    # print(u)
    if u:
        return u[1:5]
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
        if int(v) >= 30 and int(v) <= 250:
            r.append(int(v))
    return r

def average(l):
    sum = 0
    for v in l:
        sum = sum + v
    return sum / len(l)

# filepath = r'C:\Users\Arenas\Desktop\chenxiang\137097.xls'
# dataTable = app.books.open(filepath)
outputTable = app.books.open(outputPath)
outputSheet = outputTable.sheets['Sheet1']

outputIndex = 0
for filepath in filePaths:
    dataTable = app.books.open(filepath)
    userName = filepath[-10:-4]
    # print('-----------------------------')
    # print('start handle %s' %(userName))
    try:
        sheet = dataTable.sheets[userName]
        shape = sheet.range(1, 1).expand().shape
        extractData = []
        allData = sheet.range('A1:A'+str(shape[0])).value
        # allData = sheet.range('A1:A'+str(shape[0])).value
        tList = []
        GList = []
        MList = []
        NList = []
        for d in allData:
            listd = d.split(' ')
            tList.append(listd[0])
            GList.append(listd[6])
            MList.append(listd[12])
            NList.append(listd[13])
        
        extractData.append(tList)
        extractData.append(GList)
        extractData.append(MList)
        extractData.append(NList)
        # for i in range(10):
        #     extractData.append(sheet[i, 0].value)
        # print(extractData)
        index = 0
        # print(tList)
        validTList = []
        validGList = []
        validMList = []
        validNList = []
        timeOfUser = timeDict[userName]
        # print(timeOfUser)
        for i in range(len(tList)):
            if isInTimeRange(timeOfUser[0], timeOfUser[1], tList[i]) or isInTimeRange(timeOfUser[2], timeOfUser[3], tList[i]):
                validTList.append(tList[i])
                validGList.append(GList[i])
                validMList.append(MList[i])
                validNList.append(NList[i])

        # print('extract valid time range data:', len(validTList))
        validGList = cleanData(validGList)
        validMList = cleanData(validMList)
        validNList = cleanData(validNList)
        # print('cleaned GList len = ', len(validGList))
        # print('cleaned MList len = ', len(validMList))
        # print('cleaned NList len = ', len(validNList))
        # print('GList max = %d, min = %d, avg = %d' % (max(validGList), min(validGList), average(validGList)))
        # print('MList max = %d, min = %d, avg = %d' % (max(validMList), min(validMList), average(validMList)))
        # print('NList max = %d, min = %d, avg = %d' % (max(validNList), min(validNList), average(validNList)))

        print('A' + str(outputIndex + 2))
        outputSheet.range('A' + str(outputIndex + 2)).value = userName
        outputSheet.range('B' + str(outputIndex + 2)).value = min(validGList)
        outputSheet.range('C' + str(outputIndex + 2)).value = max(validGList)
        outputSheet.range('D' + str(outputIndex + 2)).value = average(validGList)
        outputSheet.range('E' + str(outputIndex + 2)).value = min(validMList)
        outputSheet.range('F' + str(outputIndex + 2)).value = max(validMList)
        outputSheet.range('G' + str(outputIndex + 2)).value = average(validMList)
        outputSheet.range('H' + str(outputIndex + 2)).value = min(validNList)
        outputSheet.range('I' + str(outputIndex + 2)).value = max(validNList)
        outputSheet.range('J' + str(outputIndex + 2)).value = average(validNList)
        outputIndex = outputIndex + 1
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