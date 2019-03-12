import xlwings as xw
import os
# data table root path
# time table path
dataPath = r'E:\my\chenxiang\test2data\data.xlsx'
outputPath = r'E:\my\chenxiang\test2data\output.xlsx'

app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = False

dataTable = app.books.open(dataPath)
outputTable = app.books.open(outputPath)

def parseNum(s):
    s = s.strip()
    # print(s)
    starIdx = -1
    plusIdx = -1
    l = []
    for i in range(len(s)):
        if s[i] == '*':
            starIdx = i
        if starIdx != -1 and s[i] == '+':
            plusIdx = i

        num = None
        if starIdx != -1 and plusIdx != -1:
            num = float(s[starIdx+1:plusIdx])
            starIdx = -1
            plusIdx = -1
        
        if i == (len(s) - 1) and starIdx != -1:
            num = float(s[starIdx+1:])

        if num:
            l.append(num)
    return l

def handleExpList(expList):
    sumList = []
    for exp in expList:
        numList = parseNum(exp)
        sumList.append(float('%.3f' % sum(numList)))
    # print(bingPoFenSumList)
    return sumList

try:
    nameData = dataTable.sheets['Sheet1'].range('A2:A630').value
    for i in range(len(nameData)):
        # nameList.append(name.strip())
        outputTable.sheets['Sheet1'].range('A' + str(i + 2)).value = nameData[i].strip()

    roomNumData = dataTable.sheets['Sheet1'].range('C2:C630').value
    for i in range(len(roomNumData)):
        # nameList.append(name.strip())
        outputTable.sheets['Sheet1'].range('B' + str(i + 2)).value = roomNumData[i]

    cowConfig = {'N':'C', 'R':'D', 'S':'E'}
    for inputCow, outputCow in cowConfig.items():
        expList = dataTable.sheets['Sheet1'].range(inputCow + '2:' + inputCow + '630').value
        sumList = handleExpList(expList)
        for i in range(len(sumList)):
            outputTable.sheets['Sheet1'].range(outputCow + str(i + 2)).value = sumList[i]

finally:
    dataTable.save()
    dataTable.close()
    outputTable.save()
    outputTable.close()
    app.quit()

