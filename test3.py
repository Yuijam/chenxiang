import xlwings as xw
import os
dataPath = r'E:\my\chenxiang\test2data\data1.xlsx'
outputPath = r'E:\my\chenxiang\test2data\output1.xlsx'

app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = False

dataTable = app.books.open(dataPath)
outputTable = app.books.open(outputPath)

def parseNum(s):
    if not s:
        return ''
    if (type(s) == str) and s.strip() == '':
        return ''
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
    for i in range(len(expList)):
        numList = parseNum(expList[i])
        if numList != '':
            sumList.append(float('%.3f' % sum(numList)))
        else:
            sumList.append('')
    return sumList

try:
    nameData = dataTable.sheets['Sheet1'].range('A2:A917').value
    for i in range(len(nameData)):
        outputTable.sheets['Sheet1'].range('A' + str(i + 2)).value = nameData[i].strip()

    roomNumData = dataTable.sheets['Sheet1'].range('C2:C917').value
    for i in range(len(roomNumData)):
        outputTable.sheets['Sheet1'].range('B' + str(i + 2)).value = roomNumData[i]

    cowConfig = {'S':'C', 'T':'D', 'U':'E'}
    for inputCow, outputCow in cowConfig.items():
        expList = dataTable.sheets['Sheet1'].range(inputCow + '2:' + inputCow + '917').value
        sumList = handleExpList(expList)
        for i in range(len(sumList)):
            outputTable.sheets['Sheet1'].range(outputCow + str(i + 2)).value = sumList[i]

finally:
    dataTable.save()
    dataTable.close()
    outputTable.save()
    outputTable.close()
    app.quit()