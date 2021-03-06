import os
import shutil

originPath = r'D:\code\chenxiang\data\data'
desPath = r'D:\code\chenxiang\data\errorData'
logFile = r'D:\code\chenxiang\logger1.log'

fileNames = []
with open(logFile, 'r') as lf:
    for line in lf.readlines():
        fileNames.append(line[line.find('handle') + 7:].replace('\n', '') + '.xls')

# print(fileNames)
for i in fileNames:
    oldFile = os.path.join(originPath, i)
    newFile = os.path.join(desPath, i)
    os.rename(oldFile, newFile)
