import openpyxl

wb1 = openpyxl.load_workbook('xxx.xlsx')
sheetTotal = wb1['Sheet1']

flagsTotal = sheetTotal['A']
resultTotal = []
for i in range(1, len(flagsTotal)):
    tmp = {}
    tmp['index'] = i+1
    tmp['value'] = flagsTotal[i].value
    resultTotal.append(tmp)

wb2 = openpyxl.load_workbook('result.xlsx')
sheetToFind = wb2.active
flagsToFind = sheetToFind['A']
resultToFind = []
for i in range(1, len(flagsToFind)):
    tmp = {}
    tmp['index'] = i+1
    tmp['value'] = flagsToFind[i].value
    resultToFind.append(tmp)

for f in resultTotal:
    isFind = True
    for toFind in resultToFind:
        if toFind['value'] == f['value'] :
            isFind = False
            break
    if isFind == True:
        print(str(f['index']) + ' : ' + f['value'])
