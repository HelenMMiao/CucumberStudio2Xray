import xlrd
import openpyxl
import re

'''This is to read from Cucumber Studio test cases'''
book = xlrd.open_workbook("CucumberStudio.xlsx")
sheet = book.sheet_by_name("CucumberStudio")
# stepGroup = ""
sheetRow = 0

def funcStepGroup():
    global sheetRow
    stepGroup = ""
    Ai, Bi = sheet.row_values(sheetRow, 0, 2)
    sheetRow += 1
    AiNext, BiNext = sheet.row_values(sheetRow, 0, 2)
    stepGroup = stepGroup + Ai + "\n"
    if AiNext:
        stepGroup += funcStepGroup()
        return stepGroup
    else:
        # print("One step is merged.")
        return stepGroup

def funcResultGroup():
    global sheetRow
    resultGroup = ""
    Ai, Bi = sheet.row_values(sheetRow, 0, 2)
    sheetRow += 1
    AiNext, BiNext = sheet.row_values(sheetRow, 0, 2)
    resultGroup = resultGroup + Bi + "\n"
    if BiNext:
        return resultGroup + funcResultGroup()
    else:
        # print("One result is merged.")
        return resultGroup

'''Merge the next to steps/results, and then put them as one step'''
def funcActionResult(actionsResults):
    actionGroup = funcStepGroup()
    resultGroup = funcResultGroup()
    actionsResults.append([actionGroup, resultGroup])
    Ai, Bi = sheet.row_values(sheetRow, 0, 2)
    if Ai or Bi:
        funcActionResult(actionsResults)
        return actionsResults
    else:
        return actionsResults



'''Write to another EXCEL file in Xray format'''
bookXray = openpyxl.Workbook()
shXray = bookXray.active
shXray.title = "SCS"

shXray['A1'] = 'TestID'
shXray['B1'] = 'Summary'
shXray['C1'] = 'Description'
shXray['D1'] = 'Step'
shXray['E1'] = 'Expected Result'
shXray['F1'] = 'Test Status'
shXray['G1'] = 'Test Type'
shXray['H1'] = 'Test Repository Path'

testID = 1
testStatus = "PEER REVIEW"
testType = "Manual"
repoPath =""
isFirstStep = True

lineXray = firstLineXray = 2
while sheetRow<67:
    Ai, Bi = sheet.row_values(sheetRow, 0, 2)

    # Write the first line for each case
    shXray[f'A{firstLineXray}'] = testID
    shXray[f'F{firstLineXray}'] = testStatus
    shXray[f'G{firstLineXray}'] = testType
    shXray[f'H{firstLineXray}'] = repoPath

    # Write into test case summary
    if Ai == "Scenario ID":
        shXray[f'B{firstLineXray}'] = Bi
        sheetRow += 1
        continue

    # Write into test case description
    if Ai == "Description":
        shXray[f'C{firstLineXray}'] = Bi
        sheetRow += 1
        continue

    #Write into test case steps
    if Ai == "Action":
        sheetRow += 1
        stepResultfor1case = []
        funcActionResult(stepResultfor1case)
        for eachStep in stepResultfor1case:
            shXray[f'D{lineXray}'] = eachStep[0]
            shXray[f'E{lineXray}'] = eachStep[1]
            shXray[f'A{lineXray}'] = testID
            lineXray += 1
        firstLineXray = lineXray
        testID += 1
        sheetRow += 1
        continue
    # Write into repo folder
    if Ai == "Folder name:":
        if not repoPath:
            repoPath = Bi
        if "flow" not in Bi:
            repoPath = Bi
        elif "Main" in Bi:
            repoPath = repoPath + "/" + Bi
            shXray[f'H{firstLineXray}'] = repoPath
        else:
            repoPath = re.sub('/(.*$)', f'/{Bi}', repoPath)
            shXray[f'H{firstLineXray}'] = repoPath
            sheetRow += 1
            continue
    sheetRow += 1
bookXray.save("XrayTC.xlsx")
