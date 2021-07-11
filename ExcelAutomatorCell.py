import openpyxl, os, xlsxwriter

#dir = r'C:\Users\Sean\Desktop\Elizabeth Script\Test Data'
#searchWord = 'example_value'
outputFile = 'Output Data.xlsx'
#sheetNumber = 0

dir = input('Enter Folder Path: ')
rowNumber = int(input('Enter row number (starting at 1): '))
colNumber = int(input('Enter column number (A=1, B=2, etc): '))
sheetNumber = int(input('Sheet Number (starting at 0): '))


def getFileList(dir):
    fileList = []
    for root, dirs, files in os.walk(dir):
        for file in files:
            if file.endswith('.xls') or file.endswith('.xlsx'):
                fileList.append(file)
    return fileList

def getDataFromFile(inputRow, inputCol, filename, sheetNumber):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.worksheets[sheetNumber]
    returnValue = sheet.cell(inputRow, inputCol).value
    return returnValue

def printListToExcel(list, filename):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    
    currentRow = 0
    for file, data in list:
        worksheet.write(currentRow, 0, file)
        worksheet.write(currentRow, 1, data)
        currentRow += 1
    
    workbook.close()

cleanedDir = dir.replace('\\', '/')

fileList = getFileList(dir)

list = []

for file in fileList:
    value = getDataFromFile(rowNumber, colNumber, dir+'/'+file, sheetNumber)
    list.append([file, value])
    
printListToExcel(list, outputFile)

print('Created Excel File\n')

end = input('Press Enter to continue')
