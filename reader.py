import zipfile
import re
import string 

def getWordContent(path):
    with zipfile.ZipFile(path, "r") as docx:
        return str(docx.read('word/document.xml').decode('ascii'))

def lettersToValue(input):
    counter = 0
    pos = len(input)
    for letter in input:
        factor = 26 ** pos
        print(pos)
        counter += factor * string.ascii_uppercase.index(letter)
        pos -= 1
    return counter

def getExcelSheetSize(dimensions):
    start = re.split('(\d+)', dimensions[0])[0:2]
    end = re.split('(\d+)', dimensions[1])[0:2]
    startPos = [lettersToValue(start[0]), int(start[1])]
    endPos = [lettersToValue(end[0]), int(end[1])]
    return [startPos, endPos]

def getExcelContent(path):
    with zipfile.ZipFile(path, "r") as docx:
        content = str(docx.read('xl/sharedStrings.xml').decode('ascii')).replace('\r\n', '')
        strings = content.split('</t></si><si><t>')
        for i in range(len(strings)):
            strings[i] = re.sub('<.*>', '', strings[i])
        sheet1 = str(docx.read('xl/worksheets/sheet1.xml').decode('ascii'))
        dimPat = re.compile('dimension ref="(.+)"/><sheetViews')
        dimensions = dimPat.findall(sheet1)[0].split(':')
        sheetSize = getExcelSheetSize(dimensions)
        return [strings, sheet1]

#print(getWordContent('word01.docx'))

#excelContent = getExcelContent('excel01.xlsx')
#print(excelContent)

print(getExcelSheetSize(['A1', 'AC5']))