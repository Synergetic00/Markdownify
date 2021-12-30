import zipfile
import re
import string 

def getWordContent(path):
    with zipfile.ZipFile(path, "r") as docx:
        return str(docx.read('word/document.xml').decode('ascii'))

def getValueOfLetters(input):
    counter = 0
    pos = 0
    for letter in input[::-1]:
        factor = 26 ** pos
        counter += factor * (string.ascii_uppercase.index(letter) + 1)
        pos += 1
    return counter - 1

def getExcelCellPosition(input):
    details = re.split('(\d+)', input)
    return [int(details[1])-1, getValueOfLetters(details[0])]

def getExcelSheetSize(dimensions):
    start = getExcelCellPosition(dimensions[0])
    end = getExcelCellPosition(dimensions[1])
    width = (end[0]-start[0]) + 1
    height = (end[1]-start[1]) + 1
    return [width, height]

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
        cellPat = re.compile('<c r="([A-Z]\w+)" t="s"><v>([0-9]+)<\/v><\/c>')
        cells = cellPat.findall(sheet1)
        array = [['' for x in range(sheetSize[1])] for y in range(sheetSize[0])]
        for cell in cells:
            pos = getExcelCellPosition(cell[0])
            array[int(pos[0])][int(pos[1])] = strings[int(cell[1])]
        return array

def makeBasicMarkdownTable(array, filename):
    output = ''
    for i in range(len(array)):
        output += '|' + '|'.join(array[i]) + '|\n'
        if i == 0:
            output += '|' + (':-:|' * len(array[0])) + '\n'
    with open(filename+'.md', 'w') as f:
        f.write(output.strip())


#print(getWordContent('word01.docx'))

makeBasicMarkdownTable(getExcelContent('Book1.xlsx'), 'mdExcel02')