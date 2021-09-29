import zipfile

def getWordContent(path):
    with zipfile.ZipFile(path, "r") as docx:
        return str(docx.read('word/document.xml').decode('ascii'))

def getExcelContent(path):
    with zipfile.ZipFile(path, "r") as docx:
        strings = str(docx.read('xl/sharedStrings.xml').decode('ascii'))
        sheet1 = str(docx.read('xl/worksheets/sheet1.xml').decode('ascii'))
        return [strings, sheet1]

#print(getWordContent('word01.docx'))
print(getExcelContent('excel01.xlsx'))