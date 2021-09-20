import zipfile

def getWordContent(path):
    with zipfile.ZipFile(path, "r") as docx:
        return str(docx.read('word/document.xml').decode('ascii'))

def getExcelContent(path):
    with zipfile.ZipFile(path, "r") as docx:
        strings = str(docx.read('xl/sharedStrings.xml').decode('ascii'))
        return [strings]

#print(getWordContent('word01.docx'))
print(getExcelContent('excel01.xlsx'))