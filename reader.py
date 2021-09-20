import zipfile

with zipfile.ZipFile("word01.docx","r") as docx:
    file = str(docx.read('word/document.xml').decode('ascii'))

print(file)