# Module 14 - Excel, Word and PDF Documents: Reading and Editing PDFs

import PyPDF2, os

os.chdir('/Users/paulmackay/Desktop/Python/Excel, Word and PDF Documents')

pfdFile = open('meetingminutes.pdf', 'rb') # Opens file in Read Binary mode.

reader = PyPDF2.PdfFileReader(pdfFile)
reader.numPages # Returns 19 as an integer

page = reader.getPage(1) # Selects the first page.
page.extractText() # Returns the text on that page as a string.

# Loops through all the pages and returns all the text
for pageNum in range(reader.numPages):
    print(reader.getPage(pageNum).extractText())

### How to combine PDF files ###

import PyPDF2

pdf1File = open('meetingminutes.pdf', 'rb')
pdf2File = open('meetingminutes2.pdf', 'rb')

reader1 = PyPDF2.PdfFileReader(pdf1File)
reader2 = PyPDF2.PdfFileReader(pdf2File)

writer = PyPDF2.PdfFileWriter()

for pageNum in range(reader1.numPages):
    page = reader1.getPage(pageNum)
    writer.addPage(page)

for pageNum in range(reader2.numPages):
    page = reader2.getPage(pageNum)
    writer.addPage(page)

outputFile = open('combinedMinutes.pdf', 'wb')
writer.write(outputFile)
outputFile.close()
pdf1File.close()
pdf2File.close()
