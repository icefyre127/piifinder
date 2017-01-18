#!/usr/bin/python
import re
import sys
import zipfile
import os
from pyPdf import PdfFileReader


'''
input1 = PdfFileReader(file("test2.pdf", "rb"))
for pageNum in range(input1.getNumPages()):
    page = input1.getPage(pageNum)
    print page.extractText() + "\n"
'''

class SearchResult:
    def __init__(self,fileName,numMatches,matches):
        self.fileName = fileName
        self.numMatches = numMatches
        self.matches = matches

def processPDF(fileName):
    pdfReader = PdfFileReader(file(fileName, "rb"))
    plainText = []
    for pageNum in range(pdfReader.getNumPages()):
        page = pdfReader.getPage(pageNum)
        plainText += page.extractText() + "\n\n"
    return ''.join(plainText)


def processOfficeXML(fileName):
    with zipfile.ZipFile(fileName) as myfile:
        

        files = myfile.namelist()
        fileExtension = fileName.split(".")[-1]

        for zfile in files:
            plainText = myfile.read(zfile)

        if fileExtension == "docx":
            plainText = re.sub("<[^>]*>","",plainText)


        return plainText


def processFile(fileName):
    with open(fileName) as myfile:
        return myfile.read()

def findMatches(fileName,data,pattern):
    matches = re.findall(pattern,data)
    if len(matches) > 0:
        uniqueMatches = list(set(matches))
        numMatches = len(uniqueMatches)
        return SearchResult(fileName, numMatches, uniqueMatches)
    

    
    
ssnPattern = "\d{3}-\d{2}-\d{4}"


'''

        matches = re.findall(ssnPattern,myfile.read())
        numMatches = len(matches)
        if numMatches > 0:
            print "%s, %s, %s" % (fileName, numMatches, matches)



        ssnList = []                
            matches = re.findall(ssnPattern,data)
            if len(matches) > 0:
                ssnList += matches
        uniqssn = list(set(ssnList))
        numMatches = len(uniqssn)
        if numMatches > 0:
            return "%s, %s, %s" % (fileName, numMatches, uniqssn)
    


        
#cool = processPDF("test2.pdf")
#print cool

output = PdfFileWriter()


# add page 1 from input1 to output document, unchanged
output.addPage(input1.getPage(0))

# add page 2 from input1, but rotated clockwise 90 degrees
output.addPage(input1.getPage(1).rotateClockwise(90))

# add page 3 from input1, rotated the other way:
output.addPage(input1.getPage(2).rotateCounterClockwise(90))
# alt: output.addPage(input1.getPage(2).rotateClockwise(270))

# add page 4 from input1, but first add a watermark from another pdf:
page4 = input1.getPage(3)
watermark = PdfFileReader(file("watermark.pdf", "rb"))
page4.mergePage(watermark.getPage(0))

# add page 5 from input1, but crop it to half size:
page5 = input1.getPage(4)
page5.mediaBox.upperRight = (
page5.mediaBox.getUpperRight_x() / 2,
page5.mediaBox.getUpperRight_y() / 2
)
output.addPage(page5)

# print how many pages input1 has:
print "document1.pdf has %s pages." % input1.getNumPages())

# finally, write "output" to document-output.pdf
outputStream = file("document-output.pdf", "wb")
output.write(outputStream)

'''


'''
myfiles = os.walk("/cygdrive/s/2016")

for item in myfiles:
    for filelist_item in item[2]:
        if "xlsx" in filelist_item:
            print "%s%s%s" % (item[0],os.path.sep,filelist_item)
'''        


fileName = sys.argv[1]
extension = fileName.split(".")[1]
#print "File Name, Number of matches, Matches"

if extension == "xlsx" or extension == "docx":
    with zipfile.ZipFile(fileName) as myfile:

        ssnList = []
        files = myfile.namelist()
  #      print files

        for zfile in files:
 #           print "DATA IN %s\n==============" % zfile
            data = myfile.read(zfile)
 #           print data
            if extension == "docx":
                data = re.sub("<[^>]*>","",data)

            matches = re.findall(ssnPattern,data)
            if len(matches) > 0:
                ssnList += matches
        uniqssn = list(set(ssnList))
        numMatches = len(uniqssn)
        if numMatches > 0:
            print "%s, %s, %s" % (fileName, numMatches, uniqssn)
else:
    with open(sys.argv[1]) as myfile:
        matches = re.findall(ssnPattern,myfile.read())
        numMatches = len(matches)
        if numMatches > 0:
            print "%s, %s, %s" % (fileName, numMatches, matches)


'''
