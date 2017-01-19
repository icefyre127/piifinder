#!/usr/bin/python
import re
import sys
import zipfile
import os
from pyPdf import PdfFileReader


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

    
    return ''.join(plainText).encode("ascii")


def processOfficeXML(fileName):
    with zipfile.ZipFile(fileName) as myfile:
        files = myfile.namelist()

        fileExtension = fileName.split(".")[-1].lower()
        plainText = []
        for zfile in files:
            plainText += myfile.read(zfile)

        plainText = ''.join(plainText)

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
    return False

    
    
ssnPattern = "\d{3}-\d{2}-\d{4}"


'''
myfiles = os.walk("/cygdrive/s/2016")

for item in myfiles:
    for filelist_item in item[2]:
        if "xlsx" in filelist_item:
            print "%s%s%s" % (item[0],os.path.sep,filelist_item)
'''        

def printReport(results):
    if (results):
        print "File Name: %s\nNumber of Matches: %s\nMatches found: %s\n" % (results.fileName, results.numMatches,results.matches)


fileName = sys.argv[1]
extension = fileName.split(".")[-1].lower()

if extension == "xlsx" or extension == "docx":
    fileContents = processOfficeXML(fileName)
    
elif extension == "pdf":
    fileContents = processPDF(fileName)
else:
    fileContents = processFile(fileName)


    
results = findMatches(fileName, fileContents,ssnPattern)

if not results:
    print "No matches found"
else:
    printReport(results)
