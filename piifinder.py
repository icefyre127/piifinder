#!/usr/bin/python
import argparse
import re
import sys
import zipfile
import os
from pyPdf import PdfFileReader
import time

class SearchResult:
    def __init__(self,fileName,numMatches,matches):
        self.fileName = fileName
        self.numMatches = numMatches
        self.matches = matches

def processPDF(fileName,verbosity):
    print "Processing: " ,fileName
    try:
        pdfReader = PdfFileReader(file(fileName, "rb"))
        plainText = []
        startTime = time.time()
        for pageNum in range(pdfReader.getNumPages()):
            page = pdfReader.getPage(pageNum)
            #print page.extractText() + "\n\n"

            if args.verbosity > 0:
                print "\t",fileName,": page [",pageNum,"/",pdfReader.getNumPages(),"]"
            plainText += page.extractText() + "\n\n"
            timeElapsed = time.time()
            if timeElapsed - startTime > 10:
                print "\tFile taking too long to process. Timing out."
        #        print plainText
                return ''.join(plainText).encode("utf-8")
            
        return ''.join(plainText).encode("utf-8")
    except Exception as e:
        print fileName, ": READ ERROR. Error: ", e.args
        return "ERROR"



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

def findQuickMatches(filename,data,pattern):
    return re.search(pattern,data)

    
    
    
ssnPattern = "\d{3}-\d{2}-\d{4}"

parser = argparse.ArgumentParser(description='Search for specific data patterns in file(s). Also supports PDF,docx and xlsx files')
parser.add_argument('-R','--recursive',action='store_true', help='Include all files in directory and in subdirectories')
parser.add_argument('-q','--quick',action='store_true', help='Use quick matching, just true or false without returning matched patterns')
parser.add_argument('location',nargs='+', help='Location (e.g. file, directory) where to search for data')
parser.add_argument('-v','--verbosity', action="count",help="increase output verbosity")
args = parser.parse_args()


fileList = []
searchFile = ""
start = time.time()
for item in args.location:
    if len(item)>1 and item[-1] == os.path.sep:
        item = item[0:-1]
    
    if os.path.isdir(item) and not args.recursive:
        for fileItem in os.listdir(item):
            if not os.path.isdir(item+os.path.sep+fileItem):
                fileList.append(item+os.path.sep+fileItem)

    
    elif os.path.isdir(item):
       for walkItem in os.walk(item):
            for fileItem in walkItem[2]:
                if walkItem[0][-1] == os.path.sep:
                    fileList.append(walkItem[0][0:-1] + os.path.sep + fileItem)
                else:
                    fileList.append(walkItem[0]+os.path.sep+fileItem)
    else:
        fileList.append(item)

        

'''
print "FILES TO SEARCH: "

for f in  fileList:
    print "\t", f
'''   
        
'''
        for walkItem in os.walk(item):
            for fileItem in walkItem[2]:
                print "\t%s%s%s" % (walkItem[0],os.path.sep,fileItem)
'''




'''
myfiles = os.walk("/cygdrive/s/2016")

for item in myfiles:
    for filelist_item in item[2]:
        if "xlsx" in filelist_item:
            print "%s%s%s" % (item[0],os.path.sep,filelist_item)
'''        

def printReport(results):
    if (results):
        print "File Name: %s  Number of Matches: %s  Matches found: %s " % (results.fileName, results.numMatches,results.matches)



excludedTypes = ["exe","zip","7z","jpeg","jpg","gif","bmp","iso","bin","ini","lnk"]
        
for fileName in fileList:
    extension = fileName.split(".")[-1].lower()

    if extension in excludedTypes:
        continue

#    print "File being processed: ", fileName
    if extension == "xlsx" or extension == "docx":
        fileContents = processOfficeXML(fileName)
    
    elif extension == "pdf":
        fileContents = processPDF(fileName,args.verbosity)
    else:
        fileContents = processFile(fileName)
        
    if fileContents == "ERROR":
        continue
    
    if args.quick:
        results = findQuickMatches(fileName, fileContents,ssnPattern)
    else:
        results = findMatches(fileName, fileContents,ssnPattern)
    
    if not results:
        print fileName,": No matches found."
    else:        
        if args.quick:
            print fileName,": match found."
        else:
            printReport(results)

end = time.time()

print (end - start)
