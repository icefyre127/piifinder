#!/usr/bin/python


import re
import sys
import zipfile
ssnPattern = "\d{3}-\d{2}-\d{4}"

fileName = sys.argv[1]
if sys.argv[1].split(".")[1] == "xlsx":
    with zipfile.ZipFile(fileName) as myfile:

        files = myfile.namelist()

        for zfile in files:
            data = myfile.read(zfile)
            matches = re.findall(ssnPattern,data)
            if len(matches) > 0:
                print "%s: %s" % (fileName, matches)
            
#        for data in zipfile.ZipFile.read(myfile,"r"):
 #           print data
        #for archiveFile in  myfile.namelist():
#         matches = re.findall(ssnPattern,zipfile.ZipFile.open(archiveFile))
         #   print matches
else:
    with open(sys.argv[1]) as myfile:
        matches = re.findall(ssnPattern,myfile.read())
        print matches
