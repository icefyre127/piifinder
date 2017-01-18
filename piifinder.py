#!/usr/bin/python
import re
import sys
import zipfile
import os
ssnPattern = "\d{3}-\d{2}-\d{4}"

'''
myfiles = os.walk("/cygdrive/s/2016")

for item in myfiles:
    for filelist_item in item[2]:
        if "xlsx" in filelist_item:
            print "%s%s%s" % (item[0],os.path.sep,filelist_item)
'''        

   
fileName = sys.argv[1]
if sys.argv[1].split(".")[1] == "xlsx" or sys.argv[1].split(".")[1] == "docx":
    with zipfile.ZipFile(fileName) as myfile:

        ssnList = []
        files = myfile.namelist()
        print files
        for zfile in files:
            data = myfile.read(zfile)
            matches = re.findall(ssnPattern,data)
            if len(matches) > 0:
                ssnList += matches
        uniqssn = list(set(ssnList))
        print "%s: %s" % (fileName, uniqssn)
        print "number of matches =  %d" % len(uniqssn)
else:
    with open(sys.argv[1]) as myfile:
        matches = re.findall(ssnPattern,myfile.read())
        print matches

