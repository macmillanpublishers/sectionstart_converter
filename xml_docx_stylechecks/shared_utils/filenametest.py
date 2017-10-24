######### IMPORT SOME STANDARD PY LIBRARIES

import sys
import os
import shutil
import re
# import codecs
# import json
# import logging
import time
import getpass
# from lxml import etree
badchar = sys.argv[1]
inputfile = sys.argv[2]
inputfilename = os.path.basename(inputfile)
inputfilename_noext = os.path.splitext(inputfilename)[0]

currentuser = getpass.getuser()
# filename = "S:\Users\padwoadmin\Documents\programming_projects\1708_2_python_ssconvertertests\bugsrug_%s.txt" % time.strftime("%m%d-%H%M%S")
filename = os.path.join(os.sep,"Users",currentuser,"Documents","programming_projects","1708_2_python_ssconvertertests","bugsrug_.txt")
# filename = os.path.join(os.sep,"Users",currentuser,"Documents","programming_projects","1708_2_python_ssconvertertests","bugsrug_%s.txt"% time.strftime("%m%d-%H%M%S"))
try:
    print "version / badchar: ",badchar
    print "inputfilename_noext : ", inputfilename_noext
    print "inputfilename : ", inputfilename
    print "inputfile : ", inputfile
    print "inputfilenamenoext encode : ", inputfilename_noext.decode('hex')#.encode("ascii")

    newfile = open(filename, 'a')
    with newfile as f:
        # for line in text:
        #     f.write("%s\n" % line)
        f.write("version / badchar: %s\n" % badchar)
        f.write("inputfilename_noext : %s\n" % inputfilename_noext)
        f.write("inputfilename : %s\n" % inputfilename)
        f.write("inputfile : %s\n" % inputfile)
        f.write("inputfilenamenoext encode : %s\n" % inputfilename_noext.decode("cp1250"))#.encode('utf8'))
        f.write("____________________\n")
            # f.write(line)
        f.close()
except Exception, e:
    print e
    newfile = open(filename, 'a')
    with newfile as f:
        # for line in text:
        #     f.write("%s\n" % line)

        f.write("%s\n" % e)
        # f.write("inputfilename_noext : %s\n" % inputfilename_noext)
        # f.write("inputfilename : %s\n" % inputfilename)
        # f.write("inputfile : %s\n" % inputfile)
        # f.write("inputfilenamenoext encode : %s\n" % inputfilename_noext.encode('utf8'))
        # f.write("____________________\n")
        #     # f.write(line)
        f.close()
