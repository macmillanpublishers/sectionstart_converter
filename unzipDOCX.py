# this is a copy of the script from our Wordmaker repo

from sys import argv

filename = argv[1]
finaldir = argv[2]
print filename
print finaldir

import os
import zipfile
import shutil
import re
import xml.etree.ElementTree as ET

def unzip_manuscript(self):

    # must be .docx or .docm
    extension = os.path.splitext(self)[1]

    if extension in ('.docx', '.docm', '.doc'):
        print "unzipping %s" % filename
        # get the contents of the Word file
        # filenames = zipfile.namelist(self)
        # print filenames
        document = zipfile.ZipFile(self, 'a')
        print document.namelist(), len(document.namelist())
        document.extractall(finaldir)
        document.close()

        return

unzip_manuscript( filename )
