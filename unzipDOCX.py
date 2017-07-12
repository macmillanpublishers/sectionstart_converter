from sys import argv

if __name__ == '__main__':
    filename = argv[1]
    finaldir = argv[2]
    print filename
    print finaldir

import os
import zipfile
import shutil
import re
import xml.etree.ElementTree as ET

def unzip_manuscript(self, finaldir):

    # must be .docx or .docm
    extension = os.path.splitext(self)[1]

    if extension in ('.docx', '.docm', '.doc', '.dotx', '.dotm'):
        print "unzipping %s" % self
        # get the contents of the Word file
        # filenames = zipfile.namelist(self)
        # print filenames
        document = zipfile.ZipFile(self, 'a')
        print document.namelist(), len(document.namelist())
        document.extractall(finaldir)
        document.close()

        return

if __name__ == '__main__':
    unzip_manuscript(filename, finaldir)
