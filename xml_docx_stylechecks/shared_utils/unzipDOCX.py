import os
import zipfile
import shutil
import re
import logging
import xml.etree.ElementTree as ET
from sys import argv


# # # # # initialize logger
logger = logging.getLogger(__name__)


# # # # # METHODS
def unzipDOCX(filename, finaldir):
    try:
        logger.debug("unzipping '%s' to '%s'" % (filename, finaldir))
        # must be .docx or .docm
        extension = os.path.splitext(filename)[1]

        if extension in ('.docx', '.docm', '.doc', '.dotx', '.dotm'):
            print "unzipping %s" % filename
            # get the contents of the Word file
            # filenames = zipfile.namelist(filename)
            # print filenames
            document = zipfile.ZipFile(filename, 'a')
            print document.namelist(), len(document.namelist())
            document.extractall(finaldir)
            document.close()
            return
        else:
            logger.error("Could not unzip %s, not a Word doctype" % filename)   
            raise
            sys.exit(1)     
    except Exception, e:    
        logger.error('Failed to unzip .doc', exc_info=True)
        sys.exit(1)


# # # # # RUN
# for running this script as a standalone:
if __name__ == '__main__':
    # handle args as a standalone
    filename = argv[1]
    finaldir = argv[2]

    # set up debug log to console
    logging.basicConfig(level=logging.DEBUG)

    unzipDOCX(filename, finaldir)
