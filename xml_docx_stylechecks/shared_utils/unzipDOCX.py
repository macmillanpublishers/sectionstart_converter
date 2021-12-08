import os
import zipfile
import shutil
import re
import logging
import xml.etree.ElementTree as ET
import sys


# # # # # initialize logger
logger = logging.getLogger(__name__)


# # # # # METHODS
def unzipDOCX(filename, finaldir, err_dict={}):
    errmsg=''
    try:
        logger.debug("unzipping '%s' to '%s'" % (filename, finaldir))
        # must be .docx or .docm
        extension = os.path.splitext(filename)[1]
        if extension in ('.docx', '.docm', '.doc', '.dotx', '.dotm'):
            document = zipfile.ZipFile(filename, 'a')
            # print document.namelist(), len(document.namelist()) # debug
            if document.namelist():
                document.extractall(finaldir)
                document.close()
            else:
                errmsg = 'cannot unzip; no document namelist'
                # allow errmsg override from calling function
                if err_dict and err_dict['no_filelist']:
                    errmsg = err_dict['no_filelist']
                logger.error('{}: filename "{}"'.format(errmsg, filename))
                raise
            return
        else:
            errmsg = 'cannot unzip; not a Word doctype'
            if err_dict and err_dict['not_docfile']:
                errmsg = err_dict['not_docfile']
            logger.error('{}: filename "{}"'.format(errmsg, filename))
            raise
    except:
        if not errmsg:
            errmsg = 'unexpected exception during unzip'
            logger.error('{}: filename "{}", dest_dir: "{}"'.format(errmsg, filename, finaldir), exc_info=True)
        raise Exception(errmsg)
        sys.exit(1)


# # # # # RUN
# for running this script as a standalone:
if __name__ == '__main__':
    # handle args as a standalone
    filename = sys.argv[1]
    finaldir = sys.argv[2]

    # set up debug log to console
    logging.basicConfig(level=logging.DEBUG)

    unzipDOCX(filename, finaldir)
