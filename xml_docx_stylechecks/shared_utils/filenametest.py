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
import dropbox

# if __name__ == '__main__':
# 	# to go up a level to read cfg when invoking from this script (for testing).
#     setupcleanuppath = os.path.join(sys.path[0], '..', 'lib', 'setup_cleanup.py')
#     import imp
#     parentpath = os.path.join(sys.path[0], '..', 'cfg.py')
#     cfg = imp.load_source('cfg', parentpath)
#     setup_cleanup = imp.load_source('os_utils', setupcleanuppath)
# else:
#     import cfg
def getSubmitterViaAPI(inputfile):
    # logger.info("Retrieve submitter info via Dropbox api...")
    submitter_email = ""
    display_name = ""
    dropboxfolder = os.path.join("C:",os.sep,"Users",currentuser,"Dropbox (Macmillan Publishers)")
    __location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
    db_access_token_txt = os.path.join(__location__,'..','..','..',"bookmaker_authkeys","access_token.txt")
    with open(db_access_token_txt) as f:
        db_access_token = f.readline()
    try:
        # dropbox api requires forward slash in path, and is a relative path (in relation to Dropbox folder)
        dropbox_relpath = inputfile.replace(dropboxfolder,"").replace("\\","/")
        dbx = dropbox.Dropbox(db_access_token)
        submitter = (dbx.files_get_metadata(dropbox_relpath).sharing_info.modified_by)
        display_name = dbx.users_get_account(submitter).name.display_name
        submitter_email = dbx.users_get_account(submitter).email
    except Exception, e:
        print e
    finally:
        return submitter_email, display_name


# from lxml import etree
badchar = sys.argv[1]
inputfile = sys.argv[2]
inputfilename = os.path.basename(inputfile)
inputfilename_noext = os.path.splitext(inputfilename)[0]
inputfilenq = inputfile
if inputfile[0] == '"':
    inputfilenq = inputfile[1:]
if inputfilenq[-1:] == '"':
    inputfilenq = inputfilenq[:-1]

currentuser = getpass.getuser()
# filename = "S:\Users\padwoadmin\Documents\programming_projects\1708_2_python_ssconvertertests\bugsrug_%s.txt" % time.strftime("%m%d-%H%M%S")
filename = os.path.join(os.sep,"Users",currentuser,"Documents","programming_projects","1708_2_python_ssconvertertests","bugsrug_.txt")
# filename = os.path.join(os.sep,"Users",currentuser,"Documents","programming_projects","1708_2_python_ssconvertertests","bugsrug_%s.txt"% time.strftime("%m%d-%H%M%S"))
try:
    print "version / badchar: ",badchar
    print "inputfilename_noext : ", inputfilename_noext
    print "inputfilename : ", inputfilename
    print "inputfile : ", inputfile
    print getSubmitterViaAPI(inputfilenq)
    # print "inputfilenamenoext encode : ", inputfilename_noext.decode('hex')#.encode("ascii")

    newfile = open(filename, 'a')
    with newfile as f:
        # for line in text:
        #     f.write("%s\n" % line)
        f.write("version / badchar: %s\n" % badchar)
        f.write("inputfilename_noext : %s\n" % inputfilename_noext)
        f.write("inputfilename : %s\n" % inputfilename)
        f.write("inputfile : %s\n" % inputfile)
        # f.write("inputfilenamenoext encode : %s\n" % inputfilename_noext.decode("cp1250"))#.encode('utf8'))
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
