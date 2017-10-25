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
newhome = os.path.join("S:",os.sep,"Users","padwoadmin","Documents","programming_projects","1708_2_python_ssconvertertests","newhome")
copydest = os.path.join("S:",os.sep,"Users","padwoadmin","Documents","programming_projects","1708_2_python_ssconvertertests","copydest")
def moveFile(pathtofile, dest):
    # Create dest directory if it does not exist (guessing whether we have a file or dir as dest, based on period in name)
    if "." not in dest:
        if not os.path.isdir(dest):
            os.makedirs(dest)
    else:
        if not os.path.isdir(os.path.dirname(dest)):
            os.makedirs(os.path.dirname(dest))
    # Move the file
    # try:
    shutil.move(pathtofile, dest)
    # except Exception, e:
    #     logger.error('Failed to move file, exiting', exc_info=True)
    #     sys.exit(1)

def copyFiletoFile(pathtofile, dest_file):
    if not os.path.isdir(os.path.dirname(dest_file)):
        os.makedirs(os.path.dirname(dest_file))
    # try:
    shutil.copyfile(pathtofile, dest_file)
    # except Exception, e:
    #     logger.error('Failed copyfile, exiting', exc_info=True)
    #     sys.exit(1)

def getSubmitterViaAPI(inputfile):
    # logger.info("Retrieve submitter info via Dropbox api...")
    submitter_email = ""
    display_name = ""
    dropboxfolder = os.path.join("C:",os.sep,"Users",currentuser,"Dropbox (Macmillan Publishers)")
    __location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
    db_access_token_txt = os.path.join(__location__,'..','..','..',"bookmaker_authkeys","access_token.txt")
    with open(db_access_token_txt) as f:
        db_access_token = f.readline()
    # try:
    # dropbox api requires forward slash in path, and is a relative path (in relation to Dropbox folder)
    # the decode cp1252 is to unencode unicode chars that were encoded by the batch file, for the api.
    dropbox_relpath = inputfile.replace(dropboxfolder,"").replace("\\","/").decode("cp1252")
    dbx = dropbox.Dropbox(db_access_token)
    submitter = (dbx.files_get_metadata(dropbox_relpath).sharing_info.modified_by)
    display_name = dbx.users_get_account(submitter).name.display_name
    submitter_email = dbx.users_get_account(submitter).email
# except Exception, e:
#     print e#.decode("utf-8")
#     newfile = open(filename, 'a')
#     with newfile as f:
#         f.write("{}\n".format(e))
#         f.close()
# finally:
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
filename = os.path.join(os.sep,"Users",currentuser,"Documents","programming_projects","1708_2_python_ssconvertertests","bugsrug_.txt")
# try:
submitterA, submitterB = getSubmitterViaAPI(inputfilenq)
# print "version / badchar: ",badchar
print "inputfilename_noext : ", inputfilename_noext
print "inputfilename : ", inputfilename
print "inputfile : ", inputfile
print submitterA, submitterB
# print "inputfilenamenoext encode : ", inputfilename_noext.decode('hex')#.encode("ascii")

newfile = open(filename, 'a')
with newfile as f:
    # for line in text:
    f.write("version / badchar: {}\n".format(badchar))
    f.write("inputfilename_noext : {}\n".format(inputfilename_noext))
    f.write("inputfilename : {}\n".format(inputfilename))
    f.write("inputfile : {}\n".format(inputfile))
    f.write("{}, {}\n".format(submitterA, submitterB))
    f.write("____________________\n")
        # f.write(line)
    f.close()
# except Exception, e:
#     print e#.decode("utf-8")
#     newfile = open(filename, 'a')
#     with newfile as f:
#         f.write("{}\n".format(e))
#         f.close()
sanitizedname = re.sub('\W','',inputfilename_noext)
copyFiletoFile(inputfilenq, os.path.join(copydest,"%s.docx" %sanitizedname))

moveFile(inputfilenq, newhome)
