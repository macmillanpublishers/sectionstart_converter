import os
import sys
import textwrap
import logging
import time

# #	# # # # # # ARGS
batchfile = sys.argv[0]
processwatch_file = sys.argv[1]
processwatch_seconds = sys.argv[2]
scriptname = sys.argv[3]
filename = sys.argv[4]

# ######### IMPORT LOCAL MODULES
import sendmail

######### LOCAL DECLARATIONS
# the path of this file: setting '__location__' allows this relative path to adhere to this file, even when invoked from a different path:
# 	https://stackoverflow.com/questions/4060221/how-to-reliably-open-a-file-in-the-same-directory-as-a-python-script
__location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
smtp_txt = os.path.join(__location__,'..','..','..',"bookmaker_authkeys","smtp.txt")
with open(smtp_txt) as f:
    smtp_address = f.readline()


# # #---------------------  METHODS
def checkPwfile(processwatch_file):
    if os.path.exists(processwatch_file):
        return True
    else:
        return False

def sendAlertmail(subject, bodytxt):
    sendmail.sendMailBasic(25, smtp_address, "workflows@macmillan.com", ["workflows@macmillan.com"], subject, bodytxt, None, None)


def watchProcess(batchfile, processwatch_file, processwatch_seconds, scriptname, filename):
    # logging.debug("I made it here") # debug
    subject = "process watch Alert: %s" % scriptname
    if checkPwfile(processwatch_file) == False:
        # set text
        bodytxt = textwrap.dedent("""
            Hello,

            Script '{batchfile}' was invoked to run '{scriptname}' on file:
                '{filename}'

            However processwatch_file '{processwatch_file}' was not created before this process watcher was invoked, indicating:
                - a problem with the batch file or
                - {scriptname} is deleting the file before this file checks it, which would be pretty weird.

            Go investigate, Sherlock!
        """).format(batchfile=batchfile, scriptname=scriptname, filename=filename, processwatch_file=processwatch_file, processwatch_seconds=processwatch_seconds)
        # sendmail
        sendAlertmail(subject, bodytxt)
        return
    time.sleep(int(processwatch_seconds))
    if checkPwfile(processwatch_file) == True:
        # set text
        bodytxt = textwrap.dedent("""
            Hello,

            Script '{batchfile}' was invoked to run '{scriptname}' on file:
                '{filename}'

            However processwatch_file '{processwatch_file}' still exists, after waiting a full {processwatch_seconds} seconds, indicating:
                - something hung up somewhere, or
                - a cleanup AND exception_cleanup scripts failed.

            Go investigate Sherlock! (try the logfiles :)
        """).format(batchfile=batchfile, scriptname=scriptname, filename=filename, processwatch_file=processwatch_file, processwatch_seconds=processwatch_seconds)
        # sendmail
        sendAlertmail(subject, bodytxt)
        return



#---------------------  MAIN
# only run if this script is being invoked directly
if __name__ == '__main__':
#     # set up debug log to file
    try:
        logfilename = "PWerr_{}_{}.txt".format(os.path.splitext(scriptname)[0], time.strftime("%y%m%d-%H%M%S"))
        logdir = os.path.dirname(processwatch_file)
        logfile = os.path.join(logdir,logfilename)
        # logdir = '/Users/username/test' # debug
        # logfile = os.path.join("S:",os.sep,"resources","logs","processLogs","test.txt") #debug

        # # # # debug block to check parameters
        # logging.basicConfig(filename=logfile,level=logging.DEBUG)   # debug
        # logging.debug("0: %s" % sys.argv[0])
        # logging.debug("1: %s" % sys.argv[1])
        # logging.debug("2: %s" % sys.argv[2])
        # logging.debug("3: %s" % sys.argv[3])
        # logging.debug("4: %s" % sys.argv[4])

        watchProcess(batchfile, processwatch_file, processwatch_seconds, scriptname, filename)

    except:
        logging.basicConfig(filename=logfile,level=logging.DEBUG)
        errmessage = 'Process_watcher err; pw_file is: {}'.format(processwatch_file)
        logging.error(errmessage, exc_info=True)
        # logging.error("test msg", exc_info=True)   # debug
        sys.exit(1)
