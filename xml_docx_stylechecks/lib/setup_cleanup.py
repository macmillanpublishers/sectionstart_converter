######### IMPORT PY LIBRARIES
import os
import shutil
import re
import uuid
import json
import sys
import collections
import logging


######### IMPORT LOCAL MODULES
if __name__ == '__main__':
    # to go up a level to read cfg (and other files) when invoking from this script (for testing).
    cfgpath = os.path.join(sys.path[0], '..', 'cfg.py')
    osutilspath = os.path.join(sys.path[0], '..', 'shared_utils', 'os_utils.py')
    import imp
    cfg = imp.load_source('cfg', cfgpath)
    os_utils = imp.load_source('os_utils', osutilspath)
    unzipDOCX = imp.load_source('unzipDOCX', osutilspath)
    import generate_report
else:
    import cfg
    import shared_utils.os_utils as os_utils
    import shared_utils.unzipDOCX as unzipDOCX
    import lib.generate_report as generate_report


######### LOCAL DECLARATIONS

# initialize logger
logger = logging.getLogger(__name__)


#---------------------  METHODS
def setupFolders(tmpdir, inputfile, inputfilename, this_outfolder):
    logger.info("Create tmpdir, create & cleanup project outfolder")

    # create new tmpdir, reset value for working file
    tmpdir = os_utils.setupTmpfolder(tmpdir)
    workingfile = os.path.join(tmpdir, inputfilename)
    os_utils.setupOutfolder(this_outfolder)

    # move inputfile to tmpdir as workingfile
    logger.info('Moving input file ({}) and template to tmpdir'.format(inputfilename))
    # os_utils.movefile(inputfile, workingfile)			# for production
    os_utils.copyFiletoFile(inputfile, workingfile)		# debug/testing only

    return tmpdir, workingfile

def copyTemplateandUnzipFiles(macmillan_template, tmpdir, workingfile, ziproot, template_ziproot):
    # move template to the tmpdir
    os_utils.copyFiletoFile(macmillan_template, os.path.join(tmpdir, os.path.basename(macmillan_template)))

    ### unzip the manuscript to ziproot, template to template_ziproot
    os_utils.rm_existing_os_object(ziproot, 'ziproot')
    os_utils.rm_existing_os_object(ziproot, 'template_ziproot')
    unzipDOCX.unzipDOCX(workingfile, ziproot)
    unzipDOCX.unzipDOCX(macmillan_template, template_ziproot)

def returnOriginal(this_outfolder, workingfile, inputfilename):
    # Return original file to user
    logger.info("Copying original file to outfolder/original_file dir")
    if not os.path.isdir(os.path.join(this_outfolder, "original_file")):
        os.makedirs(os.path.join(this_outfolder, "original_file"))
    os_utils.copyFiletoFile(workingfile, os.path.join(this_outfolder, "original_file", inputfilename))

def cleanupforReporterOrConverter(this_outfolder, workingfile, inputfilename, report_dict, stylereport_txt, alerts_json, tmpdir):
    logger.info("Running cleanup, 'cleanupforReporterOrConverter'...")

    # 1 return original_file to outfolder
    returnOriginal(this_outfolder, workingfile, inputfilename)

    # 2 if report_dict has contents, write stylereport file:
    logger.debug("Writing stylereport.txt to outfolder")
    if report_dict:
        generate_report.generateReport(report_dict, stylereport_txt)
        # and send stylereport as mail

    # 3 write our alertfile.txt if necessary
    logger.debug("Writing alerts.txt to outfolder")
    os_utils.writeAlertstoTxtfile(alerts_json, this_outfolder)

    # 4 Rm tmpdir
    logger.debug("deleting tmp folder")
    # os_utils.rm_existing_os_object(tmpdir, 'tmpdir')		# comment out for testing / debug

def cleanupforValidator(this_outfolder, workingfile, inputfilename, report_dict, stylereport_txt, alerts_json):
    logger.info("Running cleanup, 'cleanupforValidator'...")

    # 1 if report_dict has contents, write stylereport file:
    logger.debug("Writing stylereport.txt to outfolder")
    if report_dict:
        generate_report.generateReport(report_dict, stylereport_txt)
        # and send stylereport as mail

    # 2 write our alertfile.txt if necessary
    logger.debug("Writing alerts.txt to outfolder")
    os_utils.writeAlertstoTxtfile(alerts_json, this_outfolder)

def cleanupException(this_outfolder, workingfile, inputfilename, alerts_json, tmpdir, logdir, inputfilename_noext, scriptname):
    logger.warn("POST-ERROR: Running cleanup, 'cleanupException'...")

    # 1 send email to workflows

    # 2 write alertfile
    logger.info("Writing alerts.txt to outfolder")
    try:
        os_utils.writeAlertstoTxtfile(alerts_json, this_outfolder)
    except:
        logger.exception("ERROR with exception cleanup :(")

    # 3 save a copy of tmpdir to logdir for troubleshooting (since it will be deleted)
    logger.info("Backing up tmpdir to logfolder")
    try:
        os_utils.copyDir(tmpdir, os.path.join(logdir, "tmpdir_%s" % inputfilename_noext))
    except:
        logger.exception("ERROR with exception cleanup :(")

    # these two items only apply to converter and reporter
    if scriptname != "validator":
        # 4 return original_file to outfolder
        try:
            returnOriginal(this_outfolder, workingfile, inputfilename)
        except:
            logger.exception("ERROR with exception cleanup :(")

        # 5 Rm tmpdir to avoid interfering with next run
        logger.debug("deleting tmp folder")
        try:
            os_utils.rm_existing_os_object(tmpdir, 'tmpdir')		# comment out for testing / debug
        except:
            logger.exception("ERROR with exception cleanup :(")


# # only run if this script is being invoked directly
# if __name__ == '__main__':
