######### IMPORT SOME STANDARD PY LIBRARIES
import sys
import os
import zipfile
import shutil
import re
import logging
import time


######### IMPORT LOCAL MODULES
import cfg
import lib.addsectionstarts as addsectionstarts
import lib.stylereports as stylereports
import lib.generate_report as generate_report
import shared_utils.unzipDOCX as unzipDOCX
import shared_utils.os_utils as os_utils
import shared_utils.check_docx as check_docx


######### LOCAL DECLARATIONS
inputfile = cfg.inputfile
inputfilename_noext = cfg.inputfilename_noext
workingfile = cfg.workingfile
ziproot = cfg.ziproot
this_outfolder = cfg.this_outfolder
# newdocxfile = cfg.newdocxfile
stylereport_txt = cfg.stylereport_txt
report_dict = {}
template_ziproot = cfg.template_ziproot
macmillan_template = cfg.macmillan_template


######### SETUP LOGGING
logfile = os.path.join(cfg.logdir, "reporter_{}_{}.txt".format(inputfilename_noext, time.strftime("%y%m%d-%H%M%S")))
cfg.defineLogger(logfile, cfg.loglevel)
logger = logging.getLogger(__name__)


######### RUN!
# only run if this script is being invoked directly
if __name__ == '__main__':

    # create & cleanup tmpfolder, outfolder if they do not exist:
    tmpdir = os_utils.setupTmpfolder(cfg.tmpdir)
    logger.info("Create & cleanup project outfolder")
    os_utils.setupOutfolder(this_outfolder)

    logger.info('Moving input file ({}) and template to tmpdir and unzipping'.format(inputfilename_noext))

    # move file to the tmpdir
    # os_utils.movefile()            # for production
    os_utils.copyFiletoFile(inputfile, workingfile)        # debug/testing only

    # move template to the tmpdir
    os_utils.copyFiletoFile(macmillan_template, os.path.join(tmpdir, os.path.basename(macmillan_template)))

    ### unzip the manuscript to ziproot, template to template_ziproot
    os_utils.rm_existing_os_object(ziproot, 'ziproot')
    os_utils.rm_existing_os_object(ziproot, 'template_ziproot')
    unzipDOCX.unzipDOCX(workingfile, ziproot)
    unzipDOCX.unzipDOCX(macmillan_template, template_ziproot)

    ### check and compare versions, styling percentage, doc protection
    logger.info('Comparing docx version to template, checking percent styled, checking if protected doc...')
    version_result, current_version, template_version = check_docx.version_test(cfg.customprops_xml, cfg.template_customprops_xml, cfg.sectionstart_versionstring)
    percent_styled, macmillan_styled_paras, total_paras = check_docx.macmillanStyleCount(cfg.doc_xml, cfg.styles_xml)
    protection = check_docx.checkSettingsXML(cfg.settings_xml, "documentProtection")

    if version_result == "up_to_date" and percent_styled >= 50 and protection == False:
        logger.info("Proceeding! (version='%s', percent_styled='%s', protection='%s')" % (version_result, percent_styled, protection))

        # # # check section starts!
        logger.info("Checking section starts")
        report_dict = addsectionstarts.sectionStartCheck("report", report_dict)

        # # # run otherstyle report stuff!
        logger.info("Running other style report functions")
        report_dict = stylereports.styleReports(report_dict)

        # write our stylereport.json with all edits etc for
        logger.debug("Writing stylereport.json")
        os_utils.dumpJSON(report_dict, cfg.stylereport_json)

        # write our stylereport.txt
        logger.debug("Writing stylereport.txt to outfolder")
        generate_report.generateReport(report_dict, stylereport_txt)

    else:
        logger.warn("* * Skipping Style Report:")
        if percent_styled < 50:
            logger.warn("* This .docx has {} percent of paragraphs styled with Macmillan styles".format(percent_styled))
        if version_result != "up_to_date":
            logger.warn("* You must attach the newest version of the macmillan style template before running the Style Report: (this .docx's version: {}, template version: {})".format(current_version, template_version))
        if protection == True:
            logger.warn("* This .docx has protection enabled.")

    # here I guess we call some piece of the reporter_main?
    # to generate a style_report.txt from this json?
    # and write it to the outfolder
    # and maybe send an email?

    # this is where we would also capture the above errors form a structured error.json and output them
    # as part of stylereport, or separate

    ##### CLEANUP
    # Return original file to user
    logger.info("Copying original file to outfolder/original_file dir")
    if not os.path.isdir(os.path.join(this_outfolder, "original_file")):
        os.makedirs(os.path.join(this_outfolder, "original_file"))
    os_utils.copyFiletoFile(workingfile, os.path.join(this_outfolder, "original_file", cfg.inputfilename))

    # write stylereport to outfolder, + any separate warn_notices, if we didn't do it above

    # Rm tmpdir
    logger.debug("deleting tmp folder")
    # os_utils.rm_existing_os_object(tmpdir, 'tmpdir')        # comment out for testing / debug
