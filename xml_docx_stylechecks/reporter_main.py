######### IMPORT SOME STANDARD PY LIBRARIES
import sys
import os
import zipfile
import shutil
import re
import logging
import time
import inspect


######### IMPORT LOCAL MODULES
import cfg
import lib.addsectionstarts as addsectionstarts
import lib.stylereports as stylereports
import lib.generate_report as generate_report
import lib.setup_cleanup as setup_cleanup
import shared_utils.os_utils as os_utils
import shared_utils.check_docx as check_docx


######### LOCAL DECLARATIONS
inputfile = cfg.inputfile
inputfilename_noext = cfg.inputfilename_noext
workingfile = cfg.workingfile
ziproot = cfg.ziproot
this_outfolder = cfg.this_outfolder
tmpdir = cfg.tmpdir
report_dict = {}
template_ziproot = cfg.template_ziproot
macmillan_template = cfg.macmillan_template


######### SETUP LOGGING
logfile = os.path.join(cfg.logdir, "{}_{}_{}.txt".format(cfg.script_name, inputfilename_noext, time.strftime("%y%m%d-%H%M%S")))
cfg.defineLogger(logfile, cfg.loglevel)
logger = logging.getLogger(__name__)


######### RUN!
# only run if this script is being invoked directly
if __name__ == '__main__':
    try:
        ########## SETUP
        # setup & create tmpdir, cleanup outfolder (archive existing), copy infile to tmpdir
        tmpdir, workingfile = setup_cleanup.setupFolders(tmpdir, inputfile, cfg.inputfilename, this_outfolder)

        # copy template to tmpdir, unzip infile and tmpdir
        setup_cleanup.copyTemplateandUnzipFiles(macmillan_template, tmpdir, workingfile, ziproot, template_ziproot)

        ########## CHECK DOCUMENT
        ### check and compare versions, styling percentage, doc protection
        logger.info('Comparing docx version to template, checking percent styled, checking if protected doc...')
        version_result, current_version, template_version = check_docx.version_test(cfg.customprops_xml, cfg.template_customprops_xml, cfg.sectionstart_versionstring)
        percent_styled, macmillan_styled_paras, total_paras = check_docx.macmillanStyleCount(cfg.doc_xml, cfg.styles_xml)
        protection = check_docx.checkSettingsXML(cfg.settings_xml, "documentProtection")

        ########## RUN STUFF
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

        ########## SKIP RUNNING STUFF, LOG ALERTS
        else:
            logger.warn("* * Skipping Style Report:")
            if percent_styled < 50:
                errstring = "This .docx has {} percent of paragraphs styled with Macmillan styles".format(percent_styled)
                os_utils.logAlerttoJSON(cfg.alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))
            if version_result != "up_to_date":
                errstring = "You must attach the newest version of the macmillan style template before running the Style Report: (this .docx's version: {}, template version: {})".format(current_version, template_version)
                os_utils.logAlerttoJSON(cfg.alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))
            if protection == True:
                errstring = "* This .docx has protection enabled."
                os_utils.logAlerttoJSON(cfg.alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))

        ########## CLEANUP
        setup_cleanup.cleanupforReporterOrConverter(this_outfolder, workingfile, cfg.inputfilename, report_dict, cfg.stylereport_txt, cfg.alerts_json, tmpdir)

    except:
        ########## LOG ERROR INFO
        # log to logfile for dev
        logger.exception("ERROR ------------------ :")
        # log to errfile_json for user
        invokedby_script = os.path.splitext(os.path.basename(inspect.stack()[0][1]))[0]
        os_utils.logAlerttoJSON(cfg.alerts_json, "error", "A fatal error was encountered while running '%s'.\n\nPlease email workflows@macmillan.com for assistance." % invokedby_script)

        setup_cleanup.cleanupException(this_outfolder, workingfile, cfg.inputfilename, cfg.alerts_json, tmpdir, cfg.logdir, inputfilename_noext, cfg.script_name)
