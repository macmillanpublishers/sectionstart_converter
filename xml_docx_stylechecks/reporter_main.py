######### IMPORT SOME STANDARD PY LIBRARIES
import sys
import os
import zipfile
import shutil
import re
import logging
import time
# import inspect


######### IMPORT LOCAL MODULES
import cfg
import lib.addsectionstarts as addsectionstarts
import lib.stylereports as stylereports
import lib.generate_report as generate_report
import lib.setup_cleanup as setup_cleanup
import lib.usertext_templates as usertext_templates
import shared_utils.os_utils as os_utils
import shared_utils.check_docx as check_docx


######### LOCAL DECLARATIONS
inputfile = cfg.inputfile
inputfilename_noext = cfg.inputfilename_noext
# workingfile = cfg.workingfile
# ziproot = cfg.ziproot
this_outfolder = cfg.this_outfolder
tmpdir = cfg.tmpdir
report_dict = {}
# template_ziproot = cfg.template_ziproot
macmillan_template = cfg.macmillan_template
report_emailed = False


######### SETUP LOGGING
logfile = os.path.join(cfg.logdir, "{}_{}_{}.txt".format(cfg.script_name, inputfilename_noext, time.strftime("%y%m%d-%H%M%S")))
cfg.defineLogger(logfile, cfg.loglevel)
logger = logging.getLogger(__name__)


######### RUN!
# only run if this script is being invoked directly
if __name__ == '__main__':
    try:
        ########## SETUP
        # get submitter name, email
        submitter_email, display_name = setup_cleanup.getSubmitterViaAPI(cfg.inputfile)
        logger.info("Submitter name:'%s', email: '%s'" % (submitter_email, display_name))

        # setup & create tmpdir, cleanup outfolder (archive existing), copy infile to tmpdir
        tmpdir, workingfile, ziproot, template_ziproot, stylereport_json, alerts_json = setup_cleanup.setupFolders(tmpdir, inputfile, cfg.inputfilename, this_outfolder, inputfilename_noext)

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
            os_utils.dumpJSON(report_dict, stylereport_json)

        ########## SKIP RUNNING STUFF, LOG ALERTS
        else:
            logger.warn("* * Skipping Style Report:")
            if percent_styled < 50:
                errstring = usertext_templates.alerts()["notstyled"].format(percent_styled=percent_styled)
                os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))
            if version_result != "up_to_date":
                errstring = usertext_templates.alerts()["notstyled"].format(current_version=current_version, template_version=template_version)
                os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))
            if protection == True:
                errstring = usertext_templates.alerts()["protected"]
                os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))

        ########## CLEANUP
        # includes writing files to outfolder, sending mail to submitter, rm'ing tmpdir
        report_emailed = setup_cleanup.cleanupforReporterOrConverter(cfg.script_name, this_outfolder, workingfile, cfg.inputfilename, report_dict, cfg.stylereport_txt, alerts_json, tmpdir, submitter_email, display_name)

    except:
        ########## LOG ERROR INFO
        # log to logfile for dev
        logger.exception("ERROR ------------------ :")
        # # log to errfile_json for user
        # # invokedby_script = os.path.splitext(os.path.basename(inspect.stack()[0][1]))[0]
        # os_utils.logAlerttoJSON(alerts_json, "error", "A fatal error was encountered while running '%s'_main.py.\n\nPlease email workflows@macmillan.com for assistance." % invokedby_script)

        setup_cleanup.cleanupException(this_outfolder, workingfile, cfg.inputfilename, alerts_json, tmpdir, cfg.logdir, inputfilename_noext, cfg.script_name, logfile, report_emailed, submitter_email, display_name)
