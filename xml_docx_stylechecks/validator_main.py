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
import lib.attachtemplate as attachtemplate
import lib.addsectionstarts as addsectionstarts
import lib.doc_prepare as doc_prepare
import lib.generate_report as generate_report
import lib.setup_cleanup as setup_cleanup
import lib.usertext_templates as usertext_templates
import shared_utils.zipDOCX as zipDOCX
import shared_utils.os_utils as os_utils
import shared_utils.check_docx as check_docx


######### LOCAL DECLARATIONS
inputfilename_noext = cfg.inputfilename_noext
workingfile = cfg.inputfile
ziproot = cfg.ziproot
tmpdir = cfg.tmpdir
this_outfolder = cfg.this_outfolder
newdocxfile = cfg.newdocxfile
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
        # copy template to tmpdir, unzip infile and tmpdir
        setup_cleanup.copyTemplateandUnzipFiles(macmillan_template, tmpdir, workingfile, ziproot, template_ziproot)

        ########## CHECK DOCUMENT
        ### check and compare versions, styling percentage, doc protection
        logger.info('Comparing docx version to template, checking percent styled, checking if protected doc...')
        version_result, current_version, template_version = check_docx.version_test(cfg.customprops_xml, cfg.template_customprops_xml, cfg.sectionstart_versionstring)
        percent_styled, macmillan_styled_paras, total_paras = check_docx.macmillanStyleCount(cfg.doc_xml, cfg.styles_xml)
        protection = check_docx.checkSettingsXML(cfg.settings_xml, "documentProtection")

        ########## RUN STUFF
        # Basic requirements passed, proceed with validation & cleanup
        if percent_styled >= 50 and protection == False:
            logger.info("Proceeding! (percent_styled='%s', protection='%s')" % (percent_styled, protection))

            # # # Attach the template as needed
            if version_result == "no_version":
                logger.info("'Attaching' macmillan template (updating styles in styles.xml, etc)")
                # print "* This .docx did not have section start styles, attaching up-to-date template"
                docx_uptodate = attachtemplate.attachTemplate()
            elif version_result == "has_section_starts":
                logger.warn("* NOTE: Newer available version of the macmillan style template (this .docx's version: {}, template version: {})".format(current_version, template_version))

            # # # add section starts!
            logger.info("Adding SectionStart styles to the document.xml!")
            report_dict = addsectionstarts.sectionStartCheck("insert", report_dict)

            # # # run docPrepare function(s)
            report_dict = doc_prepare.docPrepare(report_dict)

            # # # # run other style report stuff for report!
            # logger.info("Running other style report functions")
            # report_dict = stylereports.styleReports(report_dict)

            ### zip ziproot up as a docx
            logger.info("Zipping updated xml into a .docx in the tmpfolder")
            os_utils.rm_existing_os_object(newdocxfile, 'validated_docx')            # < --- this should get replaced with our fancy folder rename
            zipDOCX.zipDOCX(ziproot, newdocxfile)

            # write our json for style report to tmpdir
            logger.debug("Writing stylereport.json")
            os_utils.dumpJSON(report_dict, cfg.stylereport_json)

        ########## SKIP RUNNING STUFF, LOG ALERTS
        # Doc is not styled or has protection enabled, skip python validation
        else:
            logger.warn("* * Skipping Validation:")
            if percent_styled < 50:
                errstring = usertext_templates.alerts()["notstyled"].format(percent_styled=percent_styled)
                os_utils.logAlerttoJSON(cfg.alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))
            if protection == True:
                errstring = usertext_templates.alerts()["protected"]
                os_utils.logAlerttoJSON(cfg.alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))

        ########## CLEANUP
        setup_cleanup.cleanupforValidator(this_outfolder, workingfile, cfg.inputfilename, report_dict, cfg.stylereport_txt, cfg.alerts_json)

    except:
        ########## LOG ERROR INFO
        # log to logfile for dev
        logger.exception("ERROR ------------------ :")
        # log to errfile_json for user
        # invokedby_script = os.path.splitext(os.path.basename(inspect.stack()[0][1]))[0]
        # os_utils.logAlerttoJSON(cfg.alerts_json, "error", "A fatal error was encountered while running '%s'.\n\nPlease email workflows@macmillan.com for assistance." % invokedby_script)

        setup_cleanup.cleanupException(this_outfolder, workingfile, cfg.inputfilename, cfg.alerts_json, tmpdir, cfg.logdir, inputfilename_noext, cfg.script_name, logfile)
