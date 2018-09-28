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
import lib.stylereports as stylereports
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
report_dict["validator_py_complete"] = False
template_ziproot = cfg.template_ziproot
macmillan_template = cfg.macmillan_template
alerts_json = cfg.alerts_json


######### SETUP LOGGING
logfile = os.path.join(cfg.logdir, "{}_{}_{}.txt".format(cfg.script_name, inputfilename_noext, time.strftime("%y%m%d-%H%M%S")))
if cfg.validator_logfile:
    cfg.defineLogger(cfg.validator_logfile, cfg.loglevel)
else:
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
        logger.info('Comparing docx version to template, checking percent styled, checking if protection, trackchanges...')
        version_result, current_version, template_version = check_docx.version_test(cfg.customprops_xml, cfg.template_customprops_xml, cfg.sectionstart_versionstring)
        percent_styled, macmillan_styled_paras, total_paras = check_docx.macmillanStyleCount(cfg.doc_xml, cfg.styles_xml)
        protection, tc_marker_found, trackchange_status = check_docx.getProtectionAndTrackChangesStatus(cfg.doc_xml, cfg.settings_xml)

        # log for the rest o the validator suite:
        report_dict["percent_styled"] = percent_styled

        ########## RUN STUFF
        # Basic requirements passed, proceed with validation & cleanup
        if percent_styled >= 50 and protection == "":
            logger.info("Proceeding! (percent_styled='%s', protection='%s')" % (percent_styled, protection))

            # note and accept all track changes
            if tc_marker_found == True:
                errstring = usertext_templates.alerts()["v_unaccepted_tcs"]
                os_utils.logAlerttoJSON(alerts_json, "warning", errstring)
                logger.warn("* {}".format(errstring))
                check_docx.acceptTrackChanges(cfg.doc_xml)

            # # # Attach the template as needed
            if version_result == "no_version":
                logger.info("'version_result' = '%s'. Attaching macmillan template (updating styles in styles.xml, etc)" % version_result)
                # print "* This .docx did not have section start styles, attaching up-to-date template"
                docx_uptodate = attachtemplate.attachTemplate()
            elif version_result == "newer_template_avail":
                logger.info("'version_result' = '%s'. Attaching macmillan template and adding 'Notice' alert." % version_result)
                noticestring = usertext_templates.alerts()["v_newertemplate_avail"].format(current_version=current_version, template_version=template_version)
                os_utils.logAlerttoJSON(alerts_json, "notice", noticestring)
                logger.warn("* NOTE: {}".format(noticestring))
                docx_uptodate = attachtemplate.attachTemplate()

            # # # add section starts!
            logger.info("Adding SectionStart styles to the document.xml!")
            report_dict = addsectionstarts.sectionStartCheck("insert", report_dict)

            # # # run docPrepare function(s)
            report_dict = doc_prepare.docPrepare(report_dict)

            # # # # run style report stuff for report!
            logger.info("Running style report functions")
            report_dict = stylereports.styleReports(report_dict)

            # # # remove non-printing heads:  has to run after style_report or refs for SectionStartneeded info is missing
            report_dict = doc_prepare.rmNonPrintingHeads(report_dict, cfg.doc_xml, cfg.nonprintingheads)

            ### zip ziproot up as a docx
            logger.info("Zipping updated xml into a .docx in the tmpfolder")
            os_utils.rm_existing_os_object(newdocxfile, 'validated_docx')            # < --- this should get replaced with our fancy folder rename
            zipDOCX.zipDOCX(ziproot, newdocxfile)

        ########## SKIP RUNNING STUFF, LOG ALERTS
        # Doc is not styled or has protection enabled, skip python validation
        else:
            logger.warn("* * Skipping Validation:")
            if percent_styled < 50:
                errstring = usertext_templates.alerts()["notstyled"].format(percent_styled=percent_styled)
                os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))
            if protection:
                errstring = usertext_templates.alerts()["protected"].format(protection=protection)
                os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))

        ########## CLEANUP
        # write our json for style report to tmpdir
        logger.debug("Writing stylereport.json")
        report_dict["validator_py_complete"] = True
        os_utils.dumpJSON(report_dict, cfg.stylereport_json)

        setup_cleanup.cleanupforValidator(this_outfolder, workingfile, cfg.inputfilename, report_dict, cfg.stylereport_txt, alerts_json, cfg.script_name)

    except:
        ########## LOG ERROR INFO
        # log to logfile for dev
        logger.exception("ERROR ------------------ :")
        # the last 4 parameters only apply to reporter and converter
        setup_cleanup.cleanupException(this_outfolder, workingfile, cfg.inputfilename, alerts_json, tmpdir, cfg.logdir, inputfilename_noext, cfg.script_name, cfg.validator_logfile, "", "", "", "")
