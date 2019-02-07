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
import lib.setup_cleanup as setup_cleanup
import lib.usertext_templates as usertext_templates
import lib.rsuite_validations as rsuite_validations
import shared_utils.zipDOCX as zipDOCX
import shared_utils.os_utils as os_utils
import shared_utils.check_docx as check_docx


######### LOCAL DECLARATIONS
inputfile = cfg.inputfile
inputfilename = cfg.inputfilename
inputfile_ext = cfg.inputfile_ext
inputfilename_noext = cfg.inputfilename_noext
original_inputfilename = cfg.original_inputfilename
tmpdir = cfg.tmpdir
workingfile = cfg.workingfile
ziproot = cfg.ziproot
this_outfolder = cfg.this_outfolder
newdocxfile = cfg.newdocxfile
stylereport_json = cfg.stylereport_json
alerts_json = cfg.alerts_json
template_ziproot = cfg.template_ziproot
macmillan_template = cfg.macmillan_template
report_dict = {}
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
        # get file submitter via api, copy infile to tmpdir, setup outfolder
        submitter_email, display_name, notdocx = setup_cleanup.setupforReporterOrConverter(inputfile, inputfilename, workingfile, this_outfolder, inputfile_ext)

        if notdocx == True:
            errstring = usertext_templates.alerts()["notdocx"].format(scriptname=cfg.script_name)
            os_utils.logAlerttoJSON(alerts_json, "error", errstring)
            logger.warn("* {}".format(errstring))
        else:
            # copy template to tmpdir, unzip infile and tmpdir
            logger.warn("cfg.script_name: %s, template: %s" % (cfg.script_name, macmillan_template)) # DEBUG
            setup_cleanup.copyTemplateandUnzipFiles(macmillan_template, tmpdir, workingfile, ziproot, template_ziproot)

            ########## CHECK DOCUMENT
            ### check and compare versions, styling percentage, doc protection
            logger.info('Comparing docx version to template, checking percent styled, checking if protection, trackchanges...')
            version_result, current_version, template_version = check_docx.version_test(cfg.customprops_xml, cfg.template_customprops_xml, cfg.templateversion_cutoff)
            percent_styled, macmillan_styled_paras, total_paras = check_docx.macmillanStyleCount(cfg.doc_xml, cfg.styles_xml)
            protection, tc_marker_found, trackchange_status = check_docx.getProtectionAndTrackChangesStatus(cfg.doc_xml, cfg.settings_xml)

            # note and accept all track changes
            if tc_marker_found == True:
                errstring = usertext_templates.alerts()["v_unaccepted_tcs"]
                os_utils.logAlerttoJSON(alerts_json, "warning", errstring)
                logger.warn("* {}".format(errstring))
                check_docx.acceptTrackChanges(cfg.doc_xml)
            # create warnings re: track changes:
            if trackchange_status == True:
                errstring = usertext_templates.alerts()["trackchange_enabled"]
                os_utils.logAlerttoJSON(alerts_json, "notice", errstring)
                logger.warn("* {}".format(errstring))


            ########## RUN STUFF
            if version_result == "up_to_date" and percent_styled >= 50 and protection == "":
                logger.info("Proceeding! (version='%s', percent_styled='%s', protection='%s')" % (version_result, percent_styled, protection))

                # handle docs where both style-sets exist.
                zerostylecheck = check_docx.stripDuplicateMacmillanStyles(cfg.doc_xml, cfg.styles_xml)
                if zerostylecheck:
                    logger.warn("duplicate Macmillan styles were found: legacy styles were replaced with new ones")

                # run our rsuite validations!
                report_dict = rsuite_validations.rsuiteValidations(report_dict)

                ### zip ziproot up as a docx into outfolder
                logger.info("Zipping updated xml into a .docx")
                os_utils.rm_existing_os_object(newdocxfile, 'newdocxfile')	# < --- this should get replaced with our fancy folder rename
                zipDOCX.zipDOCX(ziproot, newdocxfile)

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
                    errstring = usertext_templates.alerts()["r_err_oldtemplate"].format(current_version=current_version, template_version=template_version)
                    os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                    logger.warn("* {}".format(errstring))
                if protection:
                    errstring = usertext_templates.alerts()["protected"].format(protection=protection)
                    os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                    logger.warn("* {}".format(errstring))

        ########## CLEANUP
        # includes writing files to outfolder, sending mail to submitter, rm'ing tmpdir
        report_emailed = setup_cleanup.cleanupforReporterOrConverter(cfg.script_name, this_outfolder, workingfile, inputfilename, report_dict, cfg.stylereport_txt, alerts_json, tmpdir, submitter_email, display_name, original_inputfilename)

    except:
        ########## LOG ERROR INFO
        # log to logfile for dev
        logger.exception("ERROR ------------------ :")
        # # invokedby_script = os.path.splitext(os.path.basename(inspect.stack()[0][1]))[0]

        # go cleanup after this exception!
        setup_cleanup.cleanupException(this_outfolder, workingfile, inputfilename, alerts_json, tmpdir, cfg.logdir, inputfilename_noext, cfg.script_name, logfile, report_emailed, submitter_email, display_name, original_inputfilename)
