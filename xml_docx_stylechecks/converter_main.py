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
import lib.attachtemplate as attachtemplate
import lib.addsectionstarts as addsectionstarts
import lib.generate_report as generate_report
import lib.setup_cleanup as setup_cleanup
import lib.usertext_templates as usertext_templates
import shared_utils.zipDOCX as zipDOCX
import shared_utils.os_utils as os_utils
import shared_utils.lxml_utils as lxml_utils
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
# w_logger = logging.getLogger('w_logger')


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
            setup_cleanup.copyTemplateandUnzipFiles(macmillan_template, tmpdir, workingfile, ziproot, template_ziproot)

            ########## CHECK DOCUMENT
            ### check and compare versions, styling percentage, doc protection
            logger.info('Comparing docx version to template, checking percent styled, checking if protected doc...')
            version_result, current_version, template_version = check_docx.version_test(cfg.customprops_xml, cfg.template_customprops_xml, cfg.sectionstart_versionstring)
            percent_styled, macmillan_styled_paras, total_paras = check_docx.macmillanStyleCount(cfg.doc_xml, cfg.styles_xml)
            protection = check_docx.checkSettingsXML(cfg.settings_xml, "documentProtection")

            ########## RUN STUFF
            # Doc is styled and has no version: attach template and add section starts!
            if version_result == "no_version" and percent_styled >= 50 and protection == False:
                logger.info("Proceeding! (version='%s', percent_styled='%s', protection='%s')" % (version_result, percent_styled, protection))

                # # # attach the template
                logger.info("'Attaching' macmillan template (updating styles in styles.xml, etc)")
                # print "* This .docx did not have section start styles, attaching up-to-date template"
                docx_uptodate = attachtemplate.attachTemplate()

                # # # add section starts!
                logger.info("Adding SectionStart styles to the document.xml!")
                # report_dict = addsectionstarts.sectionStartCheck("report", report_dict).  The true is to enable autnumbering section-start para content
                report_dict = addsectionstarts.sectionStartCheck("insert", report_dict, True)

                ### zip ziproot up as a docx into outfolder
                logger.info("Zipping updated xml into a .docx")
                os_utils.rm_existing_os_object(newdocxfile, 'newdocxfile')	# < --- this should get replaced with our fancy folder rename
                zipDOCX.zipDOCX(ziproot, newdocxfile)

                # write our json for style report
                logger.debug("Writing stylereport.json")
                os_utils.dumpJSON(report_dict, stylereport_json)

            ########## SKIP RUNNING STUFF, LOG ALERTS
            # Doc is not styled or has section start styles already
            # skip attach template and skip adding section starts
            else:
                logger.warn("* * Skipping Converter:")
                if percent_styled < 50:
                    errstring = usertext_templates.alerts()["notstyled"].format(percent_styled=percent_styled)
                    os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                    logger.warn("* {}".format(errstring))
                if version_result != "no_version":
                    errstring = usertext_templates.alerts()["c_has_template"]
                    os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                    logger.warn("* {}".format(errstring))
                    if version_result == "has_section_starts":
                        noticestring = usertext_templates.alerts()["c_newertemplate_avail"].format(current_version=current_version, template_version=template_version)
                        os_utils.logAlerttoJSON(alerts_json, "notice", noticestring)
                        logger.warn("* NOTE: {}".format(noticestring))
                if protection == True:
                    errstring = usertext_templates.alerts()["protected"]
                    os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                    logger.warn("* {}".format(errstring))

        ########## CLEANUP
        # includes writing files to outfolder, sending mail to submitter, rm'ing tmpdir
        report_emailed = setup_cleanup.cleanupforReporterOrConverter(cfg.script_name, this_outfolder, workingfile, inputfilename, report_dict, cfg.stylereport_txt, alerts_json, tmpdir, submitter_email, display_name, original_inputfilename, newdocxfile)

    except:
        ########## LOG ERROR INFO
        # log to logfile for dev
        logger.exception("ERROR ------------------ :")

        # go cleanup after this exception!
        setup_cleanup.cleanupException(this_outfolder, workingfile, inputfilename, alerts_json, tmpdir, cfg.logdir, inputfilename_noext, cfg.script_name, logfile, report_emailed, submitter_email, display_name, original_inputfilename)
