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
import lib.generate_report as generate_report
import lib.setup_cleanup as setup_cleanup
# import shared_utils.unzipDOCX as unzipDOCX
import shared_utils.zipDOCX as zipDOCX
import shared_utils.os_utils as os_utils
import shared_utils.lxml_utils as lxml_utils
import shared_utils.check_docx as check_docx

######### LOCAL DECLARATIONS
inputfile = cfg.inputfile
inputfilename_noext = cfg.inputfilename_noext
tmpdir = cfg.tmpdir
workingfile = cfg.workingfile
ziproot = cfg.ziproot
this_outfolder = cfg.this_outfolder
newdocxfile = cfg.newdocxfile
report_dict = {}
template_ziproot = cfg.template_ziproot
macmillan_template = cfg.macmillan_template


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
        # setup & create tmpdir, cleanup outfolder (archive existing), copy infile to tmpdir
        tmpdir, workingfile = setup_cleanup.setupFolders(tmpdir, inputfile, cfg.inputfilename, this_outfolder)

        # copy template to tmpdir, unzip infile and tmpdir
        setup_cleanup.copyTemplateandUnzipFiles(macmillan_template, tmpdir, workingfile, ziproot, template_ziproot)

        # ## old setup:
        # # create & cleanup outfolder if it does not exist:
        # logger.info("Create tmpdir, create & cleanup project outfolder")
        # os_utils.setupOutfolder(this_outfolder)
        #
        # # move inputfile to tmpdir as workingfile
        # logger.info('Moving input file ({}) and template to tmpdir and unzipping'.format(inputfilename_noext))
        # # os_utils.movefile(inputfile, workingfile)			# for production
        # os_utils.copyFiletoFile(inputfile, workingfile)		# debug/testing only
        #
        # # move template to the tmpdir
        # os_utils.copyFiletoFile(macmillan_template, os.path.join(tmpdir, os.path.basename(macmillan_template)))
        #
        # ### unzip the manuscript to ziproot, template to template_ziproot
        # os_utils.rm_existing_os_object(ziproot, 'ziproot')
        # os_utils.rm_existing_os_object(ziproot, 'template_ziproot')
        # unzipDOCX.unzipDOCX(workingfile, ziproot)
        # unzipDOCX.unzipDOCX(macmillan_template, template_ziproot)

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

            # # debug test:
            # os_utils.logAlerttoJSON(cfg.alerts_json, "error", "You really messed up")
            # os_utils.logAlerttoJSON(cfg.alerts_json, "warning", "You  messed up a little")
            # os_utils.logAlerttoJSON(cfg.alerts_json, "notice", "You might want to stop messing up")

            # write our json for style report
            logger.debug("Writing stylereport.json")
            os_utils.dumpJSON(report_dict, cfg.stylereport_json)

            # # write our stylereport.txt
            # logger.debug("Writing stylereport.txt to outfolder")
            # generate_report.generateReport(report_dict, cfg.stylereport_txt)

        ########## SKIP RUNNING STUFF, LOG ALERTS
        # Doc is not styled or has section start styles already
        # skip attach template and skip adding section starts
        else:
            logger.warn("* * Skipping Converter:")
            if percent_styled < 50:
                errstring = "This .docx has {} percent of paragraphs styled with Macmillan styles".format(percent_styled)
                os_utils.logAlerttoJSON(cfg.alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))
            if version_result != "no_version":
                errstring = "This document has already has a template attached with section_start styles"
                os_utils.logAlerttoJSON(cfg.alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))
                if version_result == "has_section_starts":
                    noticestring = "Newer available version of the macmillan style template (this .docx's version: {}, template version: {})".format(current_version, template_version)
                    os_utils.logAlerttoJSON(cfg.alerts_json, "error", noticestring)
                    logger.warn("* NOTE: {}".format(noticestring))
            if protection == True:
                errstring = "* This .docx has protection enabled."
                os_utils.logAlerttoJSON(cfg.alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))

        # here: call some piece of the reporter_main
        # to generate a style_report.txt from this json
        # and write it to the outfolder
        # and maybe send an email

        ########## CLEANUP
        setup_cleanup.cleanupforReporterOrConverter(this_outfolder, workingfile, cfg.inputfilename, report_dict, cfg.stylereport_txt, cfg.alerts_json, tmpdir)

        # # ##old cleanup:
        # # Return original file to user
        # logger.info("Copying original file to outfolder/original_file dir")
        # if not os.path.isdir(os.path.join(this_outfolder, "original_file")):
        #     os.makedirs(os.path.join(this_outfolder, "original_file"))
        # os_utils.copyFiletoFile(workingfile, os.path.join(this_outfolder, "original_file", cfg.inputfilename))
        #
        # # write our alertfile.txt if necessary (are we using this for converter? I guess so?)
        # logger.debug("Writing alerts.txt to outfolder")
        # os_utils.writeAlertstoTxtfile(cfg.alerts_json, this_outfolder)
        #
        # # Rm tmpdir
        # logger.debug("deleting tmp folder")
        # # os_utils.rm_existing_os_object(tmpdir, 'tmpdir')		# comment out for testing / debug

    except:
        ########## LOG ERROR INFO
        # log to logfile for dev
        logger.exception("ERROR ------------------ :")
        # log to errfile_json for user
        invokedby_script = os.path.splitext(os.path.basename(inspect.stack()[0][1]))[0]
        os_utils.logAlerttoJSON(cfg.alerts_json, "error", "A fatal error was encountered while running '%s'.\n\nPlease email workflows@macmillan.com for assistance." % invokedby_script)

        setup_cleanup.cleanupException(this_outfolder, workingfile, cfg.inputfilename, cfg.alerts_json, tmpdir, cfg.logdir, inputfilename_noext, cfg.script_name)
        
        # try:
        #     logger.warn("SENDING EMAIL ALERT RE: EXCEPTION:")
        #     # send an email to us!...
        #     # TK
        #
        #     #  save a copy of tmpdir to logdir for troubleshooting (since it will be deleted)
        #     logger.info("Backing up tmpdir to logfolder")
        #     os_utils.copyDir(tmpdir, os.path.join(cfg.logdir, "tmpdir_%s" % inputfilename_noext))
        #
        #     ########## ATTEMPT CLEANUP (same cleanup as in try block above)
        #     logger.warn("RUNNING CLEANUP FROM EXCEPTION:")
        #     # Return original file to user
        #     logger.info("Copying original file to outfolder/original_file dir")
        #     if not os.path.isdir(os.path.join(this_outfolder, "original_file")):
        #         os.makedirs(os.path.join(this_outfolder, "original_file"))
        #     os_utils.copyFiletoFile(workingfile, os.path.join(this_outfolder, "original_file", cfg.inputfilename))
        #
        #     # write our alertfile.txt if necessary (are we using this for converter? I guess so?)
        #     logger.info("Writing alerts.txt to outfolder")
        #     os_utils.writeAlertstoTxtfile(cfg.alerts_json, this_outfolder)
        #
        #     # Rm tmpdir
        #     logger.info("deleting tmp folder")
        #     # os_utils.rm_existing_os_object(tmpdir, 'tmpdir')		# comment out for testing / debug
        # except:
        #     logger.exception("ERROR with exception cleanup :")
