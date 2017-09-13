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
import shared_utils.unzipDOCX as unzipDOCX
import shared_utils.zipDOCX as zipDOCX
import shared_utils.os_utils as os_utils
import shared_utils.check_docx as check_docx


######### LOCAL DECLARATIONS
inputfile = cfg.inputfile
inputfilename_noext = cfg.inputfilename_noext
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


    # create & cleanup tmpfolder, outfolder if they do not exist:
    logger.info("Create tmpdir, create & cleanup project outfolder")
    tmpdir = os_utils.setupTmpfolder(cfg.tmpdir)
    os_utils.setupOutfolder(this_outfolder)

    logger.info('Moving input file ({}) and template to tmpdir and unzipping'.format(inputfilename_noext))

    # move inputfile to tmpdir as workingfile
    # os_utils.movefile()			# for production
    os_utils.copyFiletoFile(inputfile, workingfile)		# debug/testing only

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

    # Doc is styled and has no version: attach template and add section starts!
    if version_result == "no_version" and percent_styled >= 50 and protection == False:
        logger.info("Proceeding! (version='%s', percent_styled='%s', protection='%s')" % (version_result, percent_styled, protection))

        # # # attach the template
        logger.info("'Attaching' macmillan template (updating styles in styles.xml, etc)")
        # print "* This .docx did not have section start styles, attaching up-to-date template"
        docx_uptodate = attachtemplate.attachTemplate()

        # # # add section starts!
        logger.info("Adding SectionStart styles to the document.xml!")
        # report_dict = addsectionstarts.sectionStartCheck("report", report_dict)
        report_dict = addsectionstarts.sectionStartCheck("insert", report_dict)

        ### zip ziproot up as a docx into outfolder
        logger.info("Zipping updated xml into a .docx")
        os_utils.rm_existing_os_object(newdocxfile, 'newdocxfile')	# < --- this should get replaced with our fancy folder rename
        zipDOCX.zipDOCX(ziproot, newdocxfile)

        # write our json for style report
        logger.debug("Writing stylereport.json")
        os_utils.dumpJSON(report_dict, cfg.stylereport_json)

    # Doc is not styled or has section start styles already
    # skip attach template and skip adding section starts
    else:
        logger.warn("* * Skipping Converter:")
        if percent_styled < 50:
            logger.warn("* This .docx has {} percent of paragraphs styled with Macmillan styles".format(percent_styled))
        if version_result != "no_version":
            logger.warn("* This document has already has a template attached with section_start styles")
            if version_result == "has_section_starts":
                logger.warn("* NOTE: Newer available version of the macmillan style template (this .docx's version: {}, template version: {})".format(current_version, template_version))
        if protection == True:
            logger.warn("* This .docx has protection enabled.")

    # here: call some piece of the reporter_main
    # to generate a style_report.txt from this json
    # and write it to the outfolder
    # and maybe send an email

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
    # os_utils.rm_existing_os_object(tmpdir, 'tmpdir')		# comment out for testing / debug
