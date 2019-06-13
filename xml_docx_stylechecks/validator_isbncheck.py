######### IMPORT SOME STANDARD PY LIBRARIES
import sys
import os
import zipfile
import logging
import time
# make sure to install lxml: sudo pip install lxml
from lxml import etree


######### IMPORT LOCAL MODULES
import cfg
import lib.doc_prepare as doc_prepare
import lib.setup_cleanup as setup_cleanup
import lib.usertext_templates as usertext_templates
import shared_utils.zipDOCX as zipDOCX
import shared_utils.os_utils as os_utils
import shared_utils.check_docx as check_docx
import shared_utils.lxml_utils as lxml_utils


######### LOCAL DECLARATIONS
inputfilename_noext = cfg.inputfilename_noext
workingfile = cfg.inputfile
ziproot = cfg.ziproot
tmpdir = cfg.tmpdir
this_outfolder = cfg.this_outfolder
newdocxfile = cfg.newdocxfile
isbn_dict = {}
isbn_dict["completed"] = False
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
        ### check doc protection
        logger.info('Checking doc protection, trackchanges...')
        protection, tc_marker_found, trackchange_status = check_docx.getProtectionAndTrackChangesStatus(cfg.doc_xml, cfg.settings_xml, cfg.footnotes_xml, cfg.endnotes_xml)

        # log for the rest o the validator suite:
        isbn_dict["password_protected"] = protection

        ########## RUN STUFF
        # Basic requirements passed, proceed with validation & cleanup
        if protection == "":
            logger.info("Proceeding with isbn_check! protection='%s')" % (protection))

            # get doc_root
            doc_xml = cfg.doc_xml
            doc_tree = etree.parse(doc_xml)
            doc_root = doc_tree.getroot()
            isbnstyle = lxml_utils.transformStylename(cfg.isbnstyle)
            hyperlinkstyle = lxml_utils.transformStylename(cfg.hyperlinkstyle)

            # # # scan for styled ISBNs and strip non-ISBN chars
            isbn_dict, isbn_dict["styled_isbns"] = doc_prepare.removeNonISBNsfromISBNspans(isbn_dict, doc_root, isbnstyle, cfg.isbnspanregex)

            # # # scan for unstyled ISBNs and style them. Also captures properly styled isbns that may have spanned multiple 'runs' in xml
            isbn_dict, isbn_dict["programatically_styled_isbns"] = doc_prepare.styleLooseISBNs(isbn_dict, cfg.isbnregex, cfg.isbnspanregex, doc_root, isbnstyle, hyperlinkstyle)

            # # # run it again, to clean up any isbn-styled leading/trailing txt created incidentally from the last method
            isbn_dict, isbn_dict["manuscript_isbns"] = doc_prepare.removeNonISBNsfromISBNspans(isbn_dict, doc_root, isbnstyle, cfg.isbnspanregex)

            ### write updated stuff to file:
            os_utils.writeXMLtoFile(doc_root, doc_xml)

            ### zip ziproot up as a docx
            logger.info("Zipping updated xml to replace workingfile .docx in the tmpfolder")
            # os_utils.rm_existing_os_object(newdocxfile, 'validated_docx') <-- shouldnt be necessary, zipfile.py should overwrite
            zipDOCX.zipDOCX(ziproot, workingfile) # < for prod
            # zipDOCX.zipDOCX(ziproot, newdocxfile)  # < for testing

        ########## SKIP RUNNING STUFF, LOG ALERTS
        # Doc is not styled or has protection enabled, skip python validation
        else:
            logger.warn("* * Skipping ISBN_check:")
            if protection:
                errstring = usertext_templates.alerts()["protected"].format(protection=protection)
                os_utils.logAlerttoJSON(alerts_json, "error", errstring)
                logger.warn("* {}".format(errstring))

        ########## CLEANUP
        # write our json for style report to tmpdir
        logger.debug("Writing isbn_check.json")
        isbn_dict["completed"] = True
        os_utils.dumpJSON(isbn_dict, cfg.isbn_check_json)

        setup_cleanup.cleanupforValidator(this_outfolder, workingfile, cfg.inputfilename, '', cfg.stylereport_txt, alerts_json, cfg.script_name)

    except:
        ########## LOG ERROR INFO
        # log to logfile for dev
        logger.exception("ERROR ------------------ :")
        # the last 4 parameters only apply to reporter and converter
        setup_cleanup.cleanupException(this_outfolder, workingfile, cfg.inputfilename, alerts_json, tmpdir, cfg.logdir, inputfilename_noext, cfg.script_name, cfg.validator_logfile, "", "", "", "")
