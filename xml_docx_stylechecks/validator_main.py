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
inputfilename_noext = cfg.inputfilename_noext
workingfile = cfg.inputfile
ziproot = cfg.ziproot
this_newdocxfile = os.path.join(tmpfolder, "{}_validated.docx".format(inputfilename_noext))
report_dict = {}
template_ziproot = cfg.template_ziproot
macmillan_template = cfg.macmillan_template


######### SETUP LOGGING
logfile = os.path.join(cfg.logdir, "validate-py_{}_{}.txt".format(inputfilename_noext, time.strftime("%y%m%d-%H%M%S")))
cfg.defineLogger(logfile, cfg.loglevel)
logger = logging.getLogger(__name__)


######### RUN!
# only run if this script is being invoked directly
if __name__ == '__main__':

	logger.info('Moving template to tmpdir and unzipping it and input file ({})'.format(inputfilename_noext))

	# move template to the tmpdir
	os_utils.copyFiletoFile(macmillan_template, os.path.join(cfg.tmpdir, os.path.basename(macmillan_template)))	

	### unzip the manuscript to ziproot, template to template_ziproot
	os_utils.rm_existing_os_object(ziproot, 'ziproot')
	os_utils.rm_existing_os_object(ziproot, 'template_ziproot')
	unzipDOCX.unzipDOCX(workingfile, ziproot)
	unzipDOCX.unzipDOCX(macmillan_template, template_ziproot)

	# create outfolder if it does not exist:
	logger.info("Creating project outfolder")	
	if not os.path.isdir(this_outfolder):
		os.makedirs(this_outfolder)

	### check and compare versions, styling percentage, doc protection
	logger.info('Comparing docx version to template, checking percent styled, checking if protected doc...')	
	version_result, current_version, template_version = check_docx.version_test(cfg.customprops_xml, cfg.template_customprops_xml, cfg.sectionstart_versionstring)
	percent_styled, macmillan_styled_paras, total_paras = check_docx.macmillanStyleCount(cfg.doc_xml, cfg.styles_xml)
	protection = check_docx.checkSettingsXML(cfg.settings_xml, "documentProtection")

	# Basic requirements passed, proceed with validation & cleanup 
	if percent_styled >= 50 and protection == False:
		logger.info("Proceeding! (percent_styled='%s', protection='%s')" % (version_result, percent_styled, protection))
		
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

		### zip ziproot up as a docx 
		logger.info("Zipping updated xml into a .docx in the tmpfolder")
		os_utils.rm_existing_os_object(this_newdocxfile, 'validated_docx')			# < --- this should get replaced with our fancy folder rename
		zipDOCX.zipDOCX(ziproot, this_newdocxfile)

		# write our json for style report to tmpdir
		logger.debug("Writing stylereport.json")	
		os_utils.dumpJSON(report_dict, cfg.stylereport_json)

	# Doc is not styled or has protection enabled, skip python validation
	else: 
		logger.warn("* * Skipping Validation:")
		if percent_styled < 50:
			logger.warn("* This .docx has {} percent of paragraphs styled with Macmillan styles".format(percent_styled))
		if protection == True:		
			logger.warn("* This .docx has protection enabled.")	

