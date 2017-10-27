######### IMPORT PY LIBRARIES

import os
import shutil
import re
import uuid
import json
import sys
import logging
from lxml import etree
from distutils.version import StrictVersion as Version

######### IMPORT LOCAL MODULES

if __name__ == '__main__':
	# to go up a level to read cfg when invoking from this script (for testing).
	import imp
	parentpath = os.path.join(sys.path[0], '..', 'cfg.py')
	cfg = imp.load_source('cfg', parentpath)
else:
	import cfg

######### LOCAL DECLARATIONS

# Word namespace vars
wnamespace = cfg.wnamespace
wordnamespaces = cfg.wordnamespaces

# Template dirs & files
# template_customprops_xml = cfg.template_customprops_xml

# doc files
# doc_xml = cfg.doc_xml
# styles_xml = cfg.styles_xml
# customprops_xml = cfg.customprops_xml
# sectionstart_versionstring = cfg.sectionstart_versionstring
# settings_xml = cfg.settings_xml

# initialize logger
logger = logging.getLogger(__name__)

#---------------------  METHODS

def get_docxVersion(customprops_xml):
	logger.debug("getting docx Version...")
	# default value for version comparison if no version marker exists
	# note: disutils.version.StrictVersion requires a x.x strcture, whole number is not valid
	versionstring = "0.0"
	if os.path.exists(customprops_xml):

		# get root / tree of docx customProps.xml
		customprops_tree = etree.parse(customprops_xml)
		customprops_root = customprops_tree.getroot()

		# check for version element in docx customProps.xml
		version_el = customprops_tree.xpath(".//*[@name='Version']", namespaces=wordnamespaces)
		if len(version_el) > 0:
			versionstring = customprops_tree.xpath(".//*[@name='Version']/vt:lpwstr", namespaces=wordnamespaces)[0].text

		if versionstring[0] == 'v':
			versionstring = versionstring[1:]

	logger.debug("versionstring value: '%s'" %  versionstring)
	return versionstring

def compare_docxVersions(current_version, template_version, sectionstart_version):
	logger.debug("comparing docx version to template...")
	# adding try statement since the Version library can be a little particular.
	try:
		if Version(current_version) == Version(template_version):
			version_compare = "up_to_date"
		elif Version(current_version) >= Version(sectionstart_version):
			version_compare = "has_section_starts"
		else:
			version_compare = "no_version"

		logger.debug("version_compare value: '%s'" %  version_compare)
		return version_compare
	except Exception, e:
		logger.error('Failed version compare, exiting', exc_info=True)
		sys.exit(1)

def macmillanStyleCount(doc_xml, styles_xml):
	logger.debug("Counting total paras, Macmillan styled paras...")

	doc_tree = etree.parse(doc_xml)
	doc_root = doc_tree.getroot()
	styles_tree = etree.parse(styles_xml)
	styles_root = styles_tree.getroot()

	total_paras = len(doc_tree.xpath(".//w:p", namespaces=wordnamespaces))
	macmillan_styled_paras = 0

	for para_style in doc_root.findall(".//*w:pStyle", wordnamespaces):
		# get stylename from each para
		stylename = para_style.get('{%s}val' % wnamespace)

		# search styles.xlm for corresponding full stylename
		stylesearchstring = ".//w:style[@w:styleId='%s']/w:name" % stylename
		stylematch = styles_root.find(stylesearchstring, wordnamespaces)

		# get fullname value and test for parentheses
		stylename_full = stylematch.get('{%s}val' % wnamespace)
		if '(' in stylename_full:
			# count paras with parentheses
			macmillan_styled_paras += 1

	# the multiplying by a factor with '.0' in the numerator forces the result to be a float for python 2.x
	percent_styled = (macmillan_styled_paras * 100.0) / total_paras

	logger.debug("total paras:'%s', macmillan styled:'%s', percent_styled:'%s'" % (total_paras, macmillan_styled_paras, percent_styled))
	return percent_styled, macmillan_styled_paras, total_paras

# This function consolidates version test functions
def version_test(customprops_xml, template_customprops_xml, sectionstart_versionstring):

	current_version = get_docxVersion(customprops_xml)
	template_version = get_docxVersion(template_customprops_xml)
	sectionstart_version = sectionstart_versionstring

	version_result = compare_docxVersions(current_version, template_version, sectionstart_version)
	return version_result, current_version, template_version

# to test trackchanges and doc protection, and existence of any other top level param in settings
def checkSettingsXML(settings_xml, settingstring):
	logger.debug("checking settings_xml for '%s'..." % settingstring)
	value = ""
	# get root / tree of docx customProps.xml
	settings_tree = etree.parse(settings_xml)

	# check for version element in docx settings.xml
	searchstring = ".//w:{}".format(settingstring)
	if len(settings_tree.xpath(searchstring, namespaces=wordnamespaces)) > 0:
		value = True
	else:
		value = False
	logger.debug("value for '%s' is '%s'" % (settingstring, value))
	return value

#---------------------  MAIN

# only run if this script is being invoked directly
if __name__ == '__main__':
	# set up debug log to console
	logging.basicConfig(level=logging.DEBUG)
    settings_xml = cfg.settings_xml
    
	# Testing
	protection = checkSettingsXML(settings_xml, "documentProtection")
	logger.debug("protection: '%s'" % protection)
	trackchanges = checkSettingsXML(settings_xml, "trackRevisions")
	logger.debug("trackchanges: '%s'" % trackchanges)
