######### IMPORT PY LIBRARIES

import os
import shutil
import re
import uuid
import json
import sys
import logging
# make sure to install lxml: sudo pip install lxml
from lxml import etree
from distutils.version import StrictVersion as Version


######### IMPORT LOCAL MODULES

import cfg

# initialize logger
logger = logging.getLogger(__name__)

######### LOCAL DECLARATIONS

# Local namespace vars
wnamespace = cfg.wnamespace
wordnamespaces = cfg.wordnamespaces


#---------------------  METHODS

# should move this to file handling file for re-use
def writeXMLtoFile(root, file):
    newfile = open(file, 'w')
    with newfile as f:
        f.write(etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes"))
        f.close()

# generate random hexadecimal 8 char id beginning with 00, for Word revision ID
def generate_rsid():
    idbase = uuid.uuid4().hex
    idshort = idbase[:8]
    idshort = "00" + idshort[-6:]
    idupper = idshort.upper()
    return str(idupper)

# test generated rsid for uniqueness, re-create as needed
def get_unique_rsid(settings_root, wordnamespaces):
	# generate random id
    rsiduniq = generate_rsid()
    # scan docx rsid entries to make sure we are unique
    rsid_searchstring = './/*w:rsid[@w:val="%s"]' % rsiduniq
    # create new id until we have a unique one
    while len(settings_root.findall(rsid_searchstring, wordnamespaces)) > 0:
        print rsiduniq + " already exists, generating another id"
        rsiduniq = generate_rsid()
        rsid_searchstring = './/*w:p[@w14:paraId="%s"]' % rsiduniq
    return str(rsiduniq)

## note from rsidtable branch--some types of saves / conversions appear to wipe the rsid table,
##  wherein at this point we would need to generate a new rsid table, rsidroot value, and record our rsid itself.
##  since upon further reading, rsids are really only for merging documents, I'm going to try to bypass adding rsid's if the table doesn't exist.
def setupRSID(settings_root, wordnamespaces, wnamespace, settings_xml):
    # generate a unique rsid
    new_rsid = get_unique_rsid(settings_root, wordnamespaces)
    # create risd etree element
    new_rsid_el = etree.Element("{%s}rsid" % wnamespace)
    new_rsid_el.attrib["{%s}val" % wnamespace] = new_rsid
    # insert rsid in docx's settings.xml rsids table
    rsids = settings_root.find('.//w:rsids', wordnamespaces)
    # test if rsid table is present:
    if rsids is not None:
    	rsids.append(new_rsid_el)
    	# update the settings.xmlfile
    	writeXMLtoFile(settings_root, settings_xml)
    else:
        #no rsids table present, skipping adding rsids with styles
        new_rsid = ""

    return new_rsid


def	updateCustomDocProps(customprops_xml, template_customprops_xml):
	# if customProps.xml does not exist, just copy it from the template
	if not os.path.exists(customprops_xml):
		shutil.copyfile(template_customprops_xml, customprops_xml)
	# if it DOES exist, we will to go in & copy the version element from the template
	else:
		# get template version element for copy/paste
		template_customprops_tree = etree.parse(template_customprops_xml)
		template_version_el = template_customprops_tree.xpath(".//*[@name='Version']", namespaces=wordnamespaces)[0]

		# get root / tree of docx customProps.xml
		customprops_tree = etree.parse(customprops_xml)
		customprops_root = customprops_tree.getroot()

		# check for version element in docx customProps.xml
		version_el = customprops_tree.xpath(".//*[@name='Version']", namespaces=wordnamespaces)
		if len(version_el) > 0:
			# if version docProp already exists, remove it
			customprops_root.remove(version_el[0])

		# add template's version element to docx
		customprops_root.append(template_version_el)
		# update the customProps.xml file
		writeXMLtoFile(customprops_root, customprops_xml)


def	edit_relsFile(rels_tree, template_rels_tree, rels_file):
	# check if customdocprops relationship exists in _rels/.rels
	if not rels_tree.xpath('.//*[@Target="docProps/custom.xml"]'):
		# get custom relationship element from template
		custom_rel_el = template_rels_tree.xpath('.//*[@Target="docProps/custom.xml"]')[0]
		# get its Id attribute
		custom_rel_el_Id = custom_rel_el.attrib['Id']

		# see if there's already a relationship with that rId
		if rels_tree.xpath('.//*[@Id="%s"]' % custom_rel_el_Id ):
			# increment until we find unused rId
			while rels_tree.xpath('.//*[@Id="%s"]' % custom_rel_el_Id ):
				custom_rel_el_Id = ('rId' + str(int(custom_rel_el_Id.replace('rId','')) + 1))
			# now update our element with new rId
			custom_rel_el.attrib['Id'] = custom_rel_el_Id

		# now append the new relationship to the rels_root
		rels_root = rels_tree.getroot()
		rels_root.append(custom_rel_el)
		# update the _rels/.rels file
		writeXMLtoFile(rels_root, rels_file)


def	editContentTypes(contenttypes_tree, template_contenttypes_tree, contenttypes_xml):
	contenttypes_root = contenttypes_tree.getroot()
	# if this override does not already exist...
	if contenttypes_root.find('.//*[@PartName="/docProps/custom.xml"]') is None:
		# get override element from template
		template_contenttypes_root = template_contenttypes_tree.getroot()
		override_el = template_contenttypes_root.find('.//*[@PartName="/docProps/custom.xml"]')
		# write / overwrite existing related override element in .docx
		contenttypes_root.append(override_el)
		# update the content_types.xml file
		writeXMLtoFile(contenttypes_root, contenttypes_xml)

def styleWrite(templatestyleID, templatestyle_element, new_rsid, styles_root):
	# skip the only built-in tyle with a paren in the name
	if templatestyleID != "NormalWeb":
		# delete existing element if it is present
		searchstring = ".//w:style[@w:styleId='%s']" % templatestyleID
		existing_element = styles_root.find(searchstring, wordnamespaces)
		if existing_element is not None:
			styles_root.remove(existing_element)

        # check if rsid table is present:
        if new_rsid:
    		# apply newrsid to style before we write
            ## note from rsidtable branch: this may be moot too since the macmillantemplate is auto-generated; the source styles will never have
            ##  rsid attributes. We are not really dealing in rsids anywhere else, so effectively disabling them here too is fine--
            ##  and leaving handling in place in case we want to generate rsidtable & bring them back for other scripts (would want to generate rsid attribute if no exist)
    		for child in templatestyle_element.getchildren():
    			rsid_attr = "{%s}rsid" % wnamespace
    			if child.tag == rsid_attr:
    				rsid_val_key = '{%s}val' % wnamespace
    				child.set(rsid_val_key, new_rsid)

        # add new style to the docx
        styles_root.append(templatestyle_element)

def getNewTemplateStyles(new_rsid, template_styles_tree, styles_root, styles_xml):
	# cycles through all styles in the template with parentheses in the stylename
	for stylename in template_styles_tree.xpath(".//w:style/w:name[contains(@w:val, '(')]", namespaces=wordnamespaces):
		style_element = stylename.getparent()
		styleID_attrib = "{%s}styleId" % wnamespace
		styleID = style_element.get(styleID_attrib)

		# write / overwrite style to docx
		styleWrite(styleID, style_element, new_rsid, styles_root)

	# update the styles.xml file
	writeXMLtoFile(styles_root, styles_xml)

# This function groups everything from the above and runs this script as a whole :)
def attachTemplate():

	# # Template dirs & files
	template_customprops_xml = cfg.template_customprops_xml
	template_numbering_xml = cfg.template_numbering_xml

	# # doc files
	numbering_xml = cfg.numbering_xml
	styles_xml = cfg.styles_xml
	settings_xml = cfg.settings_xml
	customprops_xml = cfg.customprops_xml
	rels_file = cfg.rels_file
	contenttypes_xml = cfg.contenttypes_xml

	# etree trees of template files
	template_styles_tree = etree.parse(cfg.template_styles_xml)
	template_rels_tree = etree.parse(cfg.template_rels_file)
	template_contenttypes_tree = etree.parse(cfg.template_contenttypes_xml)

	# docx etree roots & trees
	styles_tree = etree.parse(cfg.styles_xml)
	styles_root = styles_tree.getroot()
	settings_tree = etree.parse(cfg.settings_xml)
	settings_root = settings_tree.getroot()
	rels_tree = etree.parse(cfg.rels_file)
	contenttypes_tree = etree.parse(cfg.contenttypes_xml)

	print "attaching style template..."

	# update the customDocProps
	updateCustomDocProps(customprops_xml, template_customprops_xml)

	# get a new unique rsid, add it to the settings.xml rsids table
	new_rsid = setupRSID(settings_root, cfg.wordnamespaces, cfg.wnamespace, settings_xml)

	# get stylesfrom template and write to docx styles.xml
	getNewTemplateStyles(new_rsid, template_styles_tree, styles_root, styles_xml)

	# copy the numbering.xml from the template to the docx (overwrites if present)
	shutil.copyfile(template_numbering_xml, numbering_xml)

	# update the _rels/.rels file as needed for customProps.xml handling
	edit_relsFile(rels_tree, template_rels_tree, rels_file)

	# write / overwrite override entry for customProps.xml to content_types.xml
	editContentTypes(contenttypes_tree, template_contenttypes_tree, contenttypes_xml)

	print "template attached!"


# only run if this script is being invoked directly
if __name__ == '__main__':

	attachTemplate()
	# print "docx_uptodate value is ", docx_uptodate
