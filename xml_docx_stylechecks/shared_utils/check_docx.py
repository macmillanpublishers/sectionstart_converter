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
    os_utils = imp.load_source('os_utils', osutilspath)
else:
    import cfg
    import shared_utils.os_utils as os_utils

######### LOCAL DECLARATIONS

# Word namespace vars
wnamespace = cfg.wnamespace
wordnamespaces = cfg.wordnamespaces
collapse_trackchange_tags = cfg.collapse_trackchange_tags
del_trackchange_tags = cfg.del_trackchange_tags
all_trackchange_tags = collapse_trackchange_tags + del_trackchange_tags

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
            version_compare = "newer_template_avail"
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
def version_test(customprops_xml, template_customprops_xml, templateversion_cutoff):

    current_version = get_docxVersion(customprops_xml)
    template_version = get_docxVersion(template_customprops_xml)
    sectionstart_version = templateversion_cutoff

    version_result = compare_docxVersions(current_version, template_version, sectionstart_version)
    return version_result, current_version, template_version

# # to test trackchanges and doc protection, and existence of any other top level param in settings
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

# to test trackchanges and doc protection, and existence of any other top level param in settings
def checkSettingsXML_Attribute(settings_xml, settingstring, attr_string):
    logger.debug("checking settings_xml for '%s', getting attribute '%s'..." % (settingstring, attr_string))
    value = ""
    # get root / tree of docx customProps.xml
    settings_tree = etree.parse(settings_xml)
    settings_root = settings_tree.getroot()

    # check for version element in docx settings.xml
    searchstring = ".//w:{}[@w:{}]".format(settingstring, attr_string)
    element = settings_root.find(searchstring, wordnamespaces)
    if element is not None:
        attr_val = element.get('{%s}%s' % (wnamespace, attr_string))
    else:
        attr_val = None
    return attr_val

def checkDocProtection(settings_xml):
    logger.debug("checking to see if protection is enabled...")
    protect_type = "no_protection"
    # enforcement is the key in w:documentProtection element that determines if a doc's protection setting is enabled
    enforcement = checkSettingsXML_Attribute(settings_xml, "documentProtection", "enforcement")
    if enforcement is not None and int(enforcement) == 1:
        # edit is the key in w:documentProtection element that says what type of protection is enabled
        protect_type = checkSettingsXML_Attribute(settings_xml, "documentProtection", "edit")

    return protect_type

def checkDocXMLforUnacceptedTrackChanges(doc_xml):
    logger.debug
    tc_marker_found = False
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    for element in all_trackchange_tags:
        searchstring = ".//*w:{}".format(element)
        for found_el in doc_root.findall(searchstring, wordnamespaces):
            tc_marker_found = True
            logger.debug("found an '%s' with searchstring" % element)
    return tc_marker_found

def getProtectionAndTrackChangesStatus(doc_xml, settings_xml):
    trackchange_status = False
    tc_marker_found = checkDocXMLforUnacceptedTrackChanges(doc_xml)
    trackRevisions = checkSettingsXML(settings_xml, "trackRevisions")
    trackRevisionsVal = checkSettingsXML_Attribute(settings_xml, "trackRevisions", "val")
    protect_type = checkDocProtection(settings_xml)
    # make sure trackRevisions is present, not set to false, and not superfluous b/c we already have TC protection as on.
    if trackRevisionsVal != "false" and trackRevisions == True and protect_type != "trackedChanges":
        trackchange_status = True
    # returns of empty or 'None' value from attribute check:
    if not protect_type:
        protection = "unspecified"
    else:
        protect_strings = {"no_protection":"","readOnly":"read-only","comments":"comments-only","trackedChanges":"track-changes-only","forms":"form-fields-only"}
        protection = protect_strings[protect_type]

    return protection, tc_marker_found, trackchange_status

def acceptTrackChanges(doc_xml):
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    # delete each tag and any nested contents!
    for tag in del_trackchange_tags:
        searchstring = ".//*w:{}".format(tag)
        for found_el in doc_root.findall(searchstring, wordnamespaces):
            parent_el = found_el.getparent()
            parent_el.remove(found_el)
            # see if parent_el is a p without any child runs, if so let's scrap that too
            if parent_el.tag == "{%s}p" % wnamespace:
                run_check = parent_el.find(".//*w:r", wordnamespaces)
                run_checkB = parent_el.find(".//w:r", wordnamespaces)
                if run_check is None and run_checkB is None:
                    parent_el.getparent().remove(parent_el)
            else:
                print parent_el.tag, "{%s}p" % wnamespace

    # 'collapse' parent element, raising child elements (where there are any printing contents), then delete parent element
    for tag in collapse_trackchange_tags:
        searchstring = ".//*w:{}".format(tag)
        for found_el in doc_root.findall(searchstring, wordnamespaces):
            # print found_el    # debug
            # see if we have any children to move up a level
            run_check = found_el.find(".//w:r", wordnamespaces)
            run_checkB = found_el.find(".//*w:r", wordnamespaces)
            if run_check is not None or run_checkB is not None:
                # print "babies present"    # debug
                child_els = found_el.getchildren()
                # move children out of collapse_tag_el
                for child in list(child_els):
                    found_el.addprevious(child)
            # get rid of our old parent
            found_el.getparent().remove(found_el)
    return doc_root

def stripDuplicateMacmillanStyles(doc_xml, styles_xml):
    logger.debug("Checking for old Macmillan duplicate styles, replacing as needed...")
    zerostylecheck = False

    styles_tree = etree.parse(styles_xml)
    styles_root = styles_tree.getroot()
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    # this is for cycling through any of these xml_roots that exist
    xmlfile_dict = {doc_root:doc_xml}
    if os.path.exists(cfg.endnotes_xml):
        endnotes_tree = etree.parse(cfg.endnotes_xml)
        endnotes_root = endnotes_tree.getroot()
        xmlfile_dict[endnotes_root]=cfg.endnotes_xml
    if os.path.exists(cfg.footnotes_xml):
        footnotes_tree = etree.parse(cfg.footnotes_xml)
        footnotes_root = footnotes_tree.getroot()
        xmlfile_dict[footnotes_root]=cfg.footnotes_xml
    if os.path.exists(cfg.numbering_xml):
        numbering_tree = etree.parse(cfg.numbering_xml)
        numbering_root = numbering_tree.getroot()
        xmlfile_dict[numbering_root]=cfg.numbering_xml

    # get styles that end in zero: use xpath b/c lxml find doesn't support wildcard search in attr name
    stylematches = styles_tree.xpath("w:style[@w:styleId[contains(.,'0') and substring-after(.,'0') = '']]", namespaces=wordnamespaces)
    # print "stylematch count!!!!!!!:  %s" % len(stylematches) # debug
    if len(stylematches):
        for zerostyle in stylematches:
            # only capture macmillan styles
            if "(" in zerostyle.find(".//w:name", wordnamespaces).get('{%s}val' % wnamespace):
                zerostylecheck = True
                zerostylename = zerostyle.get('{%s}styleId' % wnamespace)
                nozerostylename = zerostylename[:-1]
                nozero_downcase_stylename = nozerostylename.lower()

                # # find style with matching stylename. (legacystyle), get its stylename, rm legacy style
                legacystyle = styles_tree.xpath("w:style[translate(@w:styleId,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = '" + nozero_downcase_stylename + "']", namespaces=wordnamespaces)
                legacystylename = legacystyle[0].get('{%s}styleId' % wnamespace)
                parent_el = legacystyle[0].getparent()
                parent_el.remove(legacystyle[0])
                # go back to zero style & correct its stylename (sans 0)
                zerostyle.attrib["{%s}styleId" % wnamespace] = nozerostylename

                # cycle through xml files to update relevant stylenames
                for xml_root in xmlfile_dict:
                    # search xml file for all references to legacystyle
                    legacystyle_searchstring = ".//w:pStyle[@w:val='%s']" % legacystylename
                    legacy_uses = xml_root.findall(legacystyle_searchstring, wordnamespaces)
                    zerostyle_searchstring = ".//w:pStyle[@w:val='%s']" % zerostylename
                    zerostyle_uses = xml_root.findall(zerostyle_searchstring, wordnamespaces)
                    # and update all style references to zerostyle & legacy
                    for changedstyle_use in zerostyle_uses + legacy_uses:
                        changedstyle_use.attrib["{%s}val" % wnamespace] = nozerostylename

    if zerostylecheck:  # write out updated xml files if zerostyles were found
        # write xml out. styles file:
        os_utils.writeXMLtoFile(styles_root, styles_xml)
        #   .. and all other xml files:
        for xml_root in xmlfile_dict:
            os_utils.writeXMLtoFile(xml_root, xmlfile_dict[xml_root])

    return zerostylecheck

#---------------------  MAIN

# # only run if this script is being invoked directly
# if __name__ == '__main__':
#     # set up debug log to console
#     logging.basicConfig(level=logging.DEBUG)
#     settings_xml = cfg.settings_xml
#     doc_xml = cfg.doc_xml
#     doc_tree = etree.parse(doc_xml)
#     doc_root = doc_tree.getroot()
#     import os_utils
# #
# # Testing 2
#     # protection, tc_marker_found, trackchange_status = getProtectionAndTrackChangesStatus(doc_xml, settings_xml)
#     # logger.debug("protection: '%s'" % protection)
#     # logger.debug("tc_marker_found: '%s'" % tc_marker_found)
#     # logger.debug("trackchange_status: '%s'" % trackchange_status)
#
#     doc_root = acceptTrackChanges(doc_xml)
#     os_utils.writeXMLtoFile(doc_root, doc_xml)
