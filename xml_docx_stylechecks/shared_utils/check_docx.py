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

    # need to use imp to import files from outside the dir
    import imp
    cfgpath = os.path.join(sys.path[0], '..', 'cfg.py')
    setup_cleanuppath = os.path.join(sys.path[0], '..', 'lib', 'setup_cleanup.py')
    usertext_templatespath = os.path.join(sys.path[0], '..', 'lib', 'usertext_templates.py')
    cfg = imp.load_source('cfg', cfgpath)
    setup_cleanup = imp.load_source('setup_cleanup', setup_cleanuppath)
    usertext_templates = imp.load_source('setup_cleanup', usertext_templatespath)
    import os_utils
    import lxml_utils
    import sendmail
else:
    import cfg
    import shared_utils.sendmail as sendmail
    import shared_utils.os_utils as os_utils
    import shared_utils.lxml_utils as lxml_utils
    import lib.setup_cleanup as setup_cleanup
    import lib.usertext_templates as usertext_templates

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

def checkFilenameChars(inputfilename):
    logger.debug("Verifying valid filename")
    filename_regex = re.compile(r"[^\w-]")
    fname_noext = os.path.splitext(inputfilename)[0]
    badchars = re.findall(filename_regex, fname_noext)
    return badchars

def filenameChecks(inputfilename):
    logger.info("Verifying valid filename (Checking fname length, if bad-characters are present)...")
    fname_check = True
    fname_toolong = False
    fname_maxlength = cfg.filename_maxlength
    # check for badchars
    badchars = checkFilenameChars(inputfilename)
    #  log and write alert as needed
    if badchars and not os.environ.get('TEST_FLAG'): # < skip for unittests
        setup_cleanup.setAlert('warning', 'bad_filename', {'badchar_array': badchars})
    # check / report on filename length, alert as needed
    if len(inputfilename) > fname_maxlength:
        fname_toolong = True
        if not os.environ.get('TEST_FLAG'): # < skip for unittests
            setup_cleanup.setAlert('error', 'fname_toolong', {'filename':inputfilename, 'fname_length':len(inputfilename), 'max_length':fname_maxlength})
    # we only return false for fname too long; to denote that we cannot process this file
    if fname_toolong == True:
        fname_check = False
    return fname_check

# compares namespace url for first element found using passed parameter's local-name
#   Currently in use to check ns-url of 'body' tag.
def compareElementNamespace(this_xmlfile, template_xmlfile, element_name, ns_reqrd=True):
    logger.debug("running compareElementNamespace for '{}' element...".format(element_name))

    # get _expected_ ns_url for body element
    template_xmltree = etree.parse(template_xmlfile)
    # # \/ use this line instead of the following one to only find element at root of document
    # template_element_tag = xmltree.xpath("/*[local-name()='document']/*[local-name()='body']")[0].tag
    template_element_tag = template_xmltree.xpath("//*[local-name()='{}']".format(element_name))[0].tag
    template_nsurl = template_element_tag.split('}')[0].replace('{', '')

    # get this document's ns_url for body element
    this_nsurl = ''
    this_xmltree = etree.parse(this_xmlfile)
    # # \/ use these 2 lines instead of the following pair to only find element at root of document
    # if this_xmltree.xpath("/*[local-name()='document']/*[local-name()='{}']".format(element_name)):
    #     this_element_tag = this_xmltree.xpath("/*[local-name()='document']/*[local-name()='{}']".format(element_name))[0].tag
    if this_xmltree.xpath("//*[local-name()='{}']".format(element_name)):
        this_element_tag = this_xmltree.xpath("//*[local-name()='{}']".format(element_name))[0].tag
        this_nsurl = this_element_tag.split('}')[0].replace('{', '')

    if this_nsurl:
        if this_nsurl == template_nsurl:
            return 'expected'
        else:
            return this_nsurl
    elif ns_reqrd == True:
        logger.error('required element "{}" missing during checkdocx.compareElementNamespace; raising exception'.format(element_name))
        raise Exception('element "{}" not present'.format(element_name))
    else:
        return 'unavailable'

def checkRqrdNamespace(xml_file):
    logger.info('verifying that this "body" namespace url in document.xml matches our template')
    namespace_check = True
    ns_url = compareElementNamespace(cfg.doc_xml, cfg.template_document_xml, 'body')
    # log alert as needed
    if ns_url != 'expected' and ns_url != 'unavailable':
        namespace_check = False
        setup_cleanup.setAlert('error', 'unexpected_namespace')
        # send notification to wf team so we can keep an eye on these
        subject = usertext_templates.subjects()["ns_notify"]
        bodytxt = usertext_templates.emailtxt()["ns_notify"].format(inputfilename=cfg.inputfilename, ns_url=ns_url,
            submitter=cfg.submitter_email, alert_text=usertext_templates.alerts()['unexpected_namespace'],
            support_email_address=cfg.support_email_address, helpurl=cfg.helpurl)
        sendmail.sendMail([cfg.alert_email_address], subject, bodytxt)
    return namespace_check

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

        # if versionstring is like 'x.x.x' just get 'x.x'
        if len(versionstring.split('.')) > 2:
            versionstring = '.'.join(versionstring.split('.')[:2])

        try:
            versionstring = str(float(versionstring))
        except:
            logging.error("version_string from doc props cannot be converted to float: '%s'. Reverting to 0.0" % versionstring)
            versionstring = "0.0"

    logger.debug("versionstring value: '%s'" %  versionstring)
    return versionstring

def compare_docxVersions(document_version, template_version, doc_version_min, doc_version_max):
    logger.debug("comparing docx version to template...")
    # adding try statement since the Version library can be a little particular.
    try:
        if document_version == "0.0":
            version_compare = "no_version"
        elif doc_version_min is not None and Version(document_version) < Version(doc_version_min):
            version_compare = "docversion_below_minimum"
        elif Version(document_version) < Version(template_version):
            version_compare = "newer_template_avail"
        elif doc_version_max is not None and Version(document_version) >= Version(doc_version_max):
            version_compare = "docversion_above_maximum"
        elif Version(document_version) >= Version(template_version):
            version_compare = "up_to_date"

        logger.info("version_compare value: '%s'" %  version_compare)
        return version_compare
    except Exception, e:
        logger.error('Failed version compare, exiting', exc_info=True)
        sys.exit(1)

def checkIfMacmillanStyle(stylename, template_styles_root):
    is_macmillan_style = False
    # search template_styles.xlm for corresponding full stylename
    stylesearchstring = ".//w:style[@w:styleId='%s']/w:name" % stylename
    stylematch = template_styles_root.find(stylesearchstring, wordnamespaces)

    if stylematch is not None:
        # get fullname value and test for parentheses: if parenthesis we have a Macmillan style
        #   (with exception of built-in style 'Normal (Web)')
        stylename_full = stylematch.get('{%s}val' % wnamespace)
        if '(' in stylename_full and stylename_full != 'Normal (Web)':
            is_macmillan_style = True

    return is_macmillan_style


def getParaStyleSummary(xml_file, template_styles_root, valid_native_pstyles, decommissioned_styles, total_paras=0, macmillan_styled_paras=0):
    xml_tree = etree.parse(xml_file)
    total_paras += len(xml_tree.xpath(".//w:p", namespaces=wordnamespaces))
    # subtract tablecell paras from total
    tablecell_paras = len(xml_tree.xpath(".//w:tc/w:p", namespaces=wordnamespaces))
    total_paras = total_paras - tablecell_paras
    xml_root = xml_tree.getroot()
    for para_style in xml_root.findall(".//*w:pStyle", wordnamespaces):
        # get stylename from each para
        stylename = para_style.get('{%s}val' % wnamespace)
        # lets not count valid native styles as 'styled' or 'not-styled'
        if stylename in valid_native_pstyles:
            total_paras -= 1
        elif stylename in decommissioned_styles or checkIfMacmillanStyle(stylename, template_styles_root):
            macmillan_styled_paras += 1

    # now subtract macmillan styled table cell paras
    for para_style in xml_root.findall(".//w:tc//w:pStyle", wordnamespaces):
        stylename = para_style.get('{%s}val' % wnamespace)
        if stylename in decommissioned_styles or checkIfMacmillanStyle(stylename, template_styles_root):
            macmillan_styled_paras -= 1
    return total_paras, macmillan_styled_paras

def macmillanStyleCount(doc_xml, template_styles_xml):
    logger.debug("Counting total paras, Macmillan styled paras...")

    template_styles_tree = etree.parse(template_styles_xml)
    template_styles_root = template_styles_tree.getroot()
    valid_native_pstyles = [cfg.footnotestyle, cfg.endnotestyle]
    # get decommissioned styles
    styleconfig_legacy_list = os_utils.readJSON(cfg.styleconfig_json)['legacy']
    decommissioned_styles = [x[1:] for x in styleconfig_legacy_list]

    # main document
    total_paras, macmillan_styled_paras = getParaStyleSummary(doc_xml, template_styles_root, valid_native_pstyles, decommissioned_styles)
    # footnotes
    if os.path.exists(cfg.footnotes_xml):
        total_paras, macmillan_styled_paras = getParaStyleSummary(cfg.footnotes_xml, template_styles_root, valid_native_pstyles, decommissioned_styles, total_paras, macmillan_styled_paras)
        total_paras -= 2 # for built-in separator paras
    # endnotes
    if os.path.exists(cfg.endnotes_xml):
        total_paras, macmillan_styled_paras = getParaStyleSummary(cfg.endnotes_xml, template_styles_root, valid_native_pstyles, decommissioned_styles, total_paras, macmillan_styled_paras)
        total_paras -= 2 # for built-in separator paras

    # the multiplying by a factor with '.0' in the numerator forces the result to be a float for python 2.x
    percent_styled = (macmillan_styled_paras * 100.0) / total_paras

    logger.info("total paras:'%s', macmillan styled:'%s', percent_styled:'%s'" % (total_paras, macmillan_styled_paras, percent_styled))
    return percent_styled, macmillan_styled_paras, total_paras

# This function consolidates version test functions
def version_test(customprops_xml, template_customprops_xml, doc_version_min, doc_version_max):

    document_version = get_docxVersion(customprops_xml)
    template_version = get_docxVersion(template_customprops_xml)
    # sectionstart_version = doc_version_min

    version_result = compare_docxVersions(document_version, template_version, doc_version_min, doc_version_max)
    return version_result, document_version, template_version

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

def checkDocXMLforUnacceptedTrackChanges(xml_file):
    logger.debug("checking for unaccepted track changes...")
    tc_marker_found = False
    doc_tree = etree.parse(xml_file)
    doc_root = doc_tree.getroot()
    for element in all_trackchange_tags:
        searchstring = ".//*w:{}".format(element)
        for found_el in doc_root.findall(searchstring, wordnamespaces):
            tc_marker_found = True
            logger.debug("found an '%s' with searchstring" % element)
    return tc_marker_found

def getProtectionAndTrackChangesStatus(doc_xml, settings_xml, footnotes_xml, endnotes_xml):
    trackchange_status, tc_marker_found, tc_marker_footnotes, tc_marker_endnotes = False, False, False, False
    tc_marker_doc = checkDocXMLforUnacceptedTrackChanges(doc_xml)
    if os.path.exists(footnotes_xml):
        tc_marker_footnotes = checkDocXMLforUnacceptedTrackChanges(footnotes_xml)
    if os.path.exists(endnotes_xml):
        tc_marker_endnotes = checkDocXMLforUnacceptedTrackChanges(endnotes_xml)
    logger.debug("tc_marker_doc: %s, tc_marker_footnotes: %s, tc_marker_endnotes: %s" %(tc_marker_doc, tc_marker_footnotes, tc_marker_endnotes))
    trackRevisions = checkSettingsXML(settings_xml, "trackRevisions")
    trackRevisionsVal = checkSettingsXML_Attribute(settings_xml, "trackRevisions", "val")
    protect_type = checkDocProtection(settings_xml)
    # determine whether we have tc_markers
    if tc_marker_doc == True or tc_marker_footnotes == True or tc_marker_endnotes == True:
        tc_marker_found = True
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

def acceptTrackChanges(xml_file):
    logger.debug("accepting tracked changes...")
    xml_tree = etree.parse(xml_file)
    xml_root = xml_tree.getroot()
    # delete each tag and any nested contents!
    for tag in del_trackchange_tags:
        searchstring = ".//*w:{}".format(tag)
        for found_el in xml_root.findall(searchstring, wordnamespaces):
            logger.debug("found del_trackchange_tag!")
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
        for found_el in xml_root.findall(searchstring, wordnamespaces):
            logger.debug("found collapse_trackchange_tag!")
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

    os_utils.writeXMLtoFile(xml_root, xml_file)

def getXMLroot(xml_file):
    if os.path.exists(xml_file):
        xml_tree = etree.parse(xml_file)
        xml_root = xml_tree.getroot()
    else:
        xml_root = None
    return xml_root

def updateStyleidInXML(oldstyle_id, newstyle_id, xml_root, root_name):
    updates_made = False
    # search xml file for all references to legacystyle
    oldstyle_searchstring = ".//*w:pStyle[@w:val='%s']" % oldstyle_id
    oldstyle_uses = xml_root.findall(oldstyle_searchstring, wordnamespaces)
    # if we don't find paraStyles, try runstyles:
    if not oldstyle_uses:
        oldstyle_searchstring = ".//*w:rStyle[@w:val='%s']" % oldstyle_id
        oldstyle_uses = xml_root.findall(oldstyle_searchstring, wordnamespaces)
    for oldstyle_use in oldstyle_uses:
        oldstyle_use.attrib["{%s}val" % wnamespace] = newstyle_id
    if oldstyle_uses:
        updates_made = True
        logging.info("replaced %s occurrences of styleID '%s' with '%s' in %s" % (len(oldstyle_uses), oldstyle_id, newstyle_id, root_name))
        # os_utils.writeXMLtoFile(xml_root, xml_file)
    return updates_made

def updateStyleUsesInStylesXML(styles_root, current_id, new_id):
    # scanstyles_root for other common uses of updated styleID (in basedOn or nextStyle vals)
    searchstring = ".//w:style/w:next[@w:val='%s']" % current_id
    searchstring_b = ".//w:style/w:basedOn[@w:val='%s']" % current_id
    el_attrs_to_update = styles_root.findall(searchstring, wordnamespaces) + styles_root.findall(searchstring_b, wordnamespaces)
    if len(el_attrs_to_update) > 0:
        for el in el_attrs_to_update:
            el.attrib["{%s}val" % wnamespace] = new_id
        logging.debug("updated %s style-element next/basedOn Style attribute values from '%s' to '%s'" % (len(el_attrs_to_update), current_id, new_id))

def updateStyleidInAllXML(oldstyle_id, newstyle_id, styles_root, doc_root, endnotes_root, footnotes_root, xmls_updated):
    # update styles_root,
    updateStyleUsesInStylesXML(styles_root, oldstyle_id, newstyle_id)
    # doc root,
    docxml_updated = updateStyleidInXML(oldstyle_id, newstyle_id, doc_root, "document.xml")
    if xmls_updated['docxml'] == False and docxml_updated == True:
        xmls_updated['docxml'] = True
    # endnotes,
    if endnotes_root is not None:
        enotexml_updated = updateStyleidInXML(oldstyle_id, newstyle_id, endnotes_root, "endnotes.xml")
        if xmls_updated['endnotes'] == False and enotexml_updated == True:
            xmls_updated['endnotes'] = True
    # footnotes
    if footnotes_root is not None:
        fnotexml_updated = updateStyleidInXML(oldstyle_id, newstyle_id, footnotes_root, "footnotes.xml")
        if xmls_updated['footnotes'] == False and fnotexml_updated == True:
            xmls_updated['footnotes'] = True
    return xmls_updated

# go through and check that each style longname, if present, has teh expected styleid.
#   5 different outcomes: -styleID good, -style not present, -styleID wrong but no conflicting styleID,
#       -styleID wrong and conflicts with legacy style, -styleID wrong ands conflicts with random style
def verifyStyleIDs(macmillanstyle_dict, legacystyle_dict, styles_root, doc_root, endnotes_root, footnotes_root):
    stylenames_updated = False
    xmls_updated = {'docxml': False, 'footnotes': False, 'endnotes': False}
    for lng_stylename, shrt_stylename in macmillanstyle_dict.iteritems():
        searchstring = ".//w:style/w:name[@w:val='%s']" % lng_stylename
        rsuite_name_el = styles_root.find(searchstring, wordnamespaces)
        # proceed if we found our style long name in current styles.xml (do nothing)
        if rsuite_name_el is not None:
            current_rs_style_el = rsuite_name_el.getparent()
            current_style_id = current_rs_style_el.get('{%s}styleId' % wnamespace)
            # the style shortname is not the expected / calculated one, we need to take some action:
            if current_style_id != shrt_stylename:
                # find if another style is using what should be our shortname... (using xpath for case-insensitive search)
                shrt_stylename_downcase = shrt_stylename.lower()
                styles_tree = etree.ElementTree(styles_root)
                target_rs_style_els = styles_tree.xpath("w:style[translate(@w:styleId,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz') = '" \
                    + shrt_stylename_downcase + "']", namespaces=wordnamespaces)
                # if not, rename the style shortname, and rename all occurences in body xml, notes xmls to match new shortname
                if len(target_rs_style_els) == 0:
                    logging.warn("style '%s' had unexpected styleID '%s', expected styleID available ('%s'), updating all xml" \
                        % (lng_stylename, current_style_id, shrt_stylename))
                    current_rs_style_el.attrib["{%s}styleId" % wnamespace] = shrt_stylename
                    xmls_updated = updateStyleidInAllXML(current_style_id, shrt_stylename, styles_root, doc_root, endnotes_root, footnotes_root, xmls_updated)
                # if so, find out if the style with our shortname is a legacy style
                else:
                    # the above xpath search returned a list, but if we're here we found exactly 1 item
                    target_rs_style_el = target_rs_style_els[0]
                    # get target style styleID; we'll need it for legacy or non-legacy cases below:
                    target_styleID = target_rs_style_el.get('{%s}styleId' % wnamespace)
                    # get target_style full stylename:
                    target_name_el = target_rs_style_el.find(".//w:name", wordnamespaces)
                    target_stylename = target_name_el.get('{%s}val' % wnamespace)
                    # if this _is_ a legacy style, we merge the new with the old
                    #   (since the longnames are almost identical, old-style uses can be presumed to be unintentional misstylings, and corrected)
                    if target_stylename in legacystyle_dict:
                        logging.warn("style '%s' had unexpected styleID '%s', expected styleID in use by legacy style ('%s'), merging the two under styleid: '%s'" \
                            % (lng_stylename, current_style_id, target_stylename, shrt_stylename))
                        # delete the legacy style_el ...
                        parent_el = target_rs_style_el.getparent()
                        parent_el.remove(target_rs_style_el)
                        # ... and rename our corresponding RS style_el with the expected styleID
                        current_rs_style_el.attrib["{%s}styleId" % wnamespace] = shrt_stylename
                        # then we update current styleid uses in the xml to new styleid
                        xmls_updated = updateStyleidInAllXML(current_style_id, shrt_stylename, styles_root, doc_root, endnotes_root, footnotes_root, xmls_updated)
                        # if the target styleID does not match styleShortname (caps could vary) we rename that globally as well:
                        if target_styleID != shrt_stylename:
                            xmls_updated = updateStyleidInAllXML(target_styleID, shrt_stylename, styles_root, doc_root, endnotes_root, footnotes_root, xmls_updated)
                    # this is not a legacy style, so we will swap style id's in styles_root, and update occurrences with new style id's in 3 xml files
                    else:
                        # get current shortname suffix, and set new shortname for rogue dupe (in case basename capitalization varies)
                        current_suffix = re.sub(r"^{}".format(shrt_stylename),'',current_style_id)
                        if current_suffix != current_style_id:
                            new_target_style_id = '{}{}'.format(target_styleID, current_suffix)
                        else:
                            new_target_style_id = current_style_id
                        logging.warn("style '%s' had unexpected styleID '%s', expected styleID ('%s') used by non-legacy style: ('%s'), swapping style ID's in all XML" \
                            % (lng_stylename, current_style_id, shrt_stylename, target_stylename))
                        # swap styleid's in styles_root
                        current_rs_style_el.attrib["{%s}styleId" % wnamespace] = shrt_stylename
                        target_rs_style_el.attrib["{%s}styleId" % wnamespace] = new_target_style_id
                        # 3 part swap in xml roots (A > C, B > A, C > B)
                        tmpstyle_id = lxml_utils.generate_id()
                        tmp_stylename = "{}{}".format(shrt_stylename, tmpstyle_id)
                        xmls_updated = updateStyleidInAllXML(target_styleID, tmp_stylename, styles_root, doc_root, endnotes_root, footnotes_root, xmls_updated)
                        xmls_updated = updateStyleidInAllXML(current_style_id, shrt_stylename, styles_root, doc_root, endnotes_root, footnotes_root, xmls_updated)
                        xmls_updated = updateStyleidInAllXML(tmp_stylename, new_target_style_id, styles_root, doc_root, endnotes_root, footnotes_root, xmls_updated)
                stylenames_updated = True
        else:
            logger.info("style longname not present: %s (style may have been deleted/modified?)" % lng_stylename)

    return stylenames_updated, xmls_updated

# a problem caused by having clashing style templates added, or rogue random styles pre-existing a Macmillan one.
def checkForDuplicateStyleIDs(macmillanstyles_json, legacystyles_json, styles_xml, doc_xml, endnotes_xml, footnotes_xml):
    logger.info("running checkDuplicateStyleIDs...")
    # create dict of style_longname : style_shortname for all macmillan styles
    macmillanstyle_data = os_utils.readJSON(macmillanstyles_json)
    macmillanstyle_dict = {}
    for lng_stylename in macmillanstyle_data:
        macmillanstyle_dict[lng_stylename] = lxml_utils.transformStylename(lng_stylename)

    # get xml_roots
    doc_root = getXMLroot(doc_xml)
    footnotes_root = getXMLroot(footnotes_xml)
    endnotes_root = getXMLroot(endnotes_xml)
    styles_root = getXMLroot(styles_xml)

    legacystyle_dict = os_utils.readJSON(legacystyles_json)
    stylenames_updated, xmls_updated = verifyStyleIDs(macmillanstyle_dict, legacystyle_dict, styles_root, doc_root, endnotes_root, footnotes_root)
    # write styles file out
    if stylenames_updated == True:
        os_utils.writeXMLtoFile(styles_root, styles_xml)
    # if changes were made to other files, write files out
    if xmls_updated['docxml'] == True:
        os_utils.writeXMLtoFile(doc_root, doc_xml)
    if xmls_updated['footnotes'] == True:
        os_utils.writeXMLtoFile(footnotes_root, footnotes_xml)
    if xmls_updated['endnotes'] == True:
        os_utils.writeXMLtoFile(endnotes_root, endnotes_xml)


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
