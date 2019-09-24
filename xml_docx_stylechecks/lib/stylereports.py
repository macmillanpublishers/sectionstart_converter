######### IMPORT PY LIBRARIES
import os
import shutil
import re
import uuid
import json
import sys
import collections
import logging
# make sure to install lxml: sudo pip install lxml
from lxml import etree


######### IMPORT LOCAL MODULES
if __name__ == '__main__':
    # to go up a level to read cfg (and other files) when invoking from this script (for testing).
    cfgpath = os.path.join(sys.path[0], '..', 'cfg.py')
    osutilspath = os.path.join(sys.path[0], '..', 'shared_utils', 'os_utils.py')
    lxmlutilspath = os.path.join(sys.path[0], '..', 'shared_utils', 'lxml_utils.py')
    import imp
    cfg = imp.load_source('cfg', cfgpath)
    os_utils = imp.load_source('os_utils', osutilspath)
    lxml_utils = imp.load_source('lxml_utils', lxmlutilspath)
else:
    import cfg
    import shared_utils.os_utils as os_utils
    import shared_utils.lxml_utils as lxml_utils

# # # to import benchmark decorator:
# decoratorspath = os.path.join(sys.path[0], '..','..','utilities','python_utils','decorators.py')
# import imp
# decorators = imp.load_source('decorators', decoratorspath)


######### LOCAL DECLARATIONS
# Local namespace vars
wnamespace = cfg.wnamespace
w14namespace = cfg.w14namespace
wordnamespaces = cfg.wordnamespaces

# initialize logger
logger = logging.getLogger(__name__)


# #---------------------  METHODS
# this is mostly identical to the function below for paras, with some varnames changed and 2 other tweaks
def logTextOfRunsWithStyle(report_dict, doc_root, stylename, report_category):
    logger.info("Logging runs styled as %s to report_dict['%s']" % (stylename, report_category))
    runs = lxml_utils.findRunsWithStyle(lxml_utils.transformStylename(stylename), doc_root)
    for run in runs:
        runtxt = lxml_utils.getParaTxt(run)
        para = run.getparent()
        report_dict = lxml_utils.logForReport(report_dict,doc_root,para,report_category,runtxt)
    return report_dict

# logging each paratxt individually, this is th emost flexible for managing at time of report generation
def logTextOfParasWithStyle(report_dict, doc_root, stylename, report_category):
    logger.info("Logging paras styled as '%s' to report_dict['%s']" % (stylename, report_category))
    paras = lxml_utils.findParasWithStyle(lxml_utils.transformStylename(stylename), doc_root)
    for para in paras:
        paratxt = lxml_utils.getParaTxt(para)
        report_dict = lxml_utils.logForReport(report_dict,doc_root,para,report_category,paratxt)
    return report_dict

def checkFirstPara(report_dict, doc_root, sectionnames, description):
    logger.info("Checking first para style to make sure it is a SectionStart..")
    firstpara = doc_root.find(".//*w:p", wordnamespaces)
    stylename = lxml_utils.getParaStyle(firstpara)
    if stylename not in sectionnames:
        logger.warn("first para style is not a required style, instead is: " + stylename)
        report_dict = lxml_utils.logForReport(report_dict,doc_root,firstpara,description,lxml_utils.getStyleLongname(stylename))
    return report_dict, firstpara

def getAllStylesUsed_RevertToBase(stylematch, macmillanstyles, report_dict, doc_root, stylename_full, para, run_style=None):
    macmillanstyle_shortnames = [lxml_utils.transformStylename(s) for s in macmillanstyles]
    basedon_element = stylematch.getparent().find(".//w:basedOn", wordnamespaces)
    if basedon_element is not None:
        basedonstyle = basedon_element.get('{%s}val' % wnamespace)
        if basedonstyle in macmillanstyle_shortnames:
            if run_style is not None:
                run_style.set(attrib_style_key, basedonstyle)
            else:
                attrib_style_key = '{%s}val' % wnamespace
                para.find(".//*w:pStyle", wordnamespaces).set(attrib_style_key, basedonstyle)
            # optionally, log to json:
            report_dict = lxml_utils.logForReport(report_dict,doc_root,para,"changed_custom_style_to_Macmillan_basestyle", "'%s', based on '%s'" % (stylename_full, basedonstyle))
        else:
            if run_style is not None:
                # log char styles
                report_dict = lxml_utils.logForReport(report_dict,doc_root,para,"non-Macmillan_charstyle_used",stylename_full)
            else:
                # log para styles not reverted to base
                report_dict = lxml_utils.logForReport(report_dict,doc_root,para,"non-Macmillan_style_used",stylename_full)
    return report_dict

def getAllStylesUsed_ProcessParaStyle(report_dict, stylename, styles_root, doc_root, macmillanstyles, sectionnames, found_para_context, container_styles, container_prefix, macmillan_styles_found_dict, macmillan_styles_found, para, call_type, bookmakerstyles):
    # search styles.xlm for corresponding full stylename so we can determine if its a Macmillan style
    stylesearchstring = ".//w:style[@w:styleId='%s']/w:name" % stylename
    stylematch = styles_root.find(stylesearchstring, wordnamespaces)

    # get fullname value and test against Macmillan style list
    stylename_full = stylematch.get('{%s}val' % wnamespace)
    if stylename_full in macmillanstyles:
        if stylename not in sectionnames and stylename not in container_styles:
            macmillan_styles_found_dict.append(found_para_context)
            macmillan_styles_found.append(stylename)
            fullstylename_with_container = container_prefix + stylename_full
            report_dict = lxml_utils.logForReport(report_dict,doc_root,para,"Macmillan_style_first_use",fullstylename_with_container)
        # skipping this check for rsuitevalidate - since it is moot. Testing by presence of container styles.
        if not container_styles:
            if stylename_full not in bookmakerstyles:
                report_dict = lxml_utils.logForReport(report_dict,doc_root,para,"non_bookmaker_macmillan_style",stylename_full)
    else:
        # if we're "validating", revert custom_styles based on Macmillan styles to base_style
        if call_type == "validate":
            report_dict = getAllStylesUsed_RevertToBase(stylematch, macmillanstyles, report_dict, doc_root, stylename_full, para)
        else:
            report_dict = lxml_utils.logForReport(report_dict,doc_root,para,"non-Macmillan_style_used",stylename_full)
    return report_dict

def getAllStylesUsed(report_dict, doc_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, call_type, valid_native_word_styles, container_starts=[], container_ends=[], runs_only=False):
    logger.info("** running function 'getAllStylesUsed'")
    styles_tree = etree.parse(styles_xml)
    styles_root = styles_tree.getroot()
    # macmillanstyle_shortnames = [lxml_utils.transformStylename(s) for s in macmillanstyledata]
    # get a list of macmillan stylenames from macmillan json, start with native word styles
    # if we want to exclude valid native word styles from report instead, would add them to conditional on line 110
    macmillanstyles = valid_native_word_styles
    for stylename in macmillanstyledata:
        macmillanstyles.append(stylename)
    macmillan_styles_found = [] # <- non-rsuite Macmillan para styles
    macmillan_styles_found_dict = []   # <- for rsuite para styles
    charstyles_found = [] # <- for all Macmillan char styles, to make sure we don't report them more than once (we are summarizing)
    # now capture / add Macmillan charstyles found in previous runs of other xml files in doc
    if "Macmillan_charstyle_first_use" in report_dict:
        for charstyle_dict in report_dict["Macmillan_charstyle_first_use"]:
            charstyles_found.append(charstyle_dict['description'])
    if "non-Macmillan_charstyle_used" in report_dict:
        for charstyle_dict in report_dict["non-Macmillan_charstyle_used"]:
            charstyles_found.append(charstyle_dict['description'])

    # adding "runs_only" option so I can re-use this to capture charstyles for footnotes/endnotes
    if runs_only == True:
        logger.info("runs_only set to: %s, we are probably scanning xml other than doc itself, just for charstyles" % runs_only)
    else:
        logger.info("logging 1st use of every Macmillan para style, and any use of other style")
        this_section = ""
        container_prefix = ""
        for para in doc_root.findall(".//*w:p", wordnamespaces):
            # get stylename from each para
            stylename = lxml_utils.getParaStyle(para)

            # track current section & container as we loop through styles
            if stylename in sectionnames:
                this_section = stylename
                container_prefix = ""
                continue
            elif stylename in container_starts:
                container_prefix = lxml_utils.getStyleLongname(stylename).split()[0] + " > "
                continue
            elif stylename in container_ends:
                container_prefix = ""
                continue

            shortstylename_with_container = container_prefix + stylename
            found_para_context = {this_section:shortstylename_with_container}

            # check index to see if style has already been noted (with section / container context where apropos)
            test_if_present = False
            if not container_starts and stylename in macmillan_styles_found:
                test_if_present = True
            elif container_starts:
                for d in macmillan_styles_found_dict:
                    if this_section in d and d[this_section] == shortstylename_with_container:
                        test_if_present = True

            # if stylename not in macmillan_styles_found, proceed to process/ log it!:
            if test_if_present == False:
                container_styles = container_starts + container_ends
                report_dict = getAllStylesUsed_ProcessParaStyle(report_dict, stylename, styles_root, doc_root, macmillanstyles, sectionnames, found_para_context, container_styles, container_prefix, macmillan_styles_found_dict, macmillan_styles_found, para, call_type, bookmakerstyles)


    # Now get runstyles!
    logger.info("logging 1st use of every Macmillan char style, and any use of other char-style")
    for run_style in doc_root.findall(".//*w:rStyle", wordnamespaces):
        # get run_stylename from each styled run
        attrib_style_key = '{%s}val' % wnamespace
        stylename = run_style.get(attrib_style_key)
        # stylename = run_style.get('{%s}val' % wnamespace)
        if stylename not in charstyles_found:
            # get para for report
            para = run_style.getparent().getparent().getparent()
            # search styles.xlm for corresponding full stylename so we can determine if its a Macmillan style
            stylesearchstring = ".//w:style[@w:styleId='%s']/w:name" % stylename
            stylematch = styles_root.find(stylesearchstring, wordnamespaces)

            # get fullname value and test against Macmillan style list
            stylename_full = stylematch.get('{%s}val' % wnamespace)
            if stylename_full in macmillanstyles and stylename_full not in charstyles_found: #we screen stylename_full here to filter results from prev runs
                charstyles_found.append(stylename)
                if stylename_full != 'annotation reference':
                    report_dict = lxml_utils.logForReport(report_dict,doc_root,para,"Macmillan_charstyle_first_use",stylename_full)
            else: #if stylename_full != "annotation reference" and stylename_full != "endnote reference":
                # if we're "validating", revert custom_styles based on Macmillan styles to base_style
                if call_type == "validate":
                    report_dict = getAllStylesUsed_RevertToBase(stylematch, macmillanstyles, report_dict, doc_root, stylename_full, para, run_style)
                elif stylename_full not in charstyles_found:
                    report_dict = lxml_utils.logForReport(report_dict,doc_root,para,"non-Macmillan_charstyle_used",stylename_full)
                    print "%s found!" % stylename_full
                    print "charstyles_found: ", charstyles_found
                    # add to the list of found charstyles:
                    charstyles_found.append(stylename)
    return report_dict

def styleReports(call_type, report_dict):
    # local vars
    # section_start_rules_json = cfg.section_start_rules_json
    # styleconfig_json = cfg.styleconfig_json
    macmillanstyles_json = cfg.macmillanstyles_json
    vbastyleconfig_json = cfg.vbastyleconfig_json
    doc_xml = cfg.doc_xml
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    styles_xml = cfg.styles_xml
    endnotes_tree = etree.parse(cfg.endnotes_xml)
    endnotes_root = endnotes_tree.getroot()
    footnotes_tree = etree.parse(cfg.footnotes_xml)
    footnotes_root = footnotes_tree.getroot()

    # read rules & macmillan styles from JSONs
    # section_start_rules = os_utils.readJSON(section_start_rules_json)
    # styleconfig_dict = os_utils.readJSON(styleconfig_json)
    macmillanstyledata = os_utils.readJSON(macmillanstyles_json)
    vbastyleconfig_dict = os_utils.readJSON(vbastyleconfig_json)
    bookmakerstyles = vbastyleconfig_dict["bookmakerstyles"]

    # get Section Start names & styles from sectionstartrules
    sectionnames = lxml_utils.getAllSectionNamesFromVSC(vbastyleconfig_dict)

    # get all Section Starts in the doc:
    report_dict = lxml_utils.sectionStartTally(report_dict, sectionnames, doc_root, "report")

    # check first para for non-section-startstyle
    report_dict, firstpara = checkFirstPara(report_dict, doc_root, sectionnames,"non_section_start_styled_firstpara")

    # log texts of illustation-holder paras
    report_dict = logTextOfParasWithStyle(report_dict, doc_root, cfg.imageholder_style, "image_holders")

    # log texts of titlepage-author paras
    report_dict = logTextOfParasWithStyle(report_dict, doc_root, cfg.authorstyle, "author_paras")

    # log texts of titlepage-title paras
    report_dict = logTextOfParasWithStyle(report_dict, doc_root, cfg.titlestyle, "title_paras")

    # log texts of titlepage-logo paras
    report_dict = logTextOfParasWithStyle(report_dict, doc_root, cfg.logostyle, "logo_paras")

    # log texts of isbn-span runs
    report_dict = logTextOfRunsWithStyle(report_dict, doc_root, cfg.isbnstyle, "isbn_spans")

    # log texts of inline illustration-holder runs
    report_dict = logTextOfRunsWithStyle(report_dict, doc_root, cfg.inline_imageholder_style, "image_holders")

    # list all styles used in the doc
    # report_dict, doc_root = getAllStylesUsed(report_dict, doc_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, call_type, valid_native_word_styles)
    report_dict = getAllStylesUsed(report_dict, doc_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, call_type, cfg.valid_native_word_styles)

    # add/update para index numbers
    logger.debug("Update all report_dict records with para_index")
    report_dict = lxml_utils.calcLocationInfoForLog(report_dict, doc_root, sectionnames, [footnotes_root, endnotes_root])

    # create sorted version of "image_holders" list in reportdict based on para_index; for reports
    if "image_holders" in report_dict:
        report_dict["image_holders__sort_by_index"] = sorted(report_dict["image_holders"], key=lambda x: x['para_index'])

    return report_dict


#---------------------  MAIN
# only run if this script is being invoked directly
if __name__ == '__main__':

    # set up debug log to console
    logging.basicConfig(level=logging.DEBUG)

    report_dict = {}
    report_dict = styleReports("report", report_dict)
    logger.debug("report_dict:  %s" % report_dict)
