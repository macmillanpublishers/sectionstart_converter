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
        report_dict = lxml_utils.logForReport(report_dict,para,report_category,runtxt)
    return report_dict

# logging each paratxt individually, this is th emost flexible for managing at time of report generation
def logTextOfParasWithStyle(report_dict, doc_root, stylename, report_category):
    logger.info("Logging paras styled as '%s' to report_dict['%s']" % (stylename, report_category))
    paras = lxml_utils.findParasWithStyle(lxml_utils.transformStylename(stylename), doc_root)
    for para in paras:
        paratxt = lxml_utils.getParaTxt(para)
        report_dict = lxml_utils.logForReport(report_dict,para,report_category,paratxt)
    return report_dict

def getAllStylesUsed(report_dict, doc_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, call_type):
    styles_tree = etree.parse(styles_xml)
    styles_root = styles_tree.getroot()
    macmillanstyle_shortnames = [lxml_utils.transformStylename(s) for s in macmillanstyledata]
    macmillan_styles_found = []
    macmillanstyles = []
    logger.info("logging 1st use of every Macmillan para style, and any use of other style")
    # get a list of macmillan stylenames from macmillan json
    for stylename in macmillanstyledata:
        macmillanstyles.append(stylename)
    for para in doc_root.findall(".//*w:p", wordnamespaces):
        # get stylename from each para
        stylename = lxml_utils.getParaStyle(para)
        if stylename not in macmillan_styles_found:

            # search styles.xlm for corresponding full stylename so we can determine if its a Macmillan style
            stylesearchstring = ".//w:style[@w:styleId='%s']/w:name" % stylename
            stylematch = styles_root.find(stylesearchstring, wordnamespaces)

            # get fullname value and test against Macmillan style list
            stylename_full = stylematch.get('{%s}val' % wnamespace)
            if stylename_full in macmillanstyles:
                if stylename not in sectionnames:
                    macmillan_styles_found.append(stylename)
                    report_dict = lxml_utils.logForReport(report_dict,para,"Macmillan_style_first_use",stylename_full)
                if stylename_full not in bookmakerstyles:
                    report_dict = lxml_utils.logForReport(report_dict,para,"non_bookmaker_macmillan_style",stylename_full)
            else:
                report_dict = lxml_utils.logForReport(report_dict,para,"non-Macmillan_style_used",stylename_full)
                # if we're "validating", revert custom_styles based on Macmillan styles to base_style
                if call_type == "validate":
                    basedon_element = stylematch.getparent().find(".//w:basedOn", wordnamespaces)
                    if basedon_element is not None:
                        basedonstyle = basedon_element.get('{%s}val' % wnamespace)
                        if basedonstyle in macmillanstyle_shortnames:
                            attrib_style_key = '{%s}val' % wnamespace
                            para.find(".//*w:pStyle", wordnamespaces).set(attrib_style_key, basedonstyle)
                            # optionally, log to json:
                            report_dict = lxml_utils.logForReport(report_dict,para,"changed_custom_style_to_Macmillan_basestyle", "'%s', based on '%s'" % (stylename_full, basedonstyle))

    # Now get runstyles!
    logger.info("logging 1st use of every Macmillan char style, and any use of other char-style")
    for run_style in doc_root.findall(".//*w:rStyle", wordnamespaces):
        # get run_stylename from each styled run
        attrib_style_key = '{%s}val' % wnamespace
        stylename = run_style.get(attrib_style_key)
        # stylename = run_style.get('{%s}val' % wnamespace)
        if stylename not in macmillan_styles_found:
            # get para for report
            para = run_style.getparent().getparent().getparent()
            # search styles.xlm for corresponding full stylename so we can determine if its a Macmillan style
            stylesearchstring = ".//w:style[@w:styleId='%s']/w:name" % stylename
            stylematch = styles_root.find(stylesearchstring, wordnamespaces)

            # get fullname value and test against Macmillan style list
            stylename_full = stylematch.get('{%s}val' % wnamespace)
            if stylename_full in macmillanstyles:
                if stylename not in sectionnames:
                    macmillan_styles_found.append(stylename)
                    report_dict = lxml_utils.logForReport(report_dict,para,"Macmillan_style_first_use",stylename_full)
            elif stylename_full != "annotation reference" and stylename_full != "endnote reference":
                report_dict = lxml_utils.logForReport(report_dict,para,"non-Macmillan_style_used",stylename_full)
                # if we're "validating", revert custom_styles based on Macmillan styles to base_style
                if call_type == "validate":
                    basedon_element = stylematch.getparent().find(".//w:basedOn", wordnamespaces)
                    if basedon_element is not None:
                        basedonstyle = basedon_element.get('{%s}val' % wnamespace)
                        if basedonstyle in macmillanstyle_shortnames:
                            run_style.set(attrib_style_key, basedonstyle)       # test this!!
                            # optionally, log to json:
                            report_dict = lxml_utils.logForReport(report_dict,para,"changed_custom_style_to_Macmillan_basestyle", "'%s', based on '%s'" % (stylename_full, basedonstyle))

    return report_dict, doc_root


def styleReports(report_dict):
    # local vars
    section_start_rules_json = cfg.section_start_rules_json
    macmillanstyles_json = cfg.macmillanstyles_json
    vbastyleconfig_json = cfg.vbastyleconfig_json
    doc_xml = cfg.doc_xml
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    styles_xml = cfg.styles_xml

    # read rules & macmillan styles from JSONs
    section_start_rules = os_utils.readJSON(section_start_rules_json)
    macmillanstyledata = os_utils.readJSON(macmillanstyles_json)
    vbastyleconfig_dict = os_utils.readJSON(vbastyleconfig_json)
    bookmakerstyles = vbastyleconfig_dict["bookmakerstyles"]

    # get Section Start names & styles from sectionstartrules
    sectionnames = lxml_utils.getAllSectionNames(section_start_rules)

    # get all Section Starts in the doc:
    report_dict = lxml_utils.sectionStartTally(report_dict, sectionnames, doc_root, "report")

    # log texts of illustation-holder paras
    report_dict = logTextOfParasWithStyle(report_dict, doc_root, cfg.illustrationholder_style, "illustration_holders")

    # log texts of titlepage-author paras
    report_dict = logTextOfParasWithStyle(report_dict, doc_root, cfg.authorstyle, "author_paras")

    # log texts of titlepage-title paras
    report_dict = logTextOfParasWithStyle(report_dict, doc_root, cfg.titlestyle, "title_paras")

    # log texts of isbn-span runs
    report_dict = logTextOfRunsWithStyle(report_dict, doc_root, cfg.isbnstyle, "isbn_spans")

    # list all styles used in the doc
    report_dict, doc_root = getAllStylesUsed(report_dict, doc_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, "report")

    # add/update para index numbers
    logger.debug("Update all report_dict records with para_index")
    report_dict = lxml_utils.calcLocationInfoForLog(report_dict, doc_root, sectionnames)

    return report_dict


#---------------------  MAIN
# only run if this script is being invoked directly
if __name__ == '__main__':

    # set up debug log to console
    logging.basicConfig(level=logging.DEBUG)

    report_dict = {}
    report_dict = styleReports(report_dict)
    logger.debug("report_dict:  %s" % report_dict)
