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




def docPrepare(report_dict):
    logger.info("* * * commencing docPrepare function...")
    # local vars
    section_start_rules_json = cfg.section_start_rules_json
    styleconfig_json = cfg.styleconfig_json
    doc_xml = cfg.doc_xml
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    # titlestylename = cfg.titlestylename

    logger.info("reading in json resource files")
    # read rules & heading-style list from JSONs
    section_start_rules = os_utils.readJSON(section_start_rules_json)
    styleconfig_dict = os_utils.readJSON(styleconfig_json)
    headingstyles = [classname[1:] for classname in styleconfig_dict["headingparas"]]

    # get Section Start names & styles from sectionstartrules
    sectionnames = lxml_utils.getAllSectionNames(section_start_rules)

    # get all Section Starts paras in the doc, add content to each para as needed:
    report_dict = lxml_utils.sectionStartTally(report_dict, sectionnames, doc_root, "insert", headingstyles)

    # autonumber contents for chapter, Appendix, Part
    report_dict = lxml_utils.autoNumberSectionParaContent(report_dict, sectionnames, cfg.autonumber_sections, doc_root)

    # write our changes back to doc.xml
    logger.debug("writing changes out to doc_xml file")
    os_utils.writeXMLtoFile(doc_root, doc_xml)

    # add/update para index numbers
    logger.debug("Update all report_dict records with para_index-")
    report_dict = lxml_utils.calcLocationInfoForLog(report_dict, doc_root, sectionnames)

    logger.info("* * * ending docPrepare function.")

    return report_dict

#---------------------  MAIN
# only run if this script is being invoked directly
if __name__ == '__main__':

    # set up debug log to console
    logging.basicConfig(level=logging.DEBUG)

    report_dict = {}
    report_dict = docPrepare(report_dict)

    logger.debug("report_dict contents:  %s" % report_dict)
