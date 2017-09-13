######### IMPORT SOME STANDARD PY LIBRARIES

import sys
import os
import shutil
import re
import json
import logging
from lxml import etree

# ######### IMPORT LOCAL MODULES
import cfg


# Local namespace vars
wnamespace = cfg.wnamespace
w14namespace = cfg.w14namespace
wordnamespaces = cfg.wordnamespaces

# initialize logger
logger = logging.getLogger(__name__)


######### METHODS
def transformStylename(stylename):
    # in js we needed to escape pound signs.  Come back and test that here
    # yep, cause Word strips em out for the style-shortname
    stylename = stylename.replace(" ","").replace("(",'').replace(")",'').replace("#",'')
    return stylename

# return all text from a paragraph (or run)
def getParaTxt(para):
    try:
        paratext = "".join([x for x in para.itertext()])
    except:
        paratext = "n-a"
    return paratext

# return the w14:paraId attribute's value for a paragraph
def getParaId(para):
    # if len(para):
    if para is not None:
        attrib_id_key = '{%s}paraId' % w14namespace
        para_id = para.get(attrib_id_key)
    else:
        para_id = 'n-a' # this could happen if we are trying to get id of a prev or next para that does not exist
    return para_id

def getParaStyle(para):      # move to lxml_utils?
    try:
        pstyle = para.find(".//*w:pStyle", wordnamespaces)
        if pstyle is not None:
            stylename = pstyle.get('{%s}val' % wnamespace)
        else:
            stylename = "Normal"    # Default paragraph style
    except:
        stylename = ""
    return stylename

# return the index value of a paragraph (within the body/root)
def getParaIndex(para):
    if para is not None:
        para_index = para.getparent().index(para)
    else:
        para_index = 'n-a' # this could happen if we are trying to get id of a prev or next para that does not exist
    return para_index

# return a dict of neighboring para elements and their text
def getNeighborParas(para):          # move to lxml_utils?
    pneighbors = {}
    try:
        # the 'len' call is what generates the error and kicks to the except statement, I think?
        pneighbors['prev'] = para.getprevious()
        len(pneighbors['prev'].tag)
        pneighbors['prevtext'] = getParaTxt(pneighbors['prev'])
        pneighbors['prevstyle'] = getParaStyle(pneighbors['prev'])
    except:
        pneighbors['prev'] = ""
        pneighbors['prevtext'] = ""
        pneighbors['prevstyle'] = ""
    try:
        # the 'len' call is what generates the error and kicks to the except statement, I think?
        pneighbors['next'] = para.getnext()
        len(pneighbors['next'].tag)
        pneighbors['nexttext'] = getParaTxt(pneighbors['next'])
        pneighbors['nextstyle'] = getParaStyle(pneighbors['next'])
    except:
        pneighbors['next'] = ""
        pneighbors['nexttext'] = ""
        pneighbors['nextstyle'] = ""
    return pneighbors

# creating a dict for secitonnames: SectionStart stylenames as keys, their 'longnames' as values
def getAllSectionNames(section_start_rules):
    sectionnames = {}
    for sectionname in section_start_rules:
        sectionnames[transformStylename(sectionname)] = sectionname
    return sectionnames

# return the last SectionStart para's content
def getSectionName(para, sectionnames):
    if para is not None:
        tmp_para = para
        stylename = getParaStyle(tmp_para)
        while stylename and stylename not in sectionnames:
            pneighbors = getNeighborParas(tmp_para)
            tmp_para = pneighbors['prev']
            stylename = pneighbors['prevstyle']
        if stylename in sectionnames:
            sectionpara_name = stylename
            sectionpara_contents = getParaTxt(tmp_para)
        else:
            sectionpara_name = 'n-a' # this could happen if we are trying to get id of a prev or next para that does not exist
            sectionpara_contents = 'n-a'
    else:
        sectionpara_name = 'n-a' # this could happen if we are trying to get id of a prev or next para that does not exist
        sectionpara_contents = 'n-a'
    return sectionpara_name, sectionpara_contents

def getContentsForSectionStart(sectionbegin_para, doc_root, headingstyles, sectionname, sectionnames):
    # get para style. If not in heading styles or sectiontypes["all"], get next (while)
    tmp_para = sectionbegin_para
    stylename = getParaStyle(tmp_para)
    while stylename and stylename not in headingstyles and stylename not in sectionnames:
        pneighbors = getNeighborParas(tmp_para)
        tmp_para = pneighbors['next']
        stylename = pneighbors['nextstyle']
    # set content equal to the content of the first heading-styled para in the section
    if stylename in headingstyles:
        newcontent = getParaTxt(tmp_para).strip()
    else:   # there was no heading-styled para
        sectionLongName = sectionnames[sectionname]
        sectionshortname = sectionLongName[8:].split()[0]
        newcontent = sectionshortname
    # newcontent = "HALLOOOOOO"  (debug)
    return newcontent

def sectionStartTally(report_dict, sectionnames, doc_root, call_type, headingstyles = []):
    logger.info("writing all paras with SectionStart styles to report_dict")
    for sectionname in sectionnames:
        paras = findParasWithStyle(sectionname, doc_root)
        for para in paras:
            # log the section start para
            #   (we can run this before content is added to paras, b/c that content is captured later in the 'calcLocationInfoForLog' method
            report_dict = logForReport(report_dict, para, "section_start_found", sectionname)
            # check to see ifthe para is empty (no contents) and if so log it, and, if 'call_type' is insert, fix it.
            if not getParaTxt(para).strip():
                report_dict = logForReport(report_dict,para,"empty_section_start_para",sectionname)
                if call_type == "insert":
                    # delete any existing run(s) that may contain whitespace items.
                    runs = para.findall(".//w:r", wordnamespaces)
                    for run in runs:
                        run.getparent().remove(run)
                    # find / create contents for Section start para
                    pneighbors = getNeighborParas(para)
                    content = getContentsForSectionStart(pneighbors['next'], doc_root, headingstyles, sectionname, sectionnames)
                    # create new run element with new content & append to our para!
                    new_para_run = etree.Element("{%s}r" % wnamespace)
                    new_para_run_text = etree.Element("{%s}t" % wnamespace)
                    new_para_run_text.text = content
                    new_para_run.append(new_para_run_text)
                    para.append(new_para_run)
    return report_dict

# a method to log paragraph id for style report etc
def logForReport(report_dict,para,category,description):
    para_dict = {}
    para_dict["para_id"] = getParaId(para)
    para_dict["description"] = description
    if category not in report_dict:
        report_dict[category] = []

    report_dict[category].append(para_dict.copy())

    return report_dict

def findParasWithStyle(stylename, doc_root):
    paras = []
    searchstring = ".//*w:pStyle[@w:val='%s']" % stylename
    for pstyle in doc_root.findall(searchstring, wordnamespaces):
        para = pstyle.getparent().getparent()
        paras.append(para)
    return paras

# once all changes havebeen made, call this to add location info for users to the changelog dicts
def calcLocationInfoForLog(report_dict, root, sectionnames):
    logger.info("calculating para_index numbers for all para_ids in 'report_dict'")
    try:
        # make sure we have contents in the dict
        if report_dict:
            for category, entries in report_dict.iteritems():
                for entry in entries:
                    for key in entry.keys():
                        if key == "para_id":
                            # Get the para object
                            searchstring = ".//*w:p[@w14:paraId='%s']" % entry[key]
                            para = root.find(searchstring, wordnamespaces)
                            # # # Get para index
                            entry['para_index'] = getParaIndex(para)
                            if entry['para_index'] == 'n-a':
                                logger.warn("couldn't get para-index for %s para (value was set to n-a)" % category)
                            # # # Get Section Name
                            entry['parent_section_start_type'], entry['parent_section_start_content']  = getSectionName(para, sectionnames)
                            if entry['parent_section_start_type'] == 'n-a' or entry['parent_section_start_content'] == 'n-a':
                                logger.warn("couldn't get section start info for %s para (value was set to n-a)" % category)
                            # # # Get 1st 10 words of para text
                            entry['para_string'] = ' '.join(getParaTxt(para).split(' ')[:10])
                            if entry['para_string'] == 'n-a':
                                logger.warn("couldn't get para_string for %s para (value was set to n-a)" % category)

        else:
            logger.warn("report_dict is empty")
        return report_dict
    except Exception, e:
        logger.error('Failed calculating para_indexes for para_ids, exiting', exc_info=True)
        sys.exit(1)
