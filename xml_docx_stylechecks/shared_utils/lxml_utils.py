######### IMPORT SOME STANDARD PY LIBRARIES

import sys
import os
import shutil
import uuid
import re
import json
import logging
from lxml import etree
# import time

# ######### IMPORT LOCAL MODULES
import cfg

######### LOCAL DECLARATIONS
styles_xml = cfg.styles_xml

# Local namespace vars
wnamespace = cfg.wnamespace
w14namespace = cfg.w14namespace
wordnamespaces = cfg.wordnamespaces

# initialize logger
logger = logging.getLogger(__name__)


######### METHODS
# generate ar random id
def generate_id():
    idbase = uuid.uuid4().hex
    idshort = idbase[:8]
    idupper = idshort.upper()
    return str(idupper)

# take a random id and make sure it is unique in the document, otherwise generate a new one, forever
def generate_para_id(doc_root):
    iduniq = generate_id()
    idsearchstring = './/*w:p[@w14:paraId="%s"]' % iduniq
    while len(doc_root.findall(idsearchstring, wordnamespaces)) > 0:
        print iduniq + " already exists, generating another id"
        iduniq = generate_id()
        idsearchstring = './/*w:p[@w14:paraId="%s"]' % iduniq
    logger.debug("generated unique para-id: '%s'" % iduniq)
    return str(iduniq)

# taken from: https://stackoverflow.com/questions/42875103/integer-to-roman-number
def int_to_Roman(num):
   val = (1000, 900,  500, 400, 100,  90, 50,  40, 10,  9,   5,  4,   1)
   syb = ('M',  'CM', 'D', 'CD','C', 'XC','L','XL','X','IX','V','IV','I')
   roman_num = ""
   for i in range(len(val)):
      count = int(num / val[i])
      roman_num += syb[i] * count
      num -= val[i] * count
   return roman_num

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
def getParaId(para, doc_root):
    # checking tag to make sure we've grabbed a paragraph element
    good_paratag = '{%s}p' % wnamespace
    if para is not None and para.tag == good_paratag:
        attrib_id_key = '{%s}paraId' % w14namespace
        para_id = para.get(attrib_id_key)
        if para_id is None:
            # create a new para_id and set it for the para!
            new_para_id = generate_para_id(doc_root)
            logger.warn("no para_id: making & setting our own: %s" % new_para_id)
            para.attrib["{%s}paraId" % w14namespace] = new_para_id
            para_id = new_para_id
    else:
        if para.tag != good_paratag:
            logger.debug("tried to set p-iD for a non para element: {}".format(para.tag))
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

def getRunStyle(run):
    try:
        rstyle = run.find(".//*w:rStyle", wordnamespaces)
        if rstyle is not None:
            stylename = rstyle.get('{%s}val' % wnamespace)
        else:
            stylename = ""    # Default paragraph style
    except:
        stylename = ""
    return stylename

# lookup longname of style in styles.xml of file.
#  save looked up values in a dict to speed up repeat lookups if desired
def getStyleLongname(styleshortname, stylenamemap={}):
    # print styleshortname#, stylenamemap
    styles_tree = etree.parse(styles_xml)
    styles_root = styles_tree.getroot()
    if styleshortname == "n-a":
        stylelongname = "not avaliable"
    elif styleshortname in stylenamemap:
        stylelongname = stylenamemap[styleshortname]
        # print "in the map!"
    else:
        # print "not in tht emap!"
        searchstring = ".//w:style[@w:styleId='%s']/w:name" % styleshortname
        stylematch = styles_root.find(searchstring, wordnamespaces)
        # get fullname value and test against Macmillan style list
        if stylematch is not None:
            stylelongname = stylematch.get('{%s}val' % wnamespace)
            stylenamemap[styleshortname] = stylelongname
        else:
            stylelongname = styleshortname
    return stylelongname

# the "Run" here would be a span / character style.
# This method is identical to the one in lxmlutils for paras except varnames and the xml keyname
def findRunsWithStyle(stylename, doc_root):
    runs = []
    searchstring = ".//*w:rStyle[@w:val='%s']" % stylename
    for rstyle in doc_root.findall(searchstring, wordnamespaces):
        run = rstyle.getparent().getparent()
        runs.append(run)
    return runs

# gets returns para of selected element
def getParaParentofElement(element):
    tmp_element = element
    while tmp_element.tag != '{%s}p' % wnamespace and tmp_element.tag != '{%s}body' % wnamespace and tmp_element.getparent() is not None:
        tmp_element = tmp_element.getparent()
    if tmp_element.tag == '{%s}body' % wnamespace:
        para = None
    else:
        para = tmp_element
    return para

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

# return a dict of neighboring run elements and their text
def getNeighborRuns(run):
    rneighbors = {}
    try:
        rneighbors['prev'] = run.getprevious()
        # the 'len' errors and kicks to the except statement when no prev sibling
        len(rneighbors['prev'].tag)
        # make sure we are capturing the previous _run_, and not other interloping element
        if rneighbors['prev'].tag != '{%s}r' % wnamespace:
            while rneighbors['prev'].tag != '{%s}r' % wnamespace:
                rneighbors['prev'] = rneighbors['prev'].getprevious()
                # again, the 'len' call results in empty values when there is no previous
                len(rneighbors['prev'].tag)
        len(rneighbors['prev'].tag)
        rneighbors['prevtext'] = getParaTxt(rneighbors['prev'])
        rneighbors['prevstyle'] = getRunStyle(rneighbors['prev'])
    except:
        rneighbors['prev'] = ""
        rneighbors['prevtext'] = ""
        rneighbors['prevstyle'] = ""
    try:
        rneighbors['next'] = run.getnext()
        # the 'len' errors and kicks to the except statement when no next sibling
        len(rneighbors['next'].tag)
        # make sure we are capturing the next _run_, and not other interloping element
        if rneighbors['next'].tag != '{%s}r' % wnamespace:
            while rneighbors['next'].tag != '{%s}r' % wnamespace:
                rneighbors['next'] = rneighbors['next'].getnext()
                # again, the 'len' call results in empty values when there is no next
                len(rneighbors['next'].tag)
        rneighbors['nexttext'] = getParaTxt(rneighbors['next'])
        rneighbors['nextstyle'] = getRunStyle(rneighbors['next'])
    except:
        rneighbors['next'] = ""
        rneighbors['nexttext'] = ""
        rneighbors['nextstyle'] = ""
    return rneighbors

# creating a dict for section_names: SectionStart stylenames as keys, their 'longnames' as values
#   from SectionStartRules file
def getAllSectionNamesFromSSR(section_start_rules):
    section_names = {}
    for section_name in section_start_rules:
        section_names[transformStylename(section_name)] = section_name
    return section_names

# creating a dict for section_names: SectionStart stylenames as keys, their 'longnames' as values
#   from VBA styleconfig file
def getAllSectionNamesFromVSC(vbastyleconfig_dict):
    section_names = {}
    for section_name in vbastyleconfig_dict["sectionstarts"]:
        section_names[transformStylename(section_name)] = section_name
    return section_names

# return the last SectionStart para's content
def getContainerName(para, section_names, container_start_styles, container_end_styles):
    allcontainerandsections = section_names + container_start_styles + container_end_styles
    containername = ""
    if para is not None:
        tmp_para = para
        stylename = getParaStyle(tmp_para)
        while stylename and stylename not in allcontainerandsections:
            pneighbors = getNeighborParas(tmp_para)
            tmp_para = pneighbors['prev']
            stylename = pneighbors['prevstyle']
        if stylename in container_start_styles:
            containername = getParaStyle(tmp_para)
    return containername

# return the last SectionStart para's content
def getSectionName(para, section_names):
    if para is not None:
        tmp_para = para
        stylename = getParaStyle(tmp_para)
        while stylename and stylename not in section_names:
            pneighbors = getNeighborParas(tmp_para)
            tmp_para = pneighbors['prev']
            stylename = pneighbors['prevstyle']
        if stylename in section_names:
            sectionpara_name = stylename
            sectionpara_contents = getParaTxt(tmp_para)
        else:
            sectionpara_name = 'n-a' # this could happen if we are trying to get id of a prev or next para that does not exist
            sectionpara_contents = ''
    else:
        sectionpara_name = 'n-a' # this could happen if we are trying to get id of a prev or next para that does not exist
        sectionpara_contents = ''
    return sectionpara_name, sectionpara_contents

def getContentsForSectionStart(sectionbegin_para, doc_root, headingstyles, section_name, section_names):
    # get para style. If not in heading styles or sectiontypes["all"], get next (while)
    tmp_para = sectionbegin_para
    stylename = getParaStyle(tmp_para)
    while stylename and stylename not in headingstyles and stylename not in section_names:
        pneighbors = getNeighborParas(tmp_para)
        tmp_para = pneighbors['next']
        stylename = pneighbors['nextstyle']
    # set content equal to the content of the first heading-styled para in the section
    #   if heading para is chapnumber or partnumber, join it with the following Chap Title / Part title para(s?)
    #   (manually setting Copyright section's name)
    if section_name == cfg.copyrightsection_stylename:
        newcontent = 'Copyright'
    elif stylename in headingstyles:
        if stylename == cfg.chapnumstyle or stylename == cfg.partnumstyle:
            pneighbors = getNeighborParas(tmp_para)
            if (stylename == cfg.chapnumstyle and pneighbors['nextstyle'] == cfg.chaptitlestyle) or (stylename == cfg.partnumstyle and pneighbors['nextstyle'] == cfg.parttitlestyle):
                newcontent = "{}: {}".format(getParaTxt(tmp_para).strip(), pneighbors['nexttext'].strip())
        else:
            newcontent = getParaTxt(tmp_para).strip()
    else:   # if there was no heading-styled para
        sectionLongName = section_names[section_name]
        sectionshortname = sectionLongName[8:].split()[0]
        newcontent = sectionshortname
    # newcontent = "HALLOOOOOO"  (debug)
    return newcontent

def addRunToPara(content, para, bool_rm_existing_contents=False):
    if bool_rm_existing_contents == True:
        # delete any existing run(s) in para
        runs = para.findall(".//w:r", wordnamespaces)
        for run in runs:
            run.getparent().remove(run)
        # create new run element with new content & append to our para!
        new_para_run = etree.Element("{%s}r" % wnamespace)
        new_para_run_text = etree.Element("{%s}t" % wnamespace)
        new_para_run_text.text = content
        new_para_run.append(new_para_run_text)
        para.append(new_para_run)

def sectionStartTally(report_dict, section_names, doc_root, call_type, headingstyles = []):
    logger.info("logging all paras with SectionStart styles, and fixing any 'empty' sectionStart paras (no content)")
    # reset from any previous tallies:
    report_dict["section_start_found"] = []
    report_dict["empty_section_start_para"] = []
    # report_dict["empty_section_start_para"] = []
    # logger.warn("start = %s" % time.strftime("%y%m%d-%H%M%S"))
    if call_type == "insert":
        logger.info("writing contents to any empty sectionstart paras")
    for pstyle in doc_root.findall(".//*w:pStyle", wordnamespaces):
        stylename = pstyle.get('{%s}val' % wnamespace)
        # logger.info(stylename) # debug
        if stylename in section_names:
            para = pstyle.getparent().getparent()
            section_name = stylename
            # log the section start para
            #   (we can run this before content is added to paras, b/c that content is captured later in the 'calcLocationInfoForLog' method
            report_dict = logForReport(report_dict,doc_root, para, "section_start_found", section_name)

            # check to see ifthe para is empty (no contents) and if so log it, and, if 'call_type' is insert, fix it.
            if not getParaTxt(para).strip():
                if call_type == "insert":
                    # find / create contents for Section start para
                    pneighbors = getNeighborParas(para)
                    content = getContentsForSectionStart(pneighbors['next'], doc_root, headingstyles, section_name, section_names)
                    # add new content to Para! ()'True' = remove existing run(s) from para that may contain whitespace)
                    addRunToPara(content, para, True)
                    # log it for report
                    report_dict = logForReport(report_dict,doc_root,para,"wrote_to_empty_section_start_para",section_name)
                else:
                    # log it for report
                    report_dict = logForReport(report_dict,doc_root,para,"empty_section_start_para",section_name)
    # logger.warn("finish = %s" % time.strftime("%y%m%d-%H%M%S"))
    return report_dict

def createMiscElement(element_name, namespace, attribute_name='', attr_val='', attr_namespace=''):
    # have to triple-curly brace curly braces in format string for a single set to be escaped!
    misc_el = etree.Element("{{{}}}{}".format(namespace, element_name))
    # body = etree.Element("{%s}body" % cfg.wnamespace)
    if attribute_name:
        misc_el.attrib["{{{}}}{}".format(attr_namespace, attribute_name)] = attr_val
    return misc_el

def createRun(runtxt, rstylename=''):
    # create run
    run = etree.Element("{%s}r" % cfg.wnamespace)
    if runtxt:
        run_text = etree.Element("{%s}t" % cfg.wnamespace)
        run_text.text = runtxt
        run.append(run_text)
    if rstylename:
        # create new run properties element
        run_props = etree.Element("{%s}rPr" % cfg.wnamespace)
        run_props_style = etree.Element("{%s}rStyle" % cfg.wnamespace)
        run_props_style.attrib["{%s}val" % cfg.wnamespace] = rstylename
        # append props element to run element
        run_props.append(run_props_style)
        run.append(run_props)
    return run

def createPara(xml_root, pstylename='', runtxt='', rstylename='', para_id=''):
    # create para
    new_para = etree.Element("{%s}p" % cfg.wnamespace)
    if para_id:
        new_para_id = para_id
    else:
        new_para_id = generate_para_id(xml_root)
    new_para.attrib["{%s}paraId" % cfg.w14namespace] = new_para_id
    # if parastyle specified, add it here
    if pstylename:
        # create new para properties element
        new_para_props = etree.Element("{%s}pPr" % cfg.wnamespace)
        # and pstyle el
        new_para_props_style = etree.Element("{%s}pStyle" % cfg.wnamespace)
        new_para_props_style.attrib["{%s}val" % cfg.wnamespace] = pstylename
        # append props element to para element
        new_para_props.append(new_para_props_style)
        new_para.append(new_para_props)
    if runtxt or rstylename:
        run = createRun(runtxt, rstylename='')
        new_para.append(run)
    return new_para

# Should revisit this using lxml builder
# def insertSectionStart(sectionstylename, sectionbegin_para, doc_root, contents=''):
def insertPara(sectionstylename, existing_para, doc_root, contents, insert_before_or_after):
    logger.debug("commencing insertPara '%s' with style: '%s' and contents: '%s'" % (insert_before_or_after, sectionstylename, contents))
    # create new para element
    new_para_id = generate_para_id(doc_root)
    new_para = etree.Element("{%s}p" % wnamespace)
    new_para.attrib["{%s}paraId" % w14namespace] = new_para_id

    # create new para properties element
    new_para_props = etree.Element("{%s}pPr" % wnamespace)
    new_para_props_style = etree.Element("{%s}pStyle" % wnamespace)
    new_para_props_style.attrib["{%s}val" % wnamespace] = sectionstylename

    # append props element to para element
    new_para_props.append(new_para_props_style)
    new_para.append(new_para_props)

    # # create run and text elements, add text, and append to para
    #   Tried using "addRunToPara" function here, but returning newpara did not return new nested items
    new_para_run = etree.Element("{%s}r" % wnamespace)
    new_para_run_text = etree.Element("{%s}t" % wnamespace)
    new_para_run_text.text = contents
    new_para_run.append(new_para_run_text)
    new_para.append(new_para_run)
    # logtext = "inserted paragraph with style '%s' and text '%s'" % (sectionstylename,contents)

    # append insert new paragraph before the selected para element
    if insert_before_or_after == "before":
        existing_para.addprevious(new_para)
    elif insert_before_or_after == "after":
        existing_para.addnext(new_para)

    logger.info("inserted '%s' paragraph." % sectionstylename)


def autoNumberSectionParaContent(report_dict, section_names, autonumber_sections, doc_root):
    logger.info("check if autonumbering is necessary for Section Start para contents")
    autonumber_section_counts = {}
    for sectionlongname in autonumber_sections:
        autonumber_section_counts[sectionlongname] = 0
    # count the occurences of generic naming for sections above
    for sectionlongname in autonumber_section_counts:
        logger.debug("SECTION: %s" % sectionlongname)
        paras = findParasWithStyle(transformStylename(sectionlongname), doc_root)
        for para in paras:
            logger.debug("PARATXT: %s, SHORTNAME: %s" % (getParaTxt(para).strip(), sectionlongname[8:].split()[0]))
            if getParaTxt(para).strip() == sectionlongname[8:].split()[0]:     # sectionlabel[8:].split()[0] gives the 'shortname' of a given section from its full Seciton style name
                autonumber_section_counts[sectionlongname] += 1
            elif getParaTxt(para).strip():      # if we find a seciton with one of these types that already has non-generic contents, we bypass auto-numbering
                autonumber_section_counts[sectionlongname] = 0
                break
    logger.debug("autonumber_section_counts: %s" % autonumber_section_counts) # debug

    # apply autonumbering for each section in "autonumber_section_counts" as applicable
    for sectionlongname, count in autonumber_section_counts.iteritems():
        if count > 1:
            logger.info("Found %s '%s's with generic names in ssparas, adding autonumbering to sspara contents" % (count, sectionlongname))
            autonum = 1
            paras = findParasWithStyle(transformStylename(sectionlongname), doc_root)
            for para in paras:
                number = autonum
                if autonumber_sections[sectionlongname] == "alpha":
                    number = chr(number+64)
                if autonumber_sections[sectionlongname] == "roman":
                    number = int_to_Roman(number)
                newcontent = "{} {}".format(sectionlongname[8:].split()[0], number)
                # add new content to Para! ()'True' = remove existing run(s) from para that may contain whitespace)
                addRunToPara(newcontent, para, True)
                # increment autonum
                autonum += 1
                # optional logging:
                logForReport(report_dict,doc_root,para,"autonumbering_applied",sectionlongname)

    return report_dict

# a method to log paragraph id for style report etc
def logForReport(report_dict,doc_root,para,category,description):
    para_dict = {}
    para_dict["para_id"] = getParaId(para, doc_root)
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

def getCalculatedParaInfo(report_dict_entry, root, section_names, para, category, rootname=''):
    entry = report_dict_entry
    # # # Assign Section-start info for notes/comments.xml & get note_id
    if rootname:
        entry['parent_section_start_type'], entry['parent_section_start_content'] = rootname, 'n-a'
        entry['para_index'] = 'n-a'
        if para.getparent() is not None:
            entry['note-or-comment_id'] = para.getparent().get('{%s}id' % wnamespace)
    else:
        # # # Get para index
        entry['para_index'] = getParaIndex(para)
        if entry['para_index'] == 'n-a':
            logger.warn("couldn't get para-index for %s para (value was set to n-a)" % category)

        # check if we have a tablecell_paras, get section info accordingly
        tablecell_tag = '{%s}tc' % wnamespace
        if para is not None and para.getparent().tag == tablecell_tag:
            entry['para_index'] = 'tablecell_para'
            # get section name, start text based on para.parent.parent.parent: (table)
            table = para.getparent().getparent().getparent()
            entry['parent_section_start_type'], entry['parent_section_start_content'] = getSectionName(table, section_names)
        else:
            # get section name, section start text etc
            entry['parent_section_start_type'], entry['parent_section_start_content'] = getSectionName(para, section_names)

        if entry['parent_section_start_type'] == 'n-a' or entry['parent_section_start_content'] == 'n-a':
            logger.warn("couldn't get section start info for %s para (value was set to n-a)" % category)
    # # # Get 1st 10 words of para text
    entry['para_string'] = ' '.join(getParaTxt(para).split(' ')[:10])
    if entry['para_string'] == 'n-a':
        logger.warn("couldn't get para_string for %s para (value was set to n-a)" % category)

    return entry

# once all changes have been made, call this to add location info for users to the changelog dicts
def calcLocationInfoForLog(report_dict, root, section_names, alt_roots=[]):
    logger.info("calculating para_index numbers for all para_ids in 'report_dict'")
    try:
        # make sure we have contents in the dict
        if report_dict:
            for category, entries in report_dict.iteritems():
                # exclude hard-coded 'marker' attributes for validator_main; also skipping deleted objects from commentxml files
                if category != 'validator_py_complete' and category != 'percent_styled' and 'deleted_objects-comments-comments' not in category:
                    for entry in entries:
                        for key in entry.keys():
                            if key == "para_id":
                                # Get the para object
                                searchstring = ".//*w:p[@w14:paraId='%s']" % entry[key]
                                para = root.find(searchstring, wordnamespaces)
                                # If we can't find para object in main doc, check endnotes.xml and footnotes.xml
                                entry_calc_done = False
                                if para is None:
                                    logger.debug("found a para not in main doc_xml")
                                    if alt_roots:
                                        for rootname, alt_root in alt_roots.iteritems():
                                            logger.debug("checking {}, searchstring {}".format(rootname, searchstring))
                                            para = alt_root.find(searchstring, wordnamespaces)
                                            if para is not None:
                                                entry = getCalculatedParaInfo(entry, alt_root, section_names, para, category, rootname)
                                                entry_calc_done = True
                                                break # < stop checking altroots since we found one
                                if entry_calc_done == False:
                                    entry = getCalculatedParaInfo(entry, root, section_names, para, category)
        else:
            logger.warn("report_dict is empty")
        return report_dict
    except Exception, e:
        logger.error('Failed calculating para_indexes for para_ids, exiting', exc_info=True)
        sys.exit(1)
