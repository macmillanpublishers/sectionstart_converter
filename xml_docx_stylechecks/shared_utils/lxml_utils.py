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
# # for benchmarking during tests:
# from shared_utils.decorators import benchmark as benchmark

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
        print (iduniq + " already exists, generating another id")
        iduniq = generate_id()
        idsearchstring = './/*w:p[@w14:paraId="%s"]' % iduniq
    logger.debug("generated unique para-id: '%s'" % iduniq)
    return str(iduniq)

def getElementCount(xml_root, element_string):
    searchstring = './/{}'.format(element_string)
    results = xml_root.findall(searchstring, wordnamespaces)
    return len(results)

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

def transformStylename(lngstylename):
    # separate rule for our (as of 12/21) 5 native word styles, where shortnames get camelcase
    if lngstylename in cfg.valid_native_word_styles:
        ls_split=lngstylename.split(' ')
        stylename = ''.join([x.capitalize() for x in ls_split])
    else:
        # matching restrictions for Wordml styleID's (observed): alphanumeric + hyphens ONLY
        #   (12/20 - switching from manual string of replacements to regex)
        stylename = re.sub('[^\w-]','',lngstylename)
        # stylename = stylename.replace(" ","").replace("(",'').replace(")",'').replace("#",'').replace("_",'')
    return stylename

# return all text from a paragraph (or run)
def getParaTxt(para):
    try:
        paratext = "".join([x for x in para.itertext()])
    except:
        paratext = "n-a"
    return paratext

# for wdv-397, occasionally a doc.xml's xml is prettified already, leading to a stylereport with extraneous newlines and whitespace
def getXmlRootfromFile(xmlfile, xmlfile_name):
    xmltree = etree.parse(xmlfile)
    xmlstr = etree.tostring(xmltree)
    nl = xmlstr.count(b'\n')
    if nl:
        logger.info("{} xml had {} newlines: converting to string and back to minify xml".format(xmlfile_name, nl))
        parser = etree.XMLParser(remove_blank_text=True)
        xmlroot = etree.XML(xmlstr, parser)
    else:
        xmlroot = xmltree.getroot()
    return xmlroot

# note: if new nsprefix already exists, with different uri, the old uri is preserved.
#   for our current purposes this is fine, we are only adding new ns if existing one is not present
def addNamespace(xmlroot, new_nsprefix, new_nsuri):
    nsmap_prefixes = [new_nsprefix]
    # get current nsmap prefixes
    for k, v in xmlroot.nsmap.items():
        nsmap_prefixes.append(k)
    # get unique values
    nsmap_prefixes = list(set(nsmap_prefixes))
    # retain current namespace, add new one:
    etree.cleanup_namespaces(xmlroot, top_nsmap={new_nsprefix: new_nsuri}, keep_ns_prefixes=nsmap_prefixes)

def checkNamespace(xmlroot, nsprefix):
    ns_present = False
    if nsprefix in xmlroot.nsmap: #and \
        # xmlroot.nsmap['{}'.format(nsprefix)] == cfg.wordnamespaces['{}'.format(nsprefix)]:
        ### ^ uncomment the above (and update related test) if we decide to restrict content updates
        ###     based on ns values differing from our defined set
        ns_present = True
    return ns_present

def verifyOrAddNamespace(xmlroot, ns_prefix, ns_uri):
    # if namespace is not already present,
    if checkNamespace(xmlroot, ns_prefix) == False:
        # add namespace to top level nsmap, then
        addNamespace(xmlroot, ns_prefix, ns_uri)
        # verify ns was added successfully
        if checkNamespace(xmlroot, ns_prefix) == True:
            # get filename through tag
            root_tag = xmlroot.tag
            if root_tag is not None:
                filename = re.sub('{.*}','',root_tag)
            else:
                filename = '<unknown>'
            logger.warning("Had to add required global namespace '{}' to {}.xml file".format(ns_prefix, filename))
        # if ns is not present after adding, exit ungracefully (user will get std processing err, wf-mail will get alert)
        else:
            logger.error("EXITING: Unsuccessful at adding required global namespace '{}' to xmlroot".format(ns_prefix))
            sys.exit(1)

# return the w14:paraId attribute's value for a paragraph
def getParaId(para, doc_root):
    # checking tag to make sure we've grabbed a paragraph element
    good_paratag = '{%s}p' % wnamespace
    if para is not None and para.tag == good_paratag:
        attrib_id_key = '{%s}paraId' % w14namespace
        para_id = para.get(attrib_id_key)
        if para_id is None:
            # check/add w14 namespace:
            verifyOrAddNamespace(doc_root, 'w14', cfg.w14namespace)
            # create a new para_id and set it for the para!
            new_para_id = generate_para_id(doc_root)
            logger.warning("no para_id: making & setting our own: %s" % new_para_id)
            para.attrib["{%s}paraId" % w14namespace] = new_para_id
            para_id = new_para_id
    else:
        if para is not None and para.tag != good_paratag:
            logger.debug("tried to set p-iD for a non para element: {}".format(para.tag))
        para_id = 'n-a' # this could happen if we are trying to get id of a prev or next para that does not exist
                        #   or for file without w14 namespace, like endnotes/footnotes
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
    if os.environ.get('TEST_FLAG'):
        return styleshortname
    else:
        # print (styleshortname#, stylenamemap)
        styles_tree = etree.parse(styles_xml)
        styles_root = styles_tree.getroot()
        if styleshortname == "n-a":
            stylelongname = "not available"
        elif styleshortname in stylenamemap:
            stylelongname = stylenamemap[styleshortname]
            # print ("in the map!")
        else:
            # print ("not in tht emap!")
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
def getSpecifiedParentofElement(current_element, target_parent):
    tmp_element = current_element
    while tmp_element.tag != '{%s}%s' % (wnamespace, target_parent) and tmp_element.tag != '{%s}body' % wnamespace and tmp_element.getparent() is not None:
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
    # logger.warning("start = %s" % time.strftime("%y%m%d-%H%M%S"))
    if call_type == "insert":
        logger.info("writing contents to any empty sectionstart paras")
    for pstyle in doc_root.findall(".//*w:pStyle", wordnamespaces):
        stylename = pstyle.get('{%s}val' % wnamespace)
        # logger.info(stylename) # debug
        if stylename in section_names:
            para = pstyle.getparent().getparent()
            section_name = stylename
            # # check if we're in a table; if so, log it as an err and 'continue' to next para
            if para.getparent().tag == '{{{}}}tc'.format(wnamespace):
                logForReport(report_dict, doc_root, para, 'illegal_style_in_table', section_name, ['section_info'], section_names)
                continue
            # log the section start para
            #   (we can run this before content is added to paras, b/c that content is captured later in the 'calcLocationInfoForLog' method
            logForReport(report_dict, doc_root, para, 'section_start_found', section_name, ['section_info'], section_names)

            # check to see ifthe para is empty (no contents) and if so log it, and, if 'call_type' is insert, fix it.
            if not getParaTxt(para).strip():
                if call_type == "insert":
                    # find / create contents for Section start para
                    pneighbors = getNeighborParas(para)
                    content = getContentsForSectionStart(pneighbors['next'], doc_root, headingstyles, section_name, section_names)
                    # add new content to Para! ()'True' = remove existing run(s) from para that may contain whitespace)
                    addRunToPara(content, para, True)
                    # log it for report
                    logForReport(report_dict, doc_root, para, 'wrote_to_empty_section_start_para', section_name, ['section_info'], section_names)
                else:
                    # log it for report
                    logForReport(report_dict, doc_root, para, 'empty_section_start_para', section_name, ['section_info'], section_names)
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
    #  check xml root namepaces before setting para_id
    verifyOrAddNamespace(xml_root, 'w14', cfg.w14namespace)
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
    #  check xml root namepaces before setting para_id
    verifyOrAddNamespace(xml_root, 'w14', cfg.w14namespace)
    new_para.attrib["{%s}paraId" % cfg.w14namespace] = new_para_id

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
    for sectionlongname, count in autonumber_section_counts.items():
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
                logForReport_old(report_dict,doc_root,para,"autonumbering_applied",sectionlongname)

    return report_dict

def logForReport(report_dict, xml_root, para, category, description='', log_extras=[], section_names={}):#, para_id=None):
    # create a new dict for this item
    para_dict = {}
    if xml_root is not None:
        para_dict['xml_file'] = etree.QName(xml_root).localname
    else:
        para_dict['xml_file'] = 'document'

    # get para_id as needed
    if cfg.loglevel == 'DEBUG' or 'para_id' in log_extras or os.environ.get('TEST_FLAG'):
        if para is not None:
            para_dict["para_id"] = getParaId(para, xml_root)
        else:
            para_dict["para_id"] = 'n-a'

    # set description
    if description:
        para_dict["description"] = description

    # handle extras
    if log_extras:
        lookup_element = para
        tablecell_tag = '{%s}tc' % wnamespace
        # if we are in a table, set our top-level lookup for index and ss info to the table
        if para.getparent().tag == tablecell_tag:
            para_dict['tablecell_para'] = True
            lookup_element = para.getparent().getparent().getparent()
        if 'section_info' in log_extras and para_dict['xml_file'] == 'document':
            para_dict['parent_section_start_type'], para_dict['parent_section_start_content'] = getSectionName(lookup_element, section_names)
        if 'para_index' in log_extras:
            para_dict['para_index'] = getParaIndex(lookup_element)
        if 'para_string' in log_extras:
            para_dict['para_string'] = ' '.join(getParaTxt(para).split(' ')[:10])

    # finally, add to category in report_dict
    if category not in report_dict:
        report_dict[category] = []
    report_dict[category].append(para_dict.copy())

    return report_dict

# a method to log paragraph id for style report etc
def logForReport_old(report_dict,doc_root,para,category,description, para_id=None):
    para_dict = {}
    if para_id is None:
        para_dict["para_id"] = getParaId(para, doc_root)
    else:
        para_dict["para_id"] = para_id
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
    tablecell_tag = '{%s}tc' % wnamespace
    # # # Assign Section-start info for notes/comments.xml & get note_id
    if rootname:
        entry['parent_section_start_type'], entry['parent_section_start_content'] = rootname, 'n-a'
        if para is not None and para.getparent().tag == tablecell_tag:
            entry['para_index'] = 'tablecell_para'
        else:
            entry['para_index'] = 'n-a'
        if para.getparent() is not None:
            entry['note-or-comment_id'] = para.getparent().get('{%s}id' % wnamespace)
    else:
        # # # Get para index
        entry['para_index'] = getParaIndex(para)
        if entry['para_index'] == 'n-a':
            logger.warning("couldn't get para-index for %s para (value was set to n-a)" % category)

        # check if we have a tablecell_paras, get section info accordingly
        if para is not None and para.getparent().tag == tablecell_tag:
            entry['para_index'] = 'tablecell_para'
            # get section name, start text based on para.parent.parent.parent: (table)
            table = para.getparent().getparent().getparent()
            entry['parent_section_start_type'], entry['parent_section_start_content'] = getSectionName(table, section_names)
        else:
            # get section name, section start text etc
            entry['parent_section_start_type'], entry['parent_section_start_content'] = getSectionName(para, section_names)

        if entry['parent_section_start_type'] == 'n-a' or entry['parent_section_start_content'] == 'n-a':
            logger.warning("couldn't get section start info for %s para (value was set to n-a)" % category)
    # # # Get 1st 10 words of para text
    entry['para_string'] = ' '.join(getParaTxt(para).split(' ')[:10])
    if entry['para_string'] == 'n-a':
        logger.warning("couldn't get para_string for %s para (value was set to n-a)" % category)

    return entry

# once all changes have been made, call this to add location info for users to the changelog dicts
def calcLocationInfoForLog(report_dict, root, section_names, alt_roots=[]):
    logger.info("calculating para_index numbers for all para_ids in 'report_dict'")
    try:
        # make sure we have contents in the dict
        if report_dict:
            for category, entries in report_dict.items():
                # exclude hard-coded 'marker' attributes for validator_main
                if category != 'validator_py_complete' and category != 'percent_styled':
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
                                        for rootname, alt_root in alt_roots.items():
                                            logger.debug("checking {}, searchstring {}".format(rootname, searchstring))
                                            para = alt_root.find(searchstring, wordnamespaces)
                                            if para is not None:
                                                entry = getCalculatedParaInfo(entry, alt_root, section_names, para, category, rootname)
                                                entry_calc_done = True
                                                break # < stop checking altroots since we found one
                                if entry_calc_done == False:
                                    entry = getCalculatedParaInfo(entry, root, section_names, para, category)
        else:
            logger.warning("report_dict is empty")
        return report_dict
    except Exception as e:
        logger.error('Failed calculating para_indexes for para_ids, exiting', exc_info=True)
        sys.exit(1)
