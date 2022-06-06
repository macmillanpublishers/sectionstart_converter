######### IMPORT PY LIBRARIES
import os
import shutil
import re
# import uuid
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

def findSectionBegin(sectionname, section_start_rules, doc_root, versatileblockparas, para, cbstring):
    # set header lists
    headers = [lxml_utils.transformStylename(s) for s in section_start_rules[sectionname][cbstring]["styles"]]
    if "optional_heading_styles" in section_start_rules[sectionname][cbstring]:
        optheaders = [lxml_utils.transformStylename(s) for s in section_start_rules[sectionname][cbstring]["optional_heading_styles"]]
        allheaders = headers + optheaders
    else:
        allheaders = headers
    allheaders_plus_versatileparas = allheaders + versatileblockparas

    # set vars for our loop & output
    pneighbors = lxml_utils.getNeighborParas(para)
    sectionbegin_para = para
    sectionbegin_tmp = para
    firstStyleOfBlock = True

    # // scan upwards through any optional headers, versatile block paras, or styles in Style list (for contiguous block criteria)
    while pneighbors['prevstyle'] in allheaders_plus_versatileparas:
        logger.debug("found leading header/versatile styled para:'%s'" % pneighbors['prevstyle'])
        # increment the loop upwards
        sectionbegin_tmp = pneighbors['prev']
        pneighbors = lxml_utils.getNeighborParas(sectionbegin_tmp)
        sectionbegin_tmp_style = lxml_utils.getParaStyle(sectionbegin_tmp)
        # adjust matching & leadingParas if we found optional header or para with style from
        #  style list directly preceding a versatile block para
        if sectionbegin_tmp_style in allheaders:
            sectionbegin_para = sectionbegin_tmp
            # this is to help us save time, now we can stop processing this particular style-match
            if sectionbegin_tmp_style in headers:
                firstStyleOfBlock = False

    return sectionbegin_para, firstStyleOfBlock

def getCBStrings(counter):
    logger.debug("incrementing contiguous-block counter")
    counterstring = "%02d" % (counter,) # forces leading 0 for single digit
    cbstring = "contiguous_block_criteria_{}".format(counterstring)
    nextcounterstring = "%02d" % (counter+1,)
    nextcbstring = "contiguous_block_criteria_{}".format(counterstring)
    return cbstring, nextcbstring

def getMatchingParas(sectionname, section_start_rules, doc_root, cbstring):
    logger.debug("getting matchingParas...")
    matchingParas = []
    for stylename in section_start_rules[sectionname][cbstring]["styles"]:
        stylename = lxml_utils.transformStylename(stylename)
        searchstring = ".//*w:pStyle[@w:val='%s']" % stylename
        for pstyle in doc_root.findall(searchstring, wordnamespaces):
            para = pstyle.getparent().getparent()
            matchingParas.append(para)
    logger.debug("found '%s' matchingParas" % len(matchingParas))
    return matchingParas

def evalFirstChild(sectionname, section_start_rules, cbstring, sectionbegin_para):
    logger.debug("evaluating first-child rule...")
    textmatch = False
    matchedtext = ""
    paracontents = lxml_utils.getParaTxt(sectionbegin_para)
    # lets be case agnostic:
    if paracontents:
        paracontents = paracontents.lower()
    textarray = section_start_rules[sectionname][cbstring]["first_child"]["text"]
    for text in textarray:
        logger.debug("matchstring: '%s', para-contents: '%s'" % (text.lower(),paracontents))
        if text.lower() in paracontents:
            textmatch = True
            matchedtext = text

    # return values based on matchFound & desired match (positive or negative)
    matchrule = section_start_rules[sectionname][cbstring]["first_child"]["match"]
    if matchrule == True and textmatch == True:
        logger.debug("found 1st child positive match: '%s'" % matchedtext)
        return True
    elif matchrule == False and textmatch == False:
        logger.debug("found 1st child negative match: '%s'" % textarray)
        return True
    else:
        logger.debug("1st child match criteria not met: '%s'" % textarray)
        return False

# evaluate previous sibling (see if there's already a section start or acceptable prevSibling style)
def precedingStyleCheck(sectionname, section_start_rules, cbstring, sectionbegin_para, sectiontypes):
    logger.debug("checking prev-sibling for existing acceptable style...")
    # get acceptable previous sibling style list:
    requiredStyles = [lxml_utils.transformStylename(s) for s in section_start_rules[sectionname][cbstring]["previous_sibling"]["required_styles"]]
    required_plus_section_styles = requiredStyles + sectiontypes["all"]
    # get preceding para style
    pneighbors = lxml_utils.getNeighborParas(sectionbegin_para)
    # check to see if previous para style is already acceptable
    if pneighbors["prevstyle"] in required_plus_section_styles:
        logger.debug("previous style already has section start style: '%s'" % pneighbors["prevstyle"])
        return True
    else:
        return False

def evalPrevUntil(sectionname, section_start_rules, cbstring, sectionbegin_para):
    logger.debug("evaluating previous until rule...")
    requiredstyles = [lxml_utils.transformStylename(style) for style in section_start_rules[sectionname][cbstring]["previous_sibling"]["required_styles"]]
    prevuntil_styles = [lxml_utils.transformStylename(style) for style in section_start_rules[sectionname][cbstring]["previous_until"]]
    required_plus_prevuntil_styles = requiredstyles + prevuntil_styles

    # get previous para style then scan upwards with while loop
    pneighbors = lxml_utils.getNeighborParas(sectionbegin_para)
    para_tmp = sectionbegin_para

    while pneighbors['prevstyle'] and pneighbors['prevstyle'] not in required_plus_prevuntil_styles:
        # increment para upwards
        para_tmp = pneighbors['prev']
        pneighbors = lxml_utils.getNeighborParas(para_tmp)

    # figure out whether we matched a prevuntil style or required style
    if pneighbors['prevstyle'] in requiredstyles:
        logger.debug("false: found required-style before prev_until-style:'%s'" % pneighbors['prevstyle'])
        return False
    elif pneighbors['prevstyle'] in prevuntil_styles:
        logger.debug("true: found required-style before prev_until-style:'%s'" % pneighbors['prevstyle'])
        return True
    elif not pneighbors['prevstyle']:
        logger.debug("false: reached the beginning of the document, which indicates erroneous styling")
        return False

def evalPosition(sectionname, section_start_rules, cbstring, sectionbegin_para, sectiontypes):
    logger.debug("evaluate 'position' rule...")
    # get previous para style then scan upwards with while loop
    pneighbors = lxml_utils.getNeighborParas(sectionbegin_para)
    while pneighbors['prevstyle'] and pneighbors['prevstyle'] not in sectiontypes["all"]:
        # increment para upwards
        para_tmp = pneighbors['prev']
        pneighbors = lxml_utils.getNeighborParas(para_tmp)
    last_sectionstart = pneighbors['prevstyle']
    # in case there were no preceding section starts:
    if last_sectionstart not in sectiontypes["all"]:
        last_sectionstart = sectiontypes["frontmatter"][0]

    # get next SectionStart style
    pneighbors = lxml_utils.getNeighborParas(sectionbegin_para)
    # para_tmp = sectionbegin_para
    while pneighbors['nextstyle'] and pneighbors['nextstyle'] not in sectiontypes["all"]:
        # increment para (down)
        para_tmp = pneighbors['next']
        pneighbors = lxml_utils.getNeighborParas(para_tmp)
    next_sectionstart = pneighbors['nextstyle']
    # in case there were no follwoing section starts:
    if next_sectionstart not in sectiontypes["all"]:
        next_sectionstart = sectiontypes["backmatter"][0]

    # the desired 'position':
    position = section_start_rules[sectionname]["position"]

    # evaluate desired position vs. position as determined by Seciton start position
    if position == "frontmatter" and last_sectionstart in sectiontypes["frontmatter"]:
        logger.debug("'frontmatter' criteria matched- prev_sectionstart: '%s'" % last_sectionstart)
        return True
    elif position == "main" and ((last_sectionstart in sectiontypes["main"]) or (next_sectionstart in sectiontypes["main"])):
        logger.debug("'main' criteria matched- betweem '%s' and '%s'" % (last_sectionstart, next_sectionstart))
        return True
    elif position == "backmatter" and next_sectionstart in sectiontypes["backmatter"]:
        logger.debug("'backmatter' criteria matched- next_sectionstart: '%s'" % next_sectionstart)
        return True
    else:
        logger.debug("'%s' criteria not matched- betweem '%s' and '%s'" % (position, last_sectionstart, next_sectionstart))
        return False

def getSectionTypes(section_start_rules):
    logger.debug("getting section type lists")
    sectiontypes = {'all':[], 'frontmatter':[], 'main':[], 'backmatter':[]}
    for sectionname, value in section_start_rules.iteritems():
        sectiontypes["all"].append(lxml_utils.transformStylename(sectionname))
        if section_start_rules[sectionname]["section_type"] == "frontmatter":
            sectiontypes["frontmatter"].append(lxml_utils.transformStylename(sectionname))
        elif section_start_rules[sectionname]["section_type"] == "main":
            sectiontypes["main"].append(lxml_utils.transformStylename(sectionname))
        elif section_start_rules[sectionname]["section_type"] == "backmatter":
            sectiontypes["backmatter"].append(lxml_utils.transformStylename(sectionname))
    return sectiontypes

def checkForParaStyle(sectionname, doc_root):
    logger.debug("checking if section start style is present (%s)..." % sectionname)
    searchstring = ".//*w:pStyle[@w:val='%s']" % sectionname
    if doc_root.find(searchstring, wordnamespaces) is None:
        logger.debug("para style does not exist")
        return False
    else:
        logger.debug("para style exists")
        return True

def evalSectionRequired(sectionname, section_start_rules, doc_root, titlestyle):
    logger.debug("evaluate section-required rule...")
    # set default return to None
    sectionbegin_para = None
    # lets see if this section start is already present:
    if checkForParaStyle(lxml_utils.transformStylename(sectionname), doc_root) == False:
        # get insert_before styles
        insertstyles = [lxml_utils.transformStylename(s) for s in section_start_rules[sectionname]["section_required"]["insert_before"]]
        # two find the first insert style, I can either find the first occurrence of each
        #   insertstyle and compare para indexes, or start at the top of the document (titlepage) and scan downwards
        # For the only section_required style in use at time of writing this, (section-chapter),
        #   the latter seems less resource intensive.
        # It's possible we would encounter a doc wihtout a titlepage, but then we have bigger problems
        searchstring = ".//*w:pStyle[@w:val='%s']" % lxml_utils.transformStylename(titlestyle)
        titlestyle = doc_root.find(searchstring, wordnamespaces)
        if titlestyle is not None:
            titlepara = titlestyle.getparent().getparent()
            # get next SectionStart style
            pneighbors = lxml_utils.getNeighborParas(titlepara)
            # para_tmp = titlepara
            while pneighbors['nextstyle'] and pneighbors['nextstyle'] not in insertstyles:
                # increment para (down)
                para_tmp = pneighbors['next']
                pneighbors = lxml_utils.getNeighborParas(para_tmp)
            next_sectionstart = pneighbors['nextstyle']
            # this needs a conditional in case there were no following insertstyles ever:
            if next_sectionstart in insertstyles:
                sectionbegin_para = pneighbors['next']
                logger.debug("section_required criteria met; 1st insertbefore_style: '%s'" % next_sectionstart)
            else:
                logger.debug("no 'insert_before' styles found, cannot insert sectionstart")
        else:
            logger.debug("no titlepageTitle para, cannot process sectionrequired")

    return sectionbegin_para

# Should revisit this using lxml builder
# def insertSectionStart(sectionstylename, sectionbegin_para, doc_root, contents=''):
# def insertSectionStart(sectionstylename, sectionbegin_para, doc_root, contents, sectionnames):
#     logger.debug("commencing insert Section Start style: '%s'..." % sectionstylename)
#     # create new para element
#     new_para_id = lxml_utils.generate_para_id(doc_root)
#     new_para = etree.Element("{%s}p" % wnamespace)
#     new_para.attrib["{%s}paraId" % w14namespace] = new_para_id
#
#     # create new para properties element
#     new_para_props = etree.Element("{%s}pPr" % wnamespace)
#     new_para_props_style = etree.Element("{%s}pStyle" % wnamespace)
#     new_para_props_style.attrib["{%s}val" % wnamespace] = sectionstylename
#
#     # append props element to para element
#     new_para_props.append(new_para_props_style)
#     new_para.append(new_para_props)
#     # contents = lxml_utils.getContentsForSectionStart(sectionbegin_para, doc_root, headingstyles, sectionstylename, sectionnames)
#
#     # # create run and text elements, add text, and append to para
#     #   Tried using "addRunToPara" function here, but returning newpara did not return new nested items
#     new_para_run = etree.Element("{%s}r" % wnamespace)
#     new_para_run_text = etree.Element("{%s}t" % wnamespace)
#     new_para_run_text.text = contents
#     new_para_run.append(new_para_run_text)
#     new_para.append(new_para_run)
#     # logtext = "inserted paragraph with style '%s' and text '%s'" % (sectionstylename,contents)
#
#     # append insert new paragraph before the selected para element
#     sectionbegin_para.addprevious(new_para)
#     logger.info("inserted '%s' paragraph with contents: '%s'" % (sectionstylename,contents))

# check paragraph directly preceding the selected (passed) para to get rid of pagebreaks/pb-paras where approrpiate
# logging these is not absolutely necessary, building it in for troubleshooting
def deletePrecedingPageBreak(para, report_dict):
    logger.debug("checking for page break in preceding para...")
    pneighbors = lxml_utils.getNeighborParas(para)
    if len(pneighbors['prev']):
        # find all pagebreaks in the preceding paragraph
        pagebreakstring = ".//*w:br[@w:type='page']"
        breaks=pneighbors['prev'].findall(pagebreakstring, wordnamespaces)
        if len(breaks) == 1 and not pneighbors['prevtext'].strip():  # we need the strip.. apparently a pb carries some whitespace value
            logger.info("empty preceding pb para, deleting it")
            # # optional - log location for debug:  (has to come before removal or the reference fails (para is gone)
            # report_dict = lxml_utils.logForReport_old(report_dict,doc_root,para,"removed_pagebreak","rm'd pagebreak para preceding inserted section-start")
            # remove pagebreak para
            pneighbors['prev'].getparent().remove(pneighbors['prev'])
        elif len(breaks) > 0 and pneighbors['prevtext'].strip():
            # could remove the last pb anyways, here, consolidate with next case; or just remove the text and the pb
            logger.info("preceding pagebreak has text contents, not deleting")
        elif len(breaks) > 1 and not pneighbors['prevtext'].strip():
            logger.info("multiple pagebreak chars in preceding para: removing the last one")
            # # optional - log location for debug:  (has to come before removal or the reference fails (para is gone)
            # report_dict = lxml_utils.logForReport_old(report_dict,doc_root,para,"removed_pagebreak","rm'd preceding pagebreak char preceding inserted section-start")
            # remove last pagebreak char from the preceding paragraph
            breaks[len(breaks)-1].getparent().remove(breaks[len(breaks)-1])
        elif len(breaks) == 0:
            logger.info("preceding para is not a pagebreak, skipping delete")
    return report_dict

# adding a parameter so this can be run to add section styles or just report on where they should be added:
#   param name is 'call_type', expects strin :"insert" or "report"
def runRule(sectionname, section_start_rules, doc_root, versatileblockparas, sectiontypes, call_type, report_dict, titlestyle, headingstyles, sectionnames):
    # cycle through for multiple contiguous blocks, apply 'last' key=True, run the rule!
    counter = 1
    cbstring, nextcbstring = getCBStrings(counter)
    while cbstring in section_start_rules[sectionname]:
        logger.debug("* Running sectionstart rule for: '%s', contiguous_block_criteria_%s" % (sectionname, counter))
        sectionstylename = lxml_utils.transformStylename(sectionname)
        # see if the section already exists ; if so and multiple is False we can move on to the next section
        sectionpresent = checkForParaStyle(lxml_utils.transformStylename(sectionname), doc_root)
        if sectionpresent == True and section_start_rules[sectionname][cbstring]["multiple"] == False:
            # increment the cbstrings
            counter += 1
            cbstring, nextcbstring = getCBStrings(counter)
            continue

        # get all paras matching "styles" for this section
        matchingParas = getMatchingParas(sectionname, section_start_rules, doc_root, cbstring)

        # Walk through the criteria for each matching para!
        for para in matchingParas:
            # optional headers, blocks of matching styles are evaluated here, returning a "sectionbegin_para"
            sectionbegin_para, firstStyleOfBlock = findSectionBegin(sectionname, section_start_rules, doc_root, versatileblockparas, para, cbstring)
            # this para-match is redundant, another matching style para starts this block, so move on the next para
            if firstStyleOfBlock == False:
                logger.debug("disqualified: not the first para in its own style block! next match")
                continue

            # evaluate 1st child if criteria is present
            if "first_child" in section_start_rules[sectionname][cbstring]:
                firstchild_results = evalFirstChild(sectionname, section_start_rules, cbstring, sectionbegin_para)
                if firstchild_results == False:
                    logger.debug("disqualified: firstchild criteria not met. next match")
                    continue

            # evaluate previous sibling (see if there's already a section start)
            if precedingStyleCheck(sectionname, section_start_rules, cbstring, sectionbegin_para, sectiontypes) == True:
                logger.debug("disqualified: previous_style is acceptable. next match")
                continue

            # // check criteria for position
            if "position" in section_start_rules[sectionname]:
                # print "position is here"
                position_results = evalPosition(sectionname, section_start_rules, cbstring, sectionbegin_para, sectiontypes)
                if position_results == False:
                    logger.debug("disqualified: position criteria not met. next match")
                    continue

            # // check criteria for previous until
            if "previous_until" in section_start_rules[sectionname][cbstring]:
                prev_until_results = evalPrevUntil(sectionname, section_start_rules, cbstring, sectionbegin_para)
                if prev_until_results == False:
                    print ("disqualified: prev_until criteria not met. next match!")
                    continue

            # if we made it this far, go ahead and insert our section start &/or log it for style report!
            logger.info("All criteria met for '%s' rule!  %sing para" % (sectionname, call_type))
            if call_type == "insert":
                report_dict = deletePrecedingPageBreak(sectionbegin_para, report_dict)
                contents = lxml_utils.getContentsForSectionStart(sectionbegin_para, doc_root, headingstyles, sectionstylename, sectionnames)
                lxml_utils.insertPara(sectionstylename, sectionbegin_para, doc_root, contents, "before")
            report_dict = lxml_utils.logForReport_old(report_dict,doc_root,sectionbegin_para,"section_start_needed",sectionname)

            # break the loop for this rule if 'multiple' value is False
            if section_start_rules[sectionname][cbstring]["multiple"] == False:
                logger.debug("'Multiple' set to 'False', moving on to next rule.")
                break

        # increment the cbstrings
        counter += 1
        cbstring, nextcbstring = getCBStrings(counter)

    # evaluate section_required rule if present (could move this into 'main', does not need to be in this function)
    if "section_required" in section_start_rules[sectionname] and section_start_rules[sectionname]["section_required"]["value"] == True:
        logger.info("section_required is true, evaluating for %s" % sectionname)
        sectionbegin_para = evalSectionRequired(sectionname, section_start_rules, doc_root, titlestyle)
        #  if we have an insertion point for Section_Required, insert Section Start styled para
        if sectionbegin_para is not None:
            report_dict = lxml_utils.logForReport_old(report_dict,doc_root, sectionbegin_para,"section_start_needed","{}".format(sectionname))
            if call_type == "insert":
                report_dict = deletePrecedingPageBreak(sectionbegin_para, report_dict)
                contents = lxml_utils.getContentsForSectionStart(sectionbegin_para, doc_root, headingstyles, sectionstylename, sectionnames)
                lxml_utils.insertPara(sectionstylename, sectionbegin_para, doc_root, contents, "before")

    return report_dict

def setRulesPriority(section_start_rules):
    logger.debug("adding rule priorities")
    for sectionname, sectionvalues in section_start_rules.iteritems():
        if "order" in sectionvalues and sectionvalues["order"] == "first":
            section_start_rules[sectionname]["priority"] = 1
        elif "section_required" in sectionvalues and "value" in sectionvalues["section_required"]:
            section_start_rules[sectionname]["priority"] = 2
        elif "position" in sectionvalues:
            section_start_rules[sectionname]["priority"] = 4
        elif "order" in sectionvalues and sectionvalues["order"] == "last":
            section_start_rules[sectionname]["priority"] = 5
        else:
            section_start_rules[sectionname]["priority"] = 3
    return section_start_rules

def sectionStartCheck(call_type, report_dict, autonumber=False):
    logger.info("* * * commencing sectionStartCheck function...")
    # local vars
    section_start_rules_json = cfg.section_start_rules_json
    styleconfig_json = cfg.styleconfig_json
    doc_xml = cfg.doc_xml
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    titlestyle = cfg.titlestyle

    logger.info("reading in json resource files")
    # read rules & versatile block para list from JSONs
    section_start_rules = os_utils.readJSON(section_start_rules_json)
    styleconfig_dict = os_utils.readJSON(styleconfig_json)
    # "list comprehension!" https://stackoverflow.com/questions/7126916/perform-a-string-operation-for-every-element-in-a-python-list
    versatileblockparas = [classname[1:] for classname in styleconfig_dict["versatileblockparas"]]
    headingstyles = [classname[1:] for classname in styleconfig_dict["headingparas"]]

    # get section types (by position, + all sections)
    sectiontypes = getSectionTypes(section_start_rules)
    # get Section Start names & styles from sectionstartrules
    sectionnames = lxml_utils.getAllSectionNamesFromSSR(section_start_rules)

    # add priorities to rules
    section_start_rules = setRulesPriority(section_start_rules)

    logger.info("Cycle through Section Start rules in order of priority")
    # cycle through rules by priority and run them, priorities 1-10
    for n in range(1,11):
        for sectionname, sectionvalues in section_start_rules.iteritems():
            if section_start_rules[sectionname]["priority"] == n:
                report_dict = runRule(sectionname, section_start_rules, doc_root, versatileblockparas, sectiontypes, call_type, report_dict, titlestyle, headingstyles, sectionnames)

    # if 'converting', add autonumbering to sectionstart para contents where applicable, write our changes back to doc.xml
    if call_type == "insert":
        logger.info("writing changes out to doc_xml file")
        # autonumber contents for chapter, Appendix, Part
        if autonumber == True:
            report_dict = lxml_utils.autoNumberSectionParaContent(report_dict, sectionnames, cfg.autonumber_sections, doc_root)
        os_utils.writeXMLtoFile(doc_root, doc_xml)

    report_dict = lxml_utils.sectionStartTally(report_dict, sectionnames, doc_root, "report")

    # add/update para index numbers
    logger.debug("Update all report_dict records with para_index-")
    report_dict = lxml_utils.calcLocationInfoForLog(report_dict, doc_root, sectiontypes["all"])

    # create sorted version of "section_start_needed" list in reportdict based on para_index: for validator report
    if "section_start_needed" in report_dict:
        report_dict["section_start_needed__sort_by_index"] = sorted(report_dict["section_start_needed"], key=lambda x: x['para_index'])

    logger.info("* * * ending sectionStartCheck function.")

    return report_dict


#---------------------  MAIN
# only run if this script is being invoked directly
if __name__ == '__main__':

    # set up debug log to console
    logging.basicConfig(level=logging.DEBUG)

    report_dict = {}
    report_dict = sectionStartCheck("report", report_dict)
    report_dict = sectionStartCheck("insert", report_dict)

    logger.debug("report_dict contents:  %s" % report_dict)
