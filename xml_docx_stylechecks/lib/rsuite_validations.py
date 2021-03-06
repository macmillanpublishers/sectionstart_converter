######### IMPORT PY LIBRARIES
import os
import json
import sys
import logging
import re
# make sure to install lxml: sudo pip install lxml
from lxml import etree

######### IMPORT LOCAL MODULES
if __name__ == '__main__':
    # to go up a level to read cfg (and other files) when invoking from this script (for testing).
    cfgpath = os.path.join(sys.path[0], '..', 'cfg.py')
    osutilspath = os.path.join(sys.path[0], '..', 'shared_utils', 'os_utils.py')
    lxmlutilspath = os.path.join(sys.path[0], '..', 'shared_utils', 'lxml_utils.py')
    import imp
    import doc_prepare
    import stylereports
    cfg = imp.load_source('cfg', cfgpath)
    os_utils = imp.load_source('os_utils', osutilspath)
    lxml_utils = imp.load_source('lxml_utils', lxmlutilspath)
else:
    import cfg
    import lib.doc_prepare as doc_prepare
    import lib.stylereports as stylereports
    import shared_utils.os_utils as os_utils
    import shared_utils.lxml_utils as lxml_utils


######### LOCAL DECLARATIONS
# Local namespace vars
wnamespace = cfg.wnamespace
w14namespace = cfg.w14namespace
wordnamespaces = cfg.wordnamespaces
xmlnamespace = cfg.xmlnamespace

# initialize logger
logger = logging.getLogger(__name__)


# #---------------------  METHODS

# # lifted most of this from stylereports.logTextOfParasWithStyle, so we can get paras with (section) context
# # if we end up needing styles within a container we can do more of a while loop, like we did with containers.
def logTextOfParasWithStyleInSection(report_dict, xml_root, sectionnames, sectionname, stylename, report_category):
    logger.info("* * * commencing logTextOfParasWithStyleInSection: section '%s', style '%s'  ..." % (sectionname, stylename))
    paras = lxml_utils.findParasWithStyle(lxml_utils.transformStylename(stylename), xml_root)
    for para in paras:
        # getSectionName returs section name and contents, we only need first arg
        current_sectionname = lxml_utils.getSectionName(para, sectionnames)[0]
        if current_sectionname == sectionname:
            paratxt = lxml_utils.getParaTxt(para)
            report_dict = lxml_utils.logForReport(report_dict,xml_root,para,report_category,paratxt)
    return report_dict

def logTextOfRunsWithStyleInSection(report_dict, xml_root, sectionnames, sectionname, stylename, report_category):
    logger.info("* * * commencing logTextOfRunsWithStyleInSection: section '%s', style '%s'  ..." % (sectionname, stylename))
    runs = lxml_utils.findRunsWithStyle(lxml_utils.transformStylename(stylename), xml_root)
    for run in runs:
        para = run.getparent()
        # getSectionName returs section name and contents, we only need first arg
        current_sectionname = lxml_utils.getSectionName(para, sectionnames)[0]
        if current_sectionname == sectionname:
            runtxt = lxml_utils.getParaTxt(run)
            # check and see if we've already captured this para for other runs
            already_captured = False
            this_para_id = lxml_utils.getParaId(para, xml_root)
            if report_category in report_dict:
                for x in report_dict[report_category]:
                    for key,value in x.iteritems():
                        if x["para_id"] == this_para_id:
                            already_captured = True
            if already_captured == False:
                report_dict = lxml_utils.logForReport(report_dict,xml_root,para,report_category,runtxt)
    return report_dict

def checkSecondPara(report_dict, xml_root, firstpara, sectionnames):
    logger.info("* * * commencing checkSecondPara function...")
    pneighbors = lxml_utils.getNeighborParas(firstpara)
    secondpara_style = pneighbors['nextstyle']
    print "secondpara_style", secondpara_style
    if secondpara_style not in sectionnames:
        logger.warn("second para style is not a Section Start style, instead is: " + secondpara_style)
        report_dict = lxml_utils.logForReport(report_dict,xml_root,pneighbors["next"],"non_section_start_styled_secondpara",lxml_utils.getStyleLongname(secondpara_style))
    return report_dict

def getContainerStarts(styleconfig_dict):
    container_start_styles = []
    for category in styleconfig_dict["containerparas"]:
        container_start_styles = container_start_styles + styleconfig_dict["containerparas"][category]
    # strip leading period
    container_start_styles = [s[1:] for s in container_start_styles]
    return container_start_styles

def deleteBookmarks(report_dict, xml_root, bookmark_items):
    logger.info("* * * commencing deleteBookmarks function...")
    start_searchstring = ".//{}".format(bookmark_items["bookmarkstart_tag"])
    for bookmark_start in xml_root.findall(start_searchstring, wordnamespaces):
            w_name = bookmark_start.get('{%s}name' % wnamespace)
            w_id = bookmark_start.get('{%s}id' % wnamespace)
            # now we find the paired 'end' that went with our start
            end_id_searchstring = ".//{}[@w:id='{}']".format(bookmark_items["bookmarkend_tag"], w_id)
            bookmark_end = xml_root.find(end_id_searchstring, wordnamespaces)
            # remove element(s) (get para first for logging)
            para = lxml_utils.getParaParentofElement(bookmark_start)
            bookmark_start.getparent().remove(bookmark_start)
            if bookmark_end is not None:
                bookmark_end.getparent().remove(bookmark_end)
            # log to report_dict as needed, logger for debug
            #   Note: we are silently deleting bookmarks of 'auto_bookmark' types, that were not inserted by users
            logger.debug("deleted bookmark named '%s'" % w_name)
            if para is not None and w_name not in bookmark_items["autobookmark_names"]:
                lxml_utils.logForReport(report_dict,xml_root,para,"deleted_objects-bookmarks","deleted bookmark named " + w_name)
    # delete any orphaned bookmark_ends silently. (these should not exist, but why not clean-up in case?)
    end_searchstring = ".//{}".format(bookmark_items["bookmarkend_tag"])
    for bookmark_end in xml_root.findall(end_searchstring, wordnamespaces):
        w_name = bookmark_end.get('{%s}name' % wnamespace)
        bookmark_end.getparent().remove(bookmark_end)
        logger.debug("deleted orphaned bookmark_end named '%s'" % w_name)
    return report_dict

# table cell (w:tc) elements require at least one w:p
#   so we cannot bilthely delete
def handleBlankParasInTables(report_dict, xml_root, para, log_category="table_blank_para", log_description="blank para found in table cell", skip_logging=False):
    tablepara = False
    # have a provision to flag only _solo_ table paras, so we can aggressively rm any blank para possible.
    #   looking back at teh original reqrement, that is both an edge case, and possibly overly intrusive
    #   So leaving provision in place with the following var; but selecting to note all tableparas as protected /
    #       deserving of flagging instead of removal, for now.
    removing_excess_tbl_blankparas = False
    # check for parent tablecell element
    if para.getparent().tag == '{{{}}}tc'.format(wnamespace):
        logger.debug("encountered blank tablecell para")
        # From this conditional, we handle whether we are only preservind _solo_blankparas, or all table paras
        if removing_excess_tbl_blankparas == True:
            paras_in_cell = para.getparent().findall(".//w:p", wordnamespaces)
            if len(paras_in_cell) == 1:
                #  para.getnext().tag != '{{{}}}p'.format(wnamespace)) and \
                # (pneighbors['prev'] is None or pneighbors['prev'].tag != '{{{}}}p'.format(wnamespace)):
                if skip_logging == False:
                    logger.info("encountered solo-table-para, tagging for report")
                    lxml_utils.logForReport(report_dict,xml_root,para,log_category,log_description)
                tablepara = True
            else:
                # print len(pneighbors['prev'])
                logger.debug("blank tablecell para has a neighbor, bouncing back to std blank para handling")
        else:
            if skip_logging == False:
                lxml_utils.logForReport(report_dict,xml_root,para,log_category,log_description)
            tablepara = True
    return report_dict, tablepara

# def removeBlankParas(xml_root, report_dict):
def removeBlankParas(report_dict, xml_root, sectionnames, container_start_styles, container_end_styles, spacebreakstyles):
    logger.info("* * * commencing removeBlankParas function...")
    specialparas = sectionnames.keys() + container_start_styles + container_end_styles + spacebreakstyles
    allparas = xml_root.findall(".//w:p", wordnamespaces)
    for para in allparas:
        # get paras with no content
        if not lxml_utils.getParaTxt(para).strip(): # or para.text is None:
            # checking for solo tablecell paras first: extremely unlikely in the 1st two cases but possible in the 3rd:
            #   and rm'ing one breaks output doc
            report_dict, tablepara = handleBlankParasInTables(report_dict, xml_root, para)
            if tablepara == False:
                # log special warnings for the report for 'special' blank paras
                parastyle = lxml_utils.getParaStyle(para)
                if parastyle in specialparas:
                    # get section info for report, since we will be unable to retrieve after para is deleted
                    sectionname, sectiontext = lxml_utils.getSectionName(para, sectionnames)
                    sectionfullname = lxml_utils.getStyleLongname(sectionname)
                    section_info = "'%s: \"%s\"'" % (sectionfullname, sectiontext)
                    # separate warning text for sectionparas versus others:
                    if parastyle in sectionnames.keys():
                        lxml_utils.logForReport(report_dict,xml_root,para,"removed_section_blank_para", sectionfullname)
                    elif parastyle in container_start_styles + container_end_styles:
                        lxml_utils.logForReport(report_dict,xml_root,para,"removed_container_blank_para","%s_%s" % (lxml_utils.getStyleLongname(parastyle), section_info))
                    elif parastyle in spacebreakstyles:
                        lxml_utils.logForReport(report_dict,xml_root,para,"removed_spacebreak_blank_para","%s_%s" % (lxml_utils.getStyleLongname(parastyle), section_info))

                # all paras are counted again so we gt a total for our count on the report
                lxml_utils.logForReport(report_dict,xml_root,para,"removed_blank_para","removed %s-styled para" % parastyle)
                # and then the blank para is removed
                para.getparent().remove(para)

    return report_dict

# so we can insert paras into test xml blocks with predictable para_ids instead of random (for assertions)
def getTestParaId(counter):
    if os.environ.get('TEST_FLAG'):
        para_id = 'p_id-{}'.format(counter)
    else:
        para_id = ''
    counter+=1
    return para_id, counter

def createDummyNotePara(xml_root, testpara_counter, note_stylename, noteref_stylename):
    # get test p_id as needed
    p_id, testpara_counter = getTestParaId(testpara_counter)
    # create basic para with empty ref-styled run
    newpara = lxml_utils.createPara(xml_root, note_stylename, '', noteref_stylename, p_id)
    # add endnote reference object
    noteref_el = lxml_utils.createMiscElement('endnoteRef', cfg.wnamespace)
    newrun = newpara[1]
    newrun.append(noteref_el)
    # and new placeholder-text run
    text_run = lxml_utils.createRun("[no text]")
    newpara.append(text_run)
    logger.info("for troubleshooting built element: {}".format(etree.tostring(newpara)))
    # * note: looked at trying to match pStyle for Endnotes, particularly w:spacing
    #   but maybe not worth it, since we may do harm by trying to guess at the style of endnotes per document
    #   (whether based on styles.xml, or another endnote, or a standard, any could be wrong)
    return newpara, testpara_counter

def handleBlankParasInNotes(report_dict, xml_root, note_stylename, noteref_stylename, note_name, note_section):
    logger.info("* * * commencing handleBlankParasInNotes function, for {}".format(note_section))
    # vars
    testpara_counter = 1
    searchstring = ".//w:{}".format(note_name)
    type = '{%s}type' % wnamespace
    separators = ['separator', 'continuationSeparator', 'continuationNotice']
    # handle notes objects with no para children
    for note in xml_root.findall(searchstring, wordnamespaces):
        if len(note) == 0: # <== note object has no children (this may not occur naturally, but why not cover?)
            dummy_notepara, testpara_counter = createDummyNotePara(xml_root, testpara_counter, note_stylename, noteref_stylename)
            note.append(dummy_notepara)
            note_id = note.get('{%s}id' % wnamespace)
            lxml_utils.logForReport(report_dict,xml_root,dummy_notepara,"found_empty_note",note_name)
        # handle note objects with children but no non-whitespace text
        elif not lxml_utils.getParaTxt(note).strip():
            uniquenote = True
            # skip separator notes
            if type in note.attrib and note.attrib[type] in separators:
                continue
            # find paras in note (could be nested in tables, hence search instead of laterally scanning siblings)
            for para in note.findall(".//w:p", wordnamespaces):
                # check if we are in a table; a table cell (w:tc) element requires at least one w:p
                #   skipping table para logging here, they will get captured again below in when cycling through paras
                report_dict, tablepara = handleBlankParasInTables(report_dict, xml_root, para, '', '', True)
                if tablepara == False:
                    para.getparent().remove(para)
                    # make sure we don't re-add dummy text and log for report for multiblank paras in the same note!
                    if uniquenote == True:
                        dummy_notepara, testpara_counter = createDummyNotePara(xml_root, testpara_counter, note_stylename, noteref_stylename)
                        note.append(dummy_notepara)
                        lxml_utils.logForReport(report_dict,xml_root,dummy_notepara,"found_empty_note",note_name)
                        uniquenote = False
                    else:
                        # log above blank para removal for summary
                        lxml_utils.logForReport(report_dict,xml_root,para,"removed_blank_para","excess blank para in empty {}".format(note_name))
                elif tablepara == True:
                    uniquenote = False

    # handle blank paras within notes that have other paras with content
    #   the search above captured any notes with solo blank paras, this captures the remainder, which can be rm'd
    allparas = xml_root.findall(".//w:p", wordnamespaces)
    for para in allparas:
        if not lxml_utils.getParaTxt(para).strip(): # or para.text is None:
            # check if we are in a table; a table cell (w:tc) element requires at least one w:p
            report_dict, tablepara = handleBlankParasInTables(report_dict, xml_root, para, 'table_blank_para_notes', 'blank para in table cell in {}'.format(note_name))
            if tablepara == False:
                # skip separator notes
                note = para.getparent()
                if type in note.attrib and note.attrib[type] in separators:
                    continue
                # log discovery
                note_id = note.get('{%s}id' % wnamespace)
                lxml_utils.logForReport(report_dict,xml_root,para,"removed_blank_para","blank para in {} note with other text; note_id: {}".format(note_name, note_id))
                # remove para
                para.getparent().remove(para)

    return report_dict

def checkContainers(report_dict, xml_root, sectionnames, container_start_styles, container_end_styles):
    logger.info("* * * commencing checkContainers function...")
    search_until_styles = sectionnames.keys() + container_start_styles + container_end_styles
    # adding counters to help determine if we need to check for orphan ENDS
    matched_cstart_count = 0
    # loop through searching for different container start styles
    for container_stylename in container_start_styles:
        # searchstring = ".//*w:pStyle[@w:val='%s']" % container_stylename
        containerstart_paras = lxml_utils.findParasWithStyle(container_stylename, xml_root)
        # containerstart_paras = xml_root.findall(searchstring, wordnamespaces)
        # loop through specific matched paras of a given style
        for start_para in containerstart_paras:
            # check if we're in a table; if so, log it as an err and 'continue' to next para
            if start_para.getparent().tag == '{{{}}}tc'.format(wnamespace):
                report_dict = lxml_utils.logForReport(report_dict,xml_root,start_para,"illegal_style_in_table",lxml_utils.getStyleLongname(container_stylename))
                continue
            # start_para = start_para_pStyle.getparent().getparent()
            pneighbors = lxml_utils.getNeighborParas(start_para)
            para_tmp = start_para

            # scan styles of next paras until we match something to stop at
            while pneighbors['nextstyle'] and pneighbors['nextstyle'] not in search_until_styles:
                # increment para downwards
                para_tmp = pneighbors['next']
                pneighbors = lxml_utils.getNeighborParas(para_tmp)
            # figure out whether we matched an END style or something else
            if pneighbors['nextstyle'] and pneighbors['nextstyle'] in container_end_styles:
                logger.debug("found a container end style before section start, container start or document end")
                matched_cstart_count = matched_cstart_count + 1
            else:
                lxml_utils.logForReport(report_dict,xml_root,start_para,"container_error",lxml_utils.getStyleLongname(container_stylename))
                if not pneighbors['nextstyle']:
                    logger.warn("container error - reached end of document before container-END styled para :( logging")
                elif pneighbors['nextstyle'] and pneighbors['nextstyle'] in sectionnames.keys():
                    logger.warn("container error - reached section-start style '%s' before container END styled para :(" % pneighbors['nextstyle'])
                else:
                    logger.warn("container error - reached container-start style '%s' before container END styled para :(" % pneighbors['nextstyle'])
    # doublecheck that no container Ends live in tables, or are orphaned:
    c_end_count = 0
    for end_stylename in container_end_styles:
        container_end_paras = lxml_utils.findParasWithStyle(end_stylename, xml_root)
        # containerstart_paras = xml_root.findall(searchstring, wordnamespaces)
        # loop through specific matched paras of a given style
        for end_para in container_end_paras:
            # check if we're in a table; if so, log it as an err and 'continue' to next para
            if end_para.getparent().tag == '{{{}}}tc'.format(wnamespace):
                report_dict = lxml_utils.logForReport(report_dict,xml_root,end_para,"illegal_style_in_table",lxml_utils.getStyleLongname(end_stylename))
            else:
                c_end_count = c_end_count + 1
    # if the counts don't match, we have an orphan END para somewhere.
    if c_end_count != matched_cstart_count:
        logger.warn("found %s matched start containers, but %s Container_ends, indicating orphaned END" % (matched_cstart_count, c_end_count))
        for end_stylename in container_end_styles:
            # the below loop is identical to teh one above, except reversed (scanning upwards from end paras for valid starts)
            logger.info("looping through container ends and scanning upwards, searching for orphan(s)")
            container_end_paras = lxml_utils.findParasWithStyle(end_stylename, xml_root)
            for end_para in container_end_paras:
                # skip illegal ends in tables, already flagged those
                if end_para.getparent().tag != '{{{}}}tc'.format(wnamespace):
                    pneighbors = lxml_utils.getNeighborParas(end_para)
                    para_tmp = end_para
                    # scan styles of next paras until we match something to stop at
                    while pneighbors['prevstyle'] and pneighbors['prevstyle'] not in search_until_styles:
                        # increment para downwards
                        para_tmp = pneighbors['prev']
                        pneighbors = lxml_utils.getNeighborParas(para_tmp)
                    # figure out whether we matched a START style or something else
                    if pneighbors['prevstyle'] and pneighbors['prevstyle'] in container_start_styles:
                        logger.debug("found container start, this END is part of a set")
                    else:
                        lxml_utils.logForReport(report_dict,xml_root,end_para,"container_end_error",lxml_utils.getStyleLongname(end_stylename))
                        if not pneighbors['prevstyle']:
                            logger.warn("orphaned container END - reached start of document before container-start para")
                        elif pneighbors['prevstyle'] and pneighbors['prevstyle'] in sectionnames.keys():
                            logger.warn("orphaned container END - reached section-start style '%s' before container START styled para" % pneighbors['prevstyle'])
                        else:
                            logger.warn("orphaned container END - reached another container-END para before container START para")

    return report_dict

def getListStylenames(styleconfig_dict):
    logger.info("* * * commencing getListStylenames function...")
    # prepare dicts of lists for collecting list styles by type & level
    li_styles_by_type = {
        "bullet": [],
        "unnum": [],
        "num": [],
        "alpha": []
    }
    li_styles_by_level = {
        "level1": [],
        "level2": [],
        "level3": []
    }

    # using list comprehensions to import stylenames, for clearing leading period(s)
    unorderedlistparaslevel1 = [s[1:] for s in styleconfig_dict["unorderedlistparaslevel1"]]
    orderedlistparaslevel1 = [s[1:] for s in styleconfig_dict["orderedlistparaslevel1"]]
    li_styles_by_level["level1"] = unorderedlistparaslevel1 + orderedlistparaslevel1

    unorderedlistparaslevel2 = [s[1:] for s in styleconfig_dict["unorderedlistparaslevel2"]]
    orderedlistparaslevel2 = [s[1:] for s in styleconfig_dict["orderedlistparaslevel2"]]
    li_styles_by_level["level2"] = unorderedlistparaslevel2 + orderedlistparaslevel2

    unorderedlistparaslevel3 = [s[1:] for s in styleconfig_dict["unorderedlistparaslevel3"]]
    orderedlistparaslevel3 = [s[1:] for s in styleconfig_dict["orderedlistparaslevel3"]]
    li_styles_by_level["level3"] = unorderedlistparaslevel3 + orderedlistparaslevel3

    listparagraphs = [s[1:] for s in styleconfig_dict["listparaparas"]]
    all_list_styles = li_styles_by_level["level1"] + li_styles_by_level["level2"] + li_styles_by_level["level3"]

    for liststyle in all_list_styles:
        if "bullet" in liststyle.lower():
            li_styles_by_type["bullet"].append(liststyle)
        elif "alpha" in liststyle.lower():
            li_styles_by_type["alpha"].append(liststyle)
        elif "unnum" in liststyle.lower():
            li_styles_by_type["unnum"].append(liststyle)
        elif "num" in liststyle.lower() and "unnum" not in liststyle.lower():
            li_styles_by_type["num"].append(liststyle)

    return li_styles_by_level, li_styles_by_type, listparagraphs, all_list_styles

def verifyListNesting(report_dict, xml_root, li_styles_by_level, li_styles_by_type, listparagraphs, all_list_styles):
    logger.info("* * * commencing verifyListNesting function...")

    #  cycle through all list styles by type
    for type in li_styles_by_type:
        for style in li_styles_by_type[type]:
            # # we are checking preceding para styles for compliance for every type of list
            # search and return matches
            list_p_elements = lxml_utils.findParasWithStyle(style, xml_root)
            # cycle through matches, get prev styles
            for list_p in list_p_elements:
                pneighbors = lxml_utils.getNeighborParas(list_p)
                prevstyle = pneighbors['prevstyle']
                # # if blank para, log as an error:
                # paratxt = lxml_utils.getParaTxt(list_p)
                # if not paratxt.strip():
                #     lxml_utils.logForReport(report_dict,xml_root,list_p,"blank_list_para","style is %s" % style)

                # determine our style's level
                for level in li_styles_by_level:
                    if style in li_styles_by_level[level]:
                        # check preceding parastyles for all listpara instances: if not same style of matching list type/level (para or non-para), log err
                        if style in listparagraphs:
                            if prevstyle not in li_styles_by_level[level] or prevstyle not in li_styles_by_type[type]:
                                lxml_utils.logForReport(report_dict,xml_root,list_p,"list_nesting_err","'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)))
                        # all other checks are for list non-paragraphs
                        else:
                            # if level1 list non-paras preceding paragraph is the same level but different type, issue warning
                            if level == "level1" and prevstyle in li_styles_by_level[level] and prevstyle not in li_styles_by_type[type]:
                                lxml_utils.logForReport(report_dict,xml_root,list_p,"list_change_warning","'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)))
                            # if level2 or level3 list non-paras preceding paragraph is the same level but different type, issue error
                            elif prevstyle in li_styles_by_level[level] and prevstyle not in li_styles_by_type[type]:
                                lxml_utils.logForReport(report_dict,xml_root,list_p,"list_change_err","'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)))
                            # list level 3 non-para, if preceded by list para, must be preceded by list level 2 or 3 (any)
                            elif level == "level3" and prevstyle not in li_styles_by_level["level2"] + li_styles_by_level["level3"]:
                                lxml_utils.logForReport(report_dict,xml_root,list_p,"list_nesting_err","'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)))
                            elif level == "level2" and prevstyle not in li_styles_by_level["level1"] + li_styles_by_level["level2"]:
                                if prevstyle in li_styles_by_level["level3"]:
                                    # go back until you find a non-level 3. IF it is a level1 or non list, error
                                    #   if it is a level 2 of a different kind, error.
                                    para_tmp = list_p
                                    while pneighbors['prevstyle'] and pneighbors['prevstyle'] in li_styles_by_level["level3"]:
                                        # increment para upwards
                                        para_tmp = pneighbors['prev']
                                        pneighbors = lxml_utils.getNeighborParas(para_tmp)
                                    if pneighbors['prevstyle'] not in li_styles_by_level["level2"] or pneighbors['prevstyle'] not in li_styles_by_type[type]:
                                        lxml_utils.logForReport(report_dict,xml_root,list_p,"list_nesting_err","'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)))
                                else:
                                    # prevstyle is not a list style at all
                                    lxml_utils.logForReport(report_dict,xml_root,list_p,"list_nesting_err","'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)))
    return report_dict

def checkEndnoteFootnoteStyles(xml_root, report_dict, note_style, sectionname):
    logger.info("* * * commencing checkEndnoteFootnoteStyles function, for %s..." % sectionname)
    # first check styles of paras
    allparas = xml_root.findall(".//w:p", wordnamespaces)
    for para in allparas:
        # note: our getParastyle function assumes paras with no pStyle are 'Normal'
        parastyle = lxml_utils.getParaStyle(para)
        if parastyle != note_style:
            # skip Endnote/Footnote separator paras (present by default, no style)
            parent_type = para.getparent().get('{%s}type' % wnamespace)
            if parent_type == 'separator' or parent_type == 'continuationSeparator' or parent_type == 'continuationNotice':
                continue
            if not os.environ.get('TEST_FLAG'):
                parastyle = lxml_utils.getStyleLongname(parastyle)
            lxml_utils.logForReport(report_dict,xml_root,para,"improperly_styled_%s" % sectionname, parastyle)
    return report_dict

def rmEndnoteFootnoteLeadingWhitespace(xml_root, report_dict, sectionname):
    logger.info("* * * commencing rmEndnoteFootnoteLeadingWhitespace function, for %s..." % sectionname)
    # now check leading text of note.
    allnotes = xml_root.findall(".//w:%s" % sectionname, wordnamespaces)
    for note in allnotes:
        note_id = note.get('{%s}id' % wnamespace)
        runs = note.findall(".//w:p/w:r", wordnamespaces)
        # cycle through runs with text, check for leading whitespace and handle accordingly
        run_removed = False
        for wr in runs:
            runtext = lxml_utils.getParaTxt(wr)
            if not runtext:
                # was just skipping blank runs, but tab characters don't register as text so need their own handling
                wtabs = wr.findall("w:tab", wordnamespaces)
                if len(wtabs):
                    for wtab in wtabs:
                        wr.remove(wtab)
                    # optionally log tab removals:
                    lxml_utils.logForReport(report_dict,xml_root,para,"%s-leading_whitespace_rmd" % sectionname,"note ref: %s (tab removed)" % note_id)
                continue
            para = wr.getparent()
            charcount = len(runtext)
            charcount_left_strip = len(runtext.lstrip())
            charcount_diff = charcount - charcount_left_strip
            # charcount_left_strip == 0 (no text, delete run, continue)
            if charcount_left_strip == 0:
                wr.getparent().remove(wr)
                run_removed = True
            # charcount_diff == 0 (no leading whitespace, break)
            elif charcount_diff == 0:
                #  we may want to comment out some of this logging, if it is logging for EVERY footnote.
                if run_removed:  # log if we previously removed a whole run
                    lxml_utils.logForReport(report_dict,xml_root,para,"%s-leading_whitespace_rmd" % sectionname,"note ref: %s (blank run removed)" % note_id)
                break
            # here we have leading whitespace in the w:r/w:t itself that must be removed. re-setting text with lstrip()
            else:
                wt = wr.find(".//w:t", wordnamespaces)
                wt.text = wt.text.lstrip()
                lxml_utils.logForReport(report_dict,xml_root,para,"%s-leading_whitespace_rmd" % sectionname,"note ref: %s" % note_id)
                break

    return report_dict

def flagCustomNoteMarks(xml_root, report_dict, ref_style_dict):
    logger.info("* * * commencing flagCustomNoteMarks function...")
    for note_type, ref_style in ref_style_dict.iteritems():
        ref_el_name = ref_style[0].lower() + ref_style[1:]
        searchstring = './/*w:{}[@w:customMarkFollows="1"]'.format(ref_el_name)
        customref_els = xml_root.findall(searchstring, wordnamespaces)
        if customref_els:
            for customref_el in customref_els:
                # get text, id of custom mark
                ref_run = customref_el.getparent()
                ref_text_el = ref_run.find(".//w:t", wordnamespaces)
                if ref_text_el is not None:
                    reftext = ref_text_el.text
                else:
                    ref_sym_el = ref_run.find(".//w:sym", wordnamespaces)
                    if ref_sym_el is not None:
                        reftext = "(custom ref-mark is symbol not text)"
                    else:
                        reftext = "(custom ref-mark is not symbol or text)"
                attrib_id_key = '{%s}id' % wnamespace
                ref_id = customref_el.get(attrib_id_key)
                para = lxml_utils.getParaParentofElement(ref_run)
                #log occurence
                lxml_utils.logForReport(report_dict,xml_root,para,"custom_%s_mark" % note_type,"custom note marker: '%s', %s id: %s" %(reftext, note_type, ref_id))
    return report_dict

def cleanNoteMarkers(report_dict, xml_root, noteref_object, note_style, report_category):
    logger.info("* * * commencing cleanNoteMarkers function, for %s..." % report_category)
    noteref_objects = xml_root.findall(".//%s" % noteref_object, wordnamespaces)
    for noteref in noteref_objects:
        note_run = noteref.getparent()
        runstyle = lxml_utils.getRunStyle(note_run)
        if runstyle != note_style:
            rstyle_obj = note_run.find(".//*w:rStyle", wordnamespaces)
            # if there is an incorrect run-style object, re-style
            if rstyle_obj is not None:
                attrib_style_key = '{%s}val' % wnamespace
                rstyle_obj.set(attrib_style_key, note_style)
                # and report it to log
                attrib_id_key = '{%s}id' % wnamespace
                note_id = noteref.get(attrib_id_key)
                lxml_utils.logForReport(report_dict, xml_root, note_run.getparent(), "note_markers_wrong_style", \
                    "restyled %s ref: no. %s (was styled as %s)" % (report_category, note_id, runstyle))
    return report_dict

def rsuiteValidations(report_dict):
    vbastyleconfig_json = cfg.vbastyleconfig_json
    styleconfig_json = cfg.styleconfig_json
    styles_xml = cfg.styles_xml
    doc_xml = cfg.doc_xml
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    # this is for writing out to any file where the xml_root is edited
    xmlfile_dict = {
        doc_root:doc_xml
        }
    # alt_roots is for inclusion in log calculations, as needed
    alt_roots = {}
    if os.path.exists(cfg.endnotes_xml):
        endnotes_tree = etree.parse(cfg.endnotes_xml)
        endnotes_root = endnotes_tree.getroot()
        xmlfile_dict[endnotes_root]=cfg.endnotes_xml
        alt_roots['Endnotes']=endnotes_root
    if os.path.exists(cfg.footnotes_xml):
        footnotes_tree = etree.parse(cfg.footnotes_xml)
        footnotes_root = footnotes_tree.getroot()
        xmlfile_dict[footnotes_root]=cfg.footnotes_xml
        alt_roots['Footnotes']=footnotes_root

    # get Section Start names & styles from vbastyleconfig_json
    #    Could pull styles from macmillan.json  with "Section-" if I don't want to use vbastyleconfig_json
    macmillanstyledata = os_utils.readJSON(cfg.macmillanstyles_json)
    vbastyleconfig_dict = os_utils.readJSON(vbastyleconfig_json)
    sectionnames = lxml_utils.getAllSectionNamesFromVSC(vbastyleconfig_dict)
    # get Container styles (shortnames only) from styleconfig_dict
    #   Could pull them from macmillan.json as well, searching for ALL CAPS.. though would need to get container-end from style_config
    #   Also might be able to reverse engineer Container longnames
    styleconfig_dict = os_utils.readJSON(styleconfig_json)
    container_start_styles = getContainerStarts(styleconfig_dict)
    container_end_styles = [s[1:] for s in styleconfig_dict["containerendparas"]] # <--strip leading period
    li_styles_by_level, li_styles_by_type, listparagraphs, all_list_styles = getListStylenames(styleconfig_dict)
    bookmakerstyles = vbastyleconfig_dict["bookmakerstyles"]
    valid_native_word_styles = cfg.valid_native_word_styles

    # These need to come first - otherwise contents (shapes) may keep blank paras from being blank
    # delete shapes, pictures, clip art etc
    report_dict, doc_root = doc_prepare.deleteObjects(report_dict, doc_root, cfg.shape_objects, "shapes")
    # delete bookmarks:
    report_dict = deleteBookmarks(report_dict, doc_root, cfg.bookmark_items)

    # delete any comments from docxml:
    report_dict, doc_root = doc_prepare.deleteObjects(report_dict, doc_root, cfg.comment_objects, "comment_ranges")
    # delete comments from commentsxml, commentsIds & commentsExtended wher present:
    comments_xmlfiles = {
        'comments_xml':cfg.comments_xml,
        'commentsExtended_xml':cfg.commentsExtended_xml,
        'commentsIds_xml':cfg.commentsIds_xml
    }
    for filename, filexml in comments_xmlfiles.iteritems():
        if os.path.exists(filexml):
            comments_tree = etree.parse(filexml)
            comments_root = comments_tree.getroot()
            xmlfile_dict[comments_root]=filexml
            report_dict, comments_root = doc_prepare.deleteObjects(report_dict, comments_root, cfg.comment_objects, "comments-%s" % filename)

    # remove blank paras from docxml, endnotes, footnotes -- only if we don't already have a critical blank para err
    # if "blank_container_para" not in report_dict and "blank_list_para" not in report_dict and "empty_section_start_para" not in empty_section_start_para:
    # for xml_root in xmlfile_dict:
    #     report_dict = removeBlankParas(report_dict, xml_root)
    report_dict = removeBlankParas(report_dict, doc_root, sectionnames, container_start_styles, container_end_styles, cfg.spacebreakstyles)
    if os.path.exists(cfg.endnotes_xml):
        report_dict = handleBlankParasInNotes(report_dict, endnotes_root, cfg.endnotestyle, cfg.endnote_ref_style, 'endnote', "Endnotes")
    if os.path.exists(cfg.footnotes_xml):
        report_dict = handleBlankParasInNotes(report_dict, footnotes_root, cfg.footnotestyle, cfg.footnote_ref_style, 'footnote', "Footnotes")

    # test / verify Container structures
    report_dict = checkContainers(report_dict, doc_root, sectionnames, container_start_styles, container_end_styles)
    # check list nesting
    report_dict = verifyListNesting(report_dict, doc_root, li_styles_by_level, li_styles_by_type, listparagraphs, all_list_styles)
    # get all Section Starts in the doc:
    report_dict = lxml_utils.sectionStartTally(report_dict, sectionnames, doc_root, "report")

    # check footnote / endnote para styles
    # rm footnote / endnote leading whitespace
    if os.path.exists(cfg.footnotes_xml):
        report_dict = checkEndnoteFootnoteStyles(footnotes_root, report_dict, cfg.footnotestyle, "footnote")
        report_dict = rmEndnoteFootnoteLeadingWhitespace(footnotes_root, report_dict, "footnote")
    if os.path.exists(cfg.endnotes_xml):
        report_dict = checkEndnoteFootnoteStyles(endnotes_root, report_dict, cfg.endnotestyle, "endnote")
        report_dict = rmEndnoteFootnoteLeadingWhitespace(endnotes_root, report_dict, "endnote")
    # log custom note markers for report
    refstyle_dict = {"endnote":cfg.endnote_ref_style, "footnote":cfg.footnote_ref_style}
    report_dict = flagCustomNoteMarks(doc_root, report_dict, refstyle_dict)

    # # log texts of titlepage-title paras
    report_dict = logTextOfParasWithStyleInSection(report_dict, doc_root, sectionnames, cfg.titlesection_stylename, cfg.titlestyle, "title_paras")
    # # log texts of titlepage-author paras
    report_dict = logTextOfParasWithStyleInSection(report_dict, doc_root, sectionnames, cfg.titlesection_stylename, cfg.authorstyle, "author_paras")
    # # # log texts of titlepage-logo paras
    report_dict = logTextOfParasWithStyleInSection(report_dict, doc_root, sectionnames, cfg.titlesection_stylename, cfg.logostyle, "logo_paras")
    # log texts of isbn-span runs
    report_dict = logTextOfRunsWithStyleInSection(report_dict, doc_root, sectionnames, cfg.copyrightsection_stylename, cfg.isbnstyle, "isbn_spans")

    # log texts of image_holders-holder paras, also checks for valid imageholder strings
    report_dict = stylereports.logTextOfParasWithStyle(report_dict, doc_root, cfg.imageholder_style, "image_holders", cfg.script_name)
    # log texts of inline illustration-holder runs, also checks for valid imageholder strings
    report_dict = stylereports.logTextOfRunsWithStyle(report_dict, doc_root, cfg.inline_imageholder_style, "image_holders", cfg.script_name)

    # check first para for non-section-Bookstyle
    booksection_stylename_short = lxml_utils.transformStylename(cfg.booksection_stylename)
    report_dict, firstpara = stylereports.checkFirstPara(report_dict, doc_root, [booksection_stylename_short], "non_section_BOOK_styled_firstpara")
    # check second para for non-section-startstyle
    report_dict = checkSecondPara(report_dict, doc_root, firstpara, sectionnames)

    # list all styles used in the doc
    # toggle 'allstyles_call_type' parameter to 'report' or 'validate' as needed:
    #   for rsuite styled docs, this means deleting non-Macmillan char styles or not
    allstyles_call_type = "validate"  # "report"
    report_dict = stylereports.getAllStylesUsed(report_dict, doc_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, allstyles_call_type, valid_native_word_styles, container_start_styles, container_end_styles)
    # running getAllStylesUsed on footnotes_root with 'runs_only = True' just to capture charstyles
    if os.path.exists(cfg.footnotes_xml):
        report_dict = stylereports.getAllStylesUsed(report_dict, footnotes_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, allstyles_call_type, valid_native_word_styles, container_start_styles, container_end_styles, True)
    if os.path.exists(cfg.endnotes_xml):
        report_dict = stylereports.getAllStylesUsed(report_dict, endnotes_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, allstyles_call_type, valid_native_word_styles, container_start_styles, container_end_styles, True)

    # removing any charstyles incorrectly / additionally applied to footnote / endnote reference markers in docxml
    #   footnotes
    report_dict = cleanNoteMarkers(report_dict, doc_root, cfg.footnote_ref_obj, cfg.footnote_ref_style, "footnote")
    #   endnotes
    report_dict = cleanNoteMarkers(report_dict, doc_root, cfg.endnote_ref_obj, cfg.endnote_ref_style, "endnote")

    # # add/update para index numbers
    logger.debug("Update all report_dict records with para_index")
    report_dict = lxml_utils.calcLocationInfoForLog(report_dict, doc_root, sectionnames, alt_roots)

    # create sorted version of "image_holders" list in reportdict based on para_index; for reports
    if "image_holders" in report_dict:
        report_dict["image_holders__sort_by_index"] = sorted(report_dict["image_holders"], key=lambda x: x['para_index'])


    # # # # # # Wrap up this parent function
    # write our changes back to xml files
    logger.debug("writing changes out to xml files")
    for xml_root in xmlfile_dict:
        os_utils.writeXMLtoFile(xml_root, xmlfile_dict[xml_root])

    logger.info("* * * ending rsuiteValidations function.")

    return report_dict



#---------------------  MAIN
# only run if this script is being invoked directly
# if __name__ == '__main__':
#
#     # set up debug log to console
#     logging.basicConfig(level=logging.DEBUG)
#
#     report_dict = {}
#     report_dict = rsuiteValidations(report_dict)
#
#     logger.debug("report_dict contents:  %s" % report_dict)
