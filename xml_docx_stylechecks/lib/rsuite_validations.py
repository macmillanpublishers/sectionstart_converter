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
    # # uncomment line below to use 'benchmark' decorator
    # from shared_utils.decorators import benchmark as benchmark


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
            lxml_utils.logForReport(report_dict, xml_root, para, report_category, paratxt, ['para_string'])
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
                    for key,value in x.items():
                        if x["para_id"] == this_para_id:
                            already_captured = True
            if already_captured == False:
                lxml_utils.logForReport(report_dict, xml_root, para, report_category, runtxt, ['para_id', 'para_string'])
    return report_dict

def checkSecondPara(report_dict, xml_root, firstpara, sectionnames):
    logger.info("* * * commencing checkSecondPara function...")
    pneighbors = lxml_utils.getNeighborParas(firstpara)
    secondpara_style = pneighbors['nextstyle']
    logger.debug("secondpara_style: {}".format(secondpara_style))
    if secondpara_style not in sectionnames:
        logger.warning("second para style is not a Section Start style, instead is: " + secondpara_style)
        lxml_utils.logForReport(report_dict, xml_root, pneighbors["next"], 'non_section_start_styled_secondpara', lxml_utils.getStyleLongname(secondpara_style))
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
                lxml_utils.logForReport(report_dict, xml_root, para, 'deleted_objects-bookmarks', 'deleted bookmark named {}'.format(w_name))
    # delete any orphaned bookmark_ends silently. (these should not exist, but why not clean-up in case?)
    end_searchstring = ".//{}".format(bookmark_items["bookmarkend_tag"])
    for bookmark_end in xml_root.findall(end_searchstring, wordnamespaces):
        w_name = bookmark_end.get('{%s}name' % wnamespace)
        bookmark_end.getparent().remove(bookmark_end)
        logger.debug("deleted orphaned bookmark_end named '%s'" % w_name)
    return report_dict

# table cell (w:tc) elements require at least one w:p
#   so we cannot bilthely delete
def handleBlankParasInTables(report_dict, xml_root, para, sectionnames, log_category="table_blank_para", log_description="blank para found in table cell", skip_logging=False):
    tablepara = False
    # have a provision to flag only _solo_ table paras, so we can aggressively rm any blank para possible.
    #   looking back at the original reqrement, that is both an edge case, and possibly overly intrusive
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
                    lxml_utils.logForReport(report_dict, xml_root, para, log_category, log_description, ['section_info'], sectionnames)
                tablepara = True
            else:
                # print len(pneighbors['prev'])
                logger.debug("blank tablecell para has a neighbor, bouncing back to std blank para handling")
        else:
            if skip_logging == False:
                lxml_utils.logForReport(report_dict, xml_root, para, log_category, log_description, ['section_info'], sectionnames)
            tablepara = True
    return report_dict, tablepara

# def removeBlankParas(xml_root, report_dict):
def removeBlankParas(report_dict, xml_root, sectionnames, container_start_styles, container_end_styles, spacebreakstyles):
    logger.info("* * * commencing removeBlankParas function...")
    specialparas = list(sectionnames) + container_start_styles + container_end_styles + spacebreakstyles
    allparas = xml_root.findall(".//w:p", wordnamespaces)
    for para in allparas:
        # get paras with no content
        if not lxml_utils.getParaTxt(para).strip(): # or para.text is None:
            # checking for solo tablecell paras first: extremely unlikely in the 1st two cases but possible in the 3rd:
            #   and rm'ing one breaks output doc
            report_dict, tablepara = handleBlankParasInTables(report_dict, xml_root, para, sectionnames)
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
                        lxml_utils.logForReport(report_dict, xml_root, para, 'removed_section_blank_para', sectionfullname)
                    elif parastyle in container_start_styles + container_end_styles:
                        lxml_utils.logForReport(report_dict, xml_root, para, 'removed_container_blank_para', '{}_{}'.format(lxml_utils.getStyleLongname(parastyle), section_info))
                    elif parastyle in spacebreakstyles:
                        lxml_utils.logForReport(report_dict, xml_root, para, 'removed_spacebreak_blank_para', '{}_{}'.format(lxml_utils.getStyleLongname(parastyle), section_info))

                # all paras are counted again so we gt a total for our count on the report
                lxml_utils.logForReport(report_dict, xml_root, para, 'removed_blank_para', 'removed {}-styled para'.format(parastyle))
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

def checkForNonSeparatorNotes(xml_root, separators, note_name, note_section):
    searchstring = ".//w:{}".format(note_name)
    type = '{%s}type' % wnamespace
    non_separator_note = False
    for note in xml_root.findall(searchstring, wordnamespaces):
        # skip separator notes
        if type in note.attrib and note.attrib[type] in separators:
            continue
        # one non-separator note is all we need to set bool
        non_separator_note = True
        break
    return non_separator_note

def checkForNotesSection(doc_root, endnotes_root, report_dict, separators, note_section_stylename):
    logger.info("* * * commencing checkForNotesSection function")
    # do we really have notes?
    notes_bool = checkForNonSeparatorNotes(endnotes_root, separators, "endnote", "Endnotes")
    if notes_bool == True:
        paras = lxml_utils.findParasWithStyle(lxml_utils.transformStylename(note_section_stylename), doc_root)
        if len(paras) == 0:
            firstpara = doc_root.find(".//*w:p", wordnamespaces)
            lxml_utils.logForReport(report_dict, None, None, 'missing_notes_section', 'Endnotes are present, Notes Section is not')
    return report_dict

def handleBlankParasInNotes(report_dict, xml_root, separators, note_stylename, noteref_stylename, note_name, note_section):
    logger.info("* * * commencing handleBlankParasInNotes function, for {}".format(note_section))
    # vars
    testpara_counter = 1
    searchstring = ".//w:{}".format(note_name)
    type = '{%s}type' % wnamespace
    # handle notes objects with no para children
    for note in xml_root.findall(searchstring, wordnamespaces):
        if len(note) == 0: # <== note object has no children (this may not occur naturally, but why not cover?)
            dummy_notepara, testpara_counter = createDummyNotePara(xml_root, testpara_counter, note_stylename, noteref_stylename)
            note.append(dummy_notepara)
            note_id = note.get('{%s}id' % wnamespace)
            lxml_utils.logForReport(report_dict, xml_root, dummy_notepara, 'found_empty_note', note_name, ['para_string'])
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
                report_dict, tablepara = handleBlankParasInTables(report_dict, xml_root, para, {}, '', '', True)
                if tablepara == False:
                    para.getparent().remove(para)
                    # make sure we don't re-add dummy text and log for report for multiblank paras in the same note!
                    if uniquenote == True:
                        dummy_notepara, testpara_counter = createDummyNotePara(xml_root, testpara_counter, note_stylename, noteref_stylename)
                        note.append(dummy_notepara)
                        lxml_utils.logForReport(report_dict, xml_root, dummy_notepara, 'found_empty_note', note_name, ['para_string'])
                        uniquenote = False
                    else:
                        # log above blank para removal for summary
                        lxml_utils.logForReport(report_dict, xml_root, para, 'removed_blank_para', 'excess blank para in empty {}'.format(note_name))
                elif tablepara == True:
                    uniquenote = False

    # handle blank paras within notes that have other paras with content
    #   the search above captured any notes with solo blank paras, this captures the remainder, which can be rm'd
    allparas = xml_root.findall(".//w:p", wordnamespaces)
    for para in allparas:
        if not lxml_utils.getParaTxt(para).strip(): # or para.text is None:
            # check if we are in a table; a table cell (w:tc) element requires at least one w:p
            report_dict, tablepara = handleBlankParasInTables(report_dict, xml_root, para, {}, 'table_blank_para_notes', 'blank para in table cell in {}'.format(note_name))
            if tablepara == False:
                # skip separator notes
                note = para.getparent()
                if type in note.attrib and note.attrib[type] in separators:
                    continue
                # log discovery
                note_id = note.get('{%s}id' % wnamespace)
                lxml_utils.logForReport(report_dict, xml_root, para, 'removed_blank_para', 'blank para in {} note with other text; note_id: {}'.format(note_name, note_id))
                # remove para
                para.getparent().remove(para)

    return report_dict

def checkContainers(report_dict, xml_root, sectionnames, container_start_styles, container_end_styles):
    logger.info("* * * commencing checkContainers function...")
    search_until_styles = list(sectionnames) + container_start_styles + container_end_styles
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
                lxml_utils.logForReport(report_dict, xml_root, start_para, 'illegal_style_in_table', lxml_utils.getStyleLongname(container_stylename), ['section_info'], sectionnames)
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
                lxml_utils.logForReport(report_dict, xml_root, start_para, 'container_error', lxml_utils.getStyleLongname(container_stylename), ['section_info'], sectionnames)
                if not pneighbors['nextstyle']:
                    logger.warning("container error - reached end of document before container-END styled para :( logging")
                elif pneighbors['nextstyle'] and pneighbors['nextstyle'] in sectionnames.keys():
                    logger.warning("container error - reached section-start style '%s' before container END styled para :(" % pneighbors['nextstyle'])
                else:
                    logger.warning("container error - reached container-start style '%s' before container END styled para :(" % pneighbors['nextstyle'])
    # doublecheck that no container Ends live in tables, or are orphaned:
    c_end_count = 0
    for end_stylename in container_end_styles:
        container_end_paras = lxml_utils.findParasWithStyle(end_stylename, xml_root)
        # containerstart_paras = xml_root.findall(searchstring, wordnamespaces)
        # loop through specific matched paras of a given style
        for end_para in container_end_paras:
            # check if we're in a table; if so, log it as an err and 'continue' to next para
            if end_para.getparent().tag == '{{{}}}tc'.format(wnamespace):
                lxml_utils.logForReport(report_dict, xml_root, end_para, 'illegal_style_in_table', lxml_utils.getStyleLongname(end_stylename), ['section_info'], sectionnames)
            else:
                c_end_count = c_end_count + 1
    # if the counts don't match, we have an orphan END para somewhere.
    if c_end_count != matched_cstart_count:
        logger.warning("found %s matched start containers, but %s Container_ends, indicating orphaned END" % (matched_cstart_count, c_end_count))
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
                        lxml_utils.logForReport(report_dict, xml_root, end_para, 'container_end_error', lxml_utils.getStyleLongname(end_stylename), ['section_info'], sectionnames)
                        if not pneighbors['prevstyle']:
                            logger.warning("orphaned container END - reached start of document before container-start para")
                        elif pneighbors['prevstyle'] and pneighbors['prevstyle'] in sectionnames.keys():
                            logger.warning("orphaned container END - reached section-start style '%s' before container START styled para" % pneighbors['prevstyle'])
                        else:
                            logger.warning("orphaned container END - reached another container-END para before container START para")

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
        "1": [],
        "2": [],
        "3": []
    }

    # using list comprehensions to import stylenames, for clearing leading period(s)
    unorderedlistparaslevel1 = [s[1:] for s in styleconfig_dict["unorderedlistparaslevel1"]]
    orderedlistparaslevel1 = [s[1:] for s in styleconfig_dict["orderedlistparaslevel1"]]
    li_styles_by_level[1] = unorderedlistparaslevel1 + orderedlistparaslevel1

    unorderedlistparaslevel2 = [s[1:] for s in styleconfig_dict["unorderedlistparaslevel2"]]
    orderedlistparaslevel2 = [s[1:] for s in styleconfig_dict["orderedlistparaslevel2"]]
    li_styles_by_level[2] = unorderedlistparaslevel2 + orderedlistparaslevel2

    unorderedlistparaslevel3 = [s[1:] for s in styleconfig_dict["unorderedlistparaslevel3"]]
    orderedlistparaslevel3 = [s[1:] for s in styleconfig_dict["orderedlistparaslevel3"]]
    li_styles_by_level[3] = unorderedlistparaslevel3 + orderedlistparaslevel3

    # get all list styles in 1 dict, with levels as their value for easy access
    all_list_styles = {i:1 for i in li_styles_by_level[1]}
    all_list_styles_2 = {i:2 for i in li_styles_by_level[2]}
    all_list_styles_3 = {i:3 for i in li_styles_by_level[3]}
    all_list_styles.update(all_list_styles_2)
    all_list_styles.update(all_list_styles_3)

    # track listparas in their own list
    listparagraphs = [s[1:] for s in styleconfig_dict["listparaparas"]]

    # get groups of list styles by type
    for liststyle in all_list_styles:
        if "bullet" in liststyle.lower():
            li_styles_by_type["bullet"].append(liststyle)
        elif "alpha" in liststyle.lower():
            li_styles_by_type["alpha"].append(liststyle)
        elif "unnum" in liststyle.lower():
            li_styles_by_type["unnum"].append(liststyle)
        elif "num" in liststyle.lower() and "unnum" not in liststyle.lower():
            li_styles_by_type["num"].append(liststyle)

    # tracking other paras that can interrupt lists without affecting nesting logic
    #   specifically extracts as per wdv-363
    nonlist_list_paras = [s[1:] for s in styleconfig_dict["extractparas"]]

    return li_styles_by_level, li_styles_by_type, listparagraphs, all_list_styles, nonlist_list_paras

def verifyListNesting(report_dict, xml_root, li_styles_by_level, li_styles_by_type, listparagraphs, list_styles, nonlist_list_paras, sectionnames):
    logger.info("* * * commencing verifyListNesting function...")
    all_list_styles = nonlist_list_paras + list(list_styles)
    #  cycle through all list styles by type
    for type in li_styles_by_type:
        for style in li_styles_by_type[type]:
            # search and return every list paragraph
            list_p_elements = lxml_utils.findParasWithStyle(style, xml_root)
            # # we are checking preceding para styles for compliance for every type of list
            # cycle through list paragraphs, get prev styles
            for list_p in list_p_elements:
                pneighbors = lxml_utils.getNeighborParas(list_p)
                prevstyle = pneighbors['prevstyle']
                # get style level from list_styles dict
                level = list_styles[style]

                # handle Unnested listparas with level > 1 (list nesting error)
                #   (Level1 list paragraphs (not list-para paragraphs) are valid when preceded by non-list items)
                #   examples: Body-Text > BL2, Title > NL3p, MainHead > UL1p
                if (level > 1 or (level == 1 and style in listparagraphs)) and prevstyle not in all_list_styles:
                    lxml_utils.logForReport(report_dict, xml_root, list_p, 'list_nesting_err', \
                        "'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)), \
                        ['section_info', 'para_string'], sectionnames)
                elif prevstyle in all_list_styles:
                    # calculate diff between current para list-level and preceding para list-level
                    if prevstyle in list_styles:
                        leveldiff = level - list_styles[prevstyle]  # < list_styles[prevstyle] = prev. para's level
                    elif prevstyle not in list_styles:
                        leveldiff = -1
                    # if prevstyle is a non_list_liststyle, or leveldiff is negative,
                    #   we need to traverse upwards till we find a real list style with a leveldiff >= 0 or a non-list style
                    #   example negative leveldiff: BL3p > NL2 <-- we need context from preceding paras to determine whether its valid
                    if leveldiff < 0:
                        para_tmp = list_p
                        while pneighbors['prevstyle'] and pneighbors['prevstyle'] in all_list_styles and leveldiff < 0:
                            # increment para upwards
                            para_tmp = pneighbors['prev']
                            pneighbors = lxml_utils.getNeighborParas(para_tmp)
                            # calc leveldiff:
                            if pneighbors['prevstyle'] and pneighbors['prevstyle'] in list_styles:
                                leveldiff = level - list_styles[pneighbors['prevstyle']]
                        # \/ L2 or L3 para never preceded by an L1; a nesting err (obscured by leading non_list_listpara(s))
                        #       (also pertains to L1, for list-para paragraphs)
                        if pneighbors['prevstyle'] not in all_list_styles and (level > 1 or (level == 1 and style in listparagraphs)):
                            lxml_utils.logForReport(report_dict, xml_root, list_p, 'list_nesting_err', \
                                "'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)), \
                                ['section_info', 'para_string'], sectionnames)
                            break
                        else:
                            prevstyle = pneighbors['prevstyle']
                    # now we take the difference in levels and flag/follow re: any problems:
                    # examples: BL1 > BL3, BL1 > NL3, BL1 > BL2p
                    #   (leveldiff == 1 is fine for non-listpara paras, like BL1 > BL2 or BL1p > NL2)
                    if leveldiff > 1 or (leveldiff == 1 and style in listparagraphs):
                        lxml_utils.logForReport(report_dict, xml_root, list_p, 'list_nesting_err', \
                            "'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)), \
                            ['section_info', 'para_string'], sectionnames)
                    # examples: BL1 > NL1, UL1p > BL1
                    #   (list change warning: only for list level = 1)
                    elif leveldiff == 0 and level == 1 and style not in listparagraphs and prevstyle not in li_styles_by_type[type]:
                        lxml_utils.logForReport(report_dict, xml_root, list_p, 'list_change_warning', \
                            "'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)), \
                            ['section_info', 'para_string'], sectionnames)
                    # examples: BL1 > NL1p, UL3p > NL3
                    #   (list change error for List paragraphs, nesting error for list-para paragraphs)
                    elif leveldiff == 0 and prevstyle not in li_styles_by_type[type]:
                        if style not in listparagraphs:
                            lxml_utils.logForReport(report_dict, xml_root, list_p, 'list_change_err', \
                                "'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)), \
                                ['section_info', 'para_string'], sectionnames)
                        else:
                            lxml_utils.logForReport(report_dict, xml_root, list_p, 'list_nesting_err', \
                                "'%s' para, preceded by: '%s' para" % (lxml_utils.getStyleLongname(style), lxml_utils.getStyleLongname(prevstyle)), \
                                ['section_info', 'para_string'], sectionnames)

    return report_dict

def duplicateSectionCheck(report_dict, section_array):
    logger.debug("* * * commencing duplicateSectionCheck function, for %s..." % section_array)
    # cycle through sections that should occur once only
    for sectionfullname in section_array:
        sectionshortname = lxml_utils.transformStylename(sectionfullname)
        section_count = 0
        # cycle through report dict sectionstart tally
        if 'section_start_found' in report_dict:
            for section_dict in report_dict['section_start_found']:
                if section_dict["description"] == sectionshortname:
                    section_count += 1
            # if we found more than one
            if section_count > 1:
                lxml_utils.logForReport(report_dict, None, None, 'too_many_section_para', "{}_{}".format(sectionfullname, section_count))
        else:
            logger.warning("function 'duplicateSectionCheck' did not find a 'section_start_found' list in report_dict")

    return report_dict

def checkForFMsectionsInBody(report_dict, fm_sectionnames, flex_sectionnames):
    logger.debug("* * * commencing checkForFMsectionsInBody function")
    # get shortnames of fm_sections
    fm_section_shortnames = [lxml_utils.transformStylename(x) for x in fm_sectionnames]
    flex_section_shortnames = [lxml_utils.transformStylename(x) for x in flex_sectionnames]
    body_begun = False

    if 'section_start_found' in report_dict:
        # cycle through sectionstarts found
        for section_dict in report_dict['section_start_found']:
            # look for first section not in fm or 'flex' list, to denote start of body
            if (section_dict["description"] not in fm_section_shortnames
                and section_dict["description"] not in flex_section_shortnames
                and body_begun == False):
                body_begun = True
            # then log any fm sections that appear once body's begun
            elif section_dict["description"] in fm_section_shortnames and body_begun == True:
                section_fullname = lxml_utils.getStyleLongname(section_dict["description"])
                # log section para_id if its present in dict, as part of description
                para_id = section_dict['para_id'] if 'para_id' in section_dict else ''
                lxml_utils.logForReport(report_dict, None, None, 'fm_section_in_body', '{}_{}'.format(section_fullname, para_id))
    return report_dict

# parse dict from checkMainheadsPerSection and log any multiple heads per section
def logMainheadMultiples(mainhead_dict, doc_root, report_dict, sectionnames):
    for stylename in mainhead_dict:
        for id, stylecount in mainhead_dict[stylename].items():
            if stylecount > 1:
                # get the section-para object from para id:
                searchstring = ".//*w:p[@w14:paraId='{}']".format(id)
                para = doc_root.find(searchstring, wordnamespaces)
                lxml_utils.logForReport(report_dict, doc_root, para, 'too_many_heading_para', '{}_{}'.format(stylename, stylecount), ['section_info', 'para_index'], sectionnames)
    return report_dict

def checkMainheadsPerSection(mainheadstyle_list, doc_root, report_dict, section_names, container_start_styles, container_end_styles):
    logger.info("* * * commencing checkMainheadsPerSection function, for {}...".format(mainheadstyle_list))
    # check document structure integrity via these previously reported items; skip if these errtypes present:
    if 'container_error' in report_dict or 'non_section_BOOK_styled_firstpara' in report_dict:
        logger.warning("Container error or non-section_Book styled first para; without section integrity we have to skip 'checkMainheadsPerSection' function")
    else:
        # for each occurence of a mainhead, find the parent section, keeping count per section/style in a dict.
        #   then we can check the dict for dupes and log to report_dict
        mainhead_dict = {}
        for stylename in mainheadstyle_list:
            mainhead_dict[stylename] = {}
            for para in lxml_utils.findParasWithStyle(lxml_utils.transformStylename(stylename), doc_root):
                section_id = getSectionOfNonContainerPara(para, doc_root, section_names, container_start_styles, container_end_styles)
                if section_id:
                    if section_id in mainhead_dict[stylename]:
                        mainhead_dict[stylename][section_id] += 1
                    else:
                        mainhead_dict[stylename][section_id] = 1
        logger.debug('contents of mainhead_dict: {}'.format(mainhead_dict))
        # separate function to process mainhead_dict
        report_dict = logMainheadMultiples(mainhead_dict, doc_root, report_dict, section_names)
    return report_dict

# occurences in a container don't count
def getSectionOfNonContainerPara(para, doc_root, section_names, container_start_styles, container_end_styles):
    tmp_para = para
    stylename = lxml_utils.getParaStyle(para)
    section_id = ''
    in_container = False
    no_previous_para = False
    # move upwards para by para until we find a sectionstart (or find no prev. para, indicating table or other nested p)
    while no_previous_para == False and stylename not in section_names:
        # this try helps us stop for nested text, like in a table
        try:
            tmp_para = tmp_para.getprevious()
        except:
            no_previous_para = True
        stylename = lxml_utils.getParaStyle(tmp_para)
        # we found a container start before we found a section-start or container-end; the init. para was in a container
        if stylename in container_start_styles and in_container == False:
            in_container = True
            logger.debug("this main head styled para is in a container, exit / skip")
            break
        # we are outside of (below) a container, but passing through one
        elif stylename in container_end_styles:
            in_container = True
        # we're done passing through a container
        elif stylename in container_start_styles and in_container == True:
            in_container = False
    # if we ended in a container then this is a styled para we don't want to count
    if in_container == False and no_previous_para == False:
        section_id = lxml_utils.getParaId(tmp_para, doc_root)
    return section_id

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
            lxml_utils.logForReport(report_dict, xml_root, para, 'improperly_styled_{}'.format(sectionname), parastyle, ['para_string'])
    return report_dict

# Logging for these could be optional, though so far in benchmark-tests it hasn't ever taken a full second as is
# @benchmark
def rmEndnoteFootnoteLeadingWhitespace(xml_root, report_dict, rootname):
    logger.info("* * * commencing rmEndnoteFootnoteLeadingWhitespace function, for %s..." % rootname)
    # now check leading text of note.
    allnotes = xml_root.findall(".//w:%s" % rootname, wordnamespaces)
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
                    lxml_utils.logForReport(report_dict, xml_root, para, '{}-leading_whitespace_rmd'.format(rootname), 'note ref: {} (tab removed)'.format(note_id))
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
                    lxml_utils.logForReport(report_dict, xml_root, para, '{}-leading_whitespace_rmd'.format(rootname), 'note ref: {} (blank run removed)'.format(note_id))
                break
            # here we have leading whitespace in the w:r/w:t itself that must be removed. re-setting text with lstrip()
            else:
                wt = wr.find(".//w:t", wordnamespaces)
                wt.text = wt.text.lstrip()
                lxml_utils.logForReport(report_dict, xml_root, para, '{}-leading_whitespace_rmd'.format(rootname), 'note ref: {}'.format(note_id))
                break

    return report_dict

def flagCustomNoteMarks(xml_root, report_dict, ref_style_dict):
    logger.info("* * * commencing flagCustomNoteMarks function...")
    for note_type, ref_style in ref_style_dict.items():
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
                lxml_utils.logForReport(report_dict, xml_root, para, 'custom_{}_mark'.format(note_type), "custom note marker: '{}', {} id: {}".format(reftext, note_type, ref_id))
    return report_dict

# wdv-389 specifies superscript styled ref marks in notes.
#   this fixes that, but does not correct superscript _formatted_ refmarks, which have rpr contents: <w:vertAlign w:val="superscript"/>
#   could add blanket handling of any non-custom notes not styled with std ref style pretty easily.
#   (this find skips custom notes, they do not have ref_el)
def fixSuperNoteMarks(xml_root, report_dict, superstyle, good_ref_style, note_type):
    logger.info("* * * commencing fixSuperNoteMarks function for %ss..." % note_type)
    ref_el = '{}Ref'.format(note_type)

    ref_searchstring = './/*w:{}'.format(ref_el)
    ref_els = xml_root.findall(ref_searchstring, wordnamespaces)

    for ref_el in ref_els:
        ref_el_run = ref_el.getparent()
        refstyle_el = ref_el_run.find(".//*w:rStyle", wordnamespaces)
        if refstyle_el is not None:
            refstyle = refstyle_el.get('{%s}val' % wnamespace)
            # print refstyle, superstyle # < for debug
            if refstyle == superstyle:

                attrib_style_key = '{%s}val' % wnamespace
                refstyle_el.set(attrib_style_key, good_ref_style)
                # log replacement
                para = lxml_utils.getParaParentofElement(ref_el_run)
                note_el = para.getparent()
                attrib_id_key = '{%s}id' % wnamespace
                ref_id = note_el.get(attrib_id_key)
                # re-using category from main body note mark-check
                lxml_utils.logForReport(report_dict, xml_root, para, 'note_markers_wrong_style', 'super_styled ref-mark in {}s, ref_id: {}'.format(note_type, ref_id))
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
                lxml_utils.logForReport(report_dict, xml_root, note_run.getparent(), 'note_markers_wrong_style', \
                    'restyled {} ref: no. {} (was styled as {})'.format(report_category, note_id, runstyle))
    return report_dict

def checkSymFonts(report_dict, xml_root, valid_symfonts):
    invalid_symfonts = []
    searchstring = './/w:sym/@w:font'
    sym_fontnames = xml_root.xpath(searchstring, namespaces=wordnamespaces)
    for fontname in sym_fontnames:
        if fontname not in valid_symfonts and fontname not in invalid_symfonts:
            invalid_symfonts.append(fontname)
    if invalid_symfonts:
        for sf in invalid_symfonts:
            # log only first occurence of this symfont
            if 'invalid_symfonts' not in report_dict or sf not in ([x['description'] for x in report_dict['invalid_symfonts']]):
                lxml_utils.logForReport(report_dict, xml_root, None, 'invalid_symfonts', sf)
    return report_dict


# @benchmark
def rsuiteValidations(report_dict):
    vbastyleconfig_json = cfg.vbastyleconfig_json
    styleconfig_json = cfg.styleconfig_json
    styles_xml = cfg.styles_xml
    doc_xml = cfg.doc_xml
    doc_root = lxml_utils.getXmlRootfromFile(doc_xml, 'doc.xml')
    # this is for writing out to any file where the xml_root is edited
    xmlfile_dict = {
        doc_root:doc_xml
        }
    if os.path.exists(cfg.endnotes_xml):
        endnotes_root = lxml_utils.getXmlRootfromFile(cfg.endnotes_xml, 'endnotes.xml')
        xmlfile_dict[endnotes_root]=cfg.endnotes_xml
    if os.path.exists(cfg.footnotes_xml):
        footnotes_root = lxml_utils.getXmlRootfromFile(cfg.footnotes_xml, 'footnotes.xml')
        xmlfile_dict[footnotes_root]=cfg.footnotes_xml

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
    li_styles_by_level, li_styles_by_type, listparagraphs, all_list_styles, nonlist_list_paras = getListStylenames(styleconfig_dict)
    bookmakerstyles = vbastyleconfig_dict["bookmakerstyles"]
    valid_native_word_styles = cfg.valid_native_word_styles
    # get decommissioned styles
    styleconfig_legacy_list = os_utils.readJSON(cfg.styleconfig_json)['legacy']
    decommissioned_styles = [lxml_utils.getStyleLongname(x[1:]) for x in styleconfig_legacy_list]

    # These need to come first - otherwise contents (shapes) may keep blank paras from being blank
    # delete shapes, pictures, clip art etc
    report_dict, doc_root = doc_prepare.deleteObjects(report_dict, doc_root, cfg.shape_objects, "shapes")
    # delete bookmarks:
    report_dict = deleteBookmarks(report_dict, doc_root, cfg.bookmark_items)
    #  delete shapes and bookmarks from notes xml if present
    if os.path.exists(cfg.endnotes_xml):
        report_dict, endnotes_root = doc_prepare.deleteObjects(report_dict, endnotes_root, cfg.shape_objects, "shapes")
        report_dict = deleteBookmarks(report_dict, endnotes_root, cfg.bookmark_items)
    if os.path.exists(cfg.footnotes_xml):
        report_dict, footnotes_root = doc_prepare.deleteObjects(report_dict, footnotes_root, cfg.shape_objects, "shapes")
        report_dict = deleteBookmarks(report_dict, footnotes_root, cfg.bookmark_items)

    # delete any comments from docxml:
    report_dict, doc_root = doc_prepare.deleteObjects(report_dict, doc_root, cfg.comment_objects, "comment_ranges")
    # delete comments from commentsxml, commentsIds & commentsExtended wher present:
    comments_xmlfiles = {
        'comments_xml':cfg.comments_xml,
        'commentsExtended_xml':cfg.commentsExtended_xml,
        'commentsIds_xml':cfg.commentsIds_xml
    }
    for filename, filexml in comments_xmlfiles.items():
        if os.path.exists(filexml):
            comments_tree = etree.parse(filexml)
            comments_root = comments_tree.getroot()
            xmlfile_dict[comments_root]=filexml
            report_dict, comments_root = doc_prepare.deleteObjects(report_dict, comments_root, cfg.comment_objects, "comments-%s" % filename)

    # remove blank paras from docxml, endnotes, footnotes -- only if we don't already have a critical blank para err
    # if "blank_container_para" not in report_dict and "blank_list_para" not in report_dict and "empty_section_start_para" not in empty_section_start_para:
    report_dict = removeBlankParas(report_dict, doc_root, sectionnames, container_start_styles, container_end_styles, cfg.spacebreakstyles)
    if os.path.exists(cfg.endnotes_xml):
        report_dict = handleBlankParasInNotes(report_dict, endnotes_root, cfg.note_separator_types, cfg.endnotestyle, cfg.endnote_ref_style, 'endnote', "Endnotes")
    if os.path.exists(cfg.footnotes_xml):
        report_dict = handleBlankParasInNotes(report_dict, footnotes_root, cfg.note_separator_types, cfg.footnotestyle, cfg.footnote_ref_style, 'footnote', "Footnotes")

    # test / verify Container structures
    report_dict = checkContainers(report_dict, doc_root, sectionnames, container_start_styles, container_end_styles)
    # check list nesting
    report_dict = verifyListNesting(report_dict, doc_root, li_styles_by_level, li_styles_by_type, listparagraphs, all_list_styles, nonlist_list_paras, sectionnames)
    # get all Section Starts in the doc:
    report_dict = lxml_utils.sectionStartTally(report_dict, sectionnames, doc_root, "report")
    # check for sections that should only appear once
    report_dict = duplicateSectionCheck(report_dict, [cfg.booksection_stylename, cfg.notessection_stylename])
    # check for FM sections in main body
    report_dict = checkForFMsectionsInBody(report_dict, cfg.fm_style_list, cfg.fm_flex_style_list)

    # check footnote / endnote para styles
    # rm footnote / endnote leading whitespace
    # handle note refs that are styled 'super' wdv-344
    if os.path.exists(cfg.footnotes_xml):
        report_dict = rmEndnoteFootnoteLeadingWhitespace(footnotes_root, report_dict, "footnote")
        report_dict = checkEndnoteFootnoteStyles(footnotes_root, report_dict, cfg.footnotestyle, "footnote")
        report_dict = fixSuperNoteMarks(footnotes_root, report_dict, cfg.superscriptstyle, cfg.footnote_ref_style, 'footnote')
    if os.path.exists(cfg.endnotes_xml):
        report_dict = rmEndnoteFootnoteLeadingWhitespace(endnotes_root, report_dict, "endnote")
        report_dict = checkEndnoteFootnoteStyles(endnotes_root, report_dict, cfg.endnotestyle, "endnote")
        report_dict = fixSuperNoteMarks(endnotes_root, report_dict, cfg.superscriptstyle, cfg.endnote_ref_style, 'endnote')
        # endnotes only: make sure Notes section is present
        report_dict = checkForNotesSection(doc_root, endnotes_root, report_dict, cfg.note_separator_types, cfg.notessection_stylename)
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
    report_dict = stylereports.logTextOfParasWithStyle(report_dict, doc_root, cfg.imageholder_style, "image_holders", sectionnames, cfg.script_name)
    # log texts of inline illustration-holder runs, also checks for valid imageholder strings
    report_dict = stylereports.logTextOfRunsWithStyle(report_dict, doc_root, cfg.inline_imageholder_style, "image_holders", sectionnames, cfg.script_name)

    # check first para for non-section-Bookstyle
    booksection_stylename_short = lxml_utils.transformStylename(cfg.booksection_stylename)
    report_dict, firstpara = stylereports.checkFirstPara(report_dict, doc_root, [booksection_stylename_short], "non_section_BOOK_styled_firstpara")
    # check second para for non-section-startstyle
    report_dict = checkSecondPara(report_dict, doc_root, firstpara, sectionnames)

    # check for more than one title, main head, or subtitle per section (not in containers)
    report_dict = checkMainheadsPerSection([cfg.titlestyle, cfg.subtitlestyle, cfg.mainheadstyle], doc_root, report_dict, sectionnames, container_start_styles, container_end_styles)

    # list all styles used in the doc
    # toggle 'allstyles_call_type' parameter to 'report' or 'validate' as needed:
    #   for rsuite styled docs, this means deleting non-Macmillan char styles or not
    allstyles_call_type = "validate"  # "report"
    report_dict = stylereports.getAllStylesUsed(report_dict, doc_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, allstyles_call_type, valid_native_word_styles, decommissioned_styles, container_start_styles, container_end_styles)
    # running getAllStylesUsed on footnotes_root with 'runs_only = True' just to capture charstyles
    if os.path.exists(cfg.footnotes_xml):
        report_dict = stylereports.getAllStylesUsed(report_dict, footnotes_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, allstyles_call_type, valid_native_word_styles, decommissioned_styles, container_start_styles, container_end_styles, True)
    if os.path.exists(cfg.endnotes_xml):
        report_dict = stylereports.getAllStylesUsed(report_dict, endnotes_root, styles_xml, sectionnames, macmillanstyledata, bookmakerstyles, allstyles_call_type, valid_native_word_styles, decommissioned_styles, container_start_styles, container_end_styles, True)

    # removing any charstyles incorrectly / additionally applied to footnote / endnote reference markers in docxml
    #   footnotes
    report_dict = cleanNoteMarkers(report_dict, doc_root, cfg.footnote_ref_obj, cfg.footnote_ref_style, "footnote")
    #   endnotes
    report_dict = cleanNoteMarkers(report_dict, doc_root, cfg.endnote_ref_obj, cfg.endnote_ref_style, "endnote")

    # check everywhere for invalid symfonts
    report_dict = checkSymFonts(report_dict, doc_root, cfg.valid_symfonts)
    if os.path.exists(cfg.footnotes_xml):
        report_dict = checkSymFonts(report_dict, footnotes_root, cfg.valid_symfonts)
    if os.path.exists(cfg.endnotes_xml):
        report_dict = checkSymFonts(report_dict, endnotes_root, cfg.valid_symfonts)


    # create sorted version of "image_holders" list in reportdict based on para_index; for reports
    if "image_holders" in report_dict:
        report_dict["image_holders__sort_by_index"] = sorted(report_dict["image_holders"], key=lambda x: x['para_index'])
    # create sorted version of "image_holders" list in reportdict based on para_index; for reports
    if "too_many_heading_para" in report_dict:
        report_dict["too_many_heading_para__sort_by_index"] = sorted(report_dict["too_many_heading_para"], key=lambda x: x['para_index'])
        report_dict["too_many_heading_para"] = []



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
