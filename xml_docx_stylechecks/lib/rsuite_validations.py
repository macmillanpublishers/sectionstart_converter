######### IMPORT PY LIBRARIES
import os
import json
import sys
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
    import doc_prepare
    cfg = imp.load_source('cfg', cfgpath)
    os_utils = imp.load_source('os_utils', osutilspath)
    lxml_utils = imp.load_source('lxml_utils', lxmlutilspath)
else:
    import cfg
    import lib.doc_prepare as doc_prepare
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

def getContainerStarts(styleconfig_dict):
    container_start_styles = []
    for category in styleconfig_dict["containerparas"]:
        container_start_styles = container_start_styles + styleconfig_dict["containerparas"][category]
    # strip leading period
    container_start_styles = [s[1:] for s in container_start_styles]
    return container_start_styles

def removeBlankParas(xml_root, report_dict):
    logger.info("* * * commencing removeBlankParas function...")
    allparas = xml_root.findall(".//w:p", wordnamespaces)
    for para in allparas:
        if not lxml_utils.getParaTxt(para).strip():
            parastyle = lxml_utils.getParaStyle(para)
            lxml_utils.logForReport(report_dict,xml_root,para,"removed_blank_para","style was %s" % parastyle)
            para.getparent().remove(para)
    return report_dict

def checkContainers(report_dict, xml_root, sectionnames, container_start_styles, container_end_styles):
    logger.info("* * * commencing checkContainers function...")
    search_until_styles = sectionnames.keys() + container_start_styles + container_end_styles
    # loop through searching for different container start styles
    for container_stylename in container_start_styles:
        searchstring = ".//*w:pStyle[@w:val='%s']" % container_stylename
        containerstart_paras = xml_root.findall(searchstring, wordnamespaces)
        # loop through specific matched paras of a given style
        for start_para_pStyle in containerstart_paras:
            start_para = start_para_pStyle.getparent().getparent()
            pneighbors = lxml_utils.getNeighborParas(start_para)
            para_tmp = start_para
            # take a quick moment to log any blank patas we find, since those will generate errors:
            paratxt = lxml_utils.getParaTxt(start_para)
            if not paratxt.strip():
                lxml_utils.logForReport(report_dict,xml_root,start_para,"blank_container_para","style is %s" % container_stylename)

            # scan styles of next paras until we match something to stop at
            while pneighbors['nextstyle'] and pneighbors['nextstyle'] not in search_until_styles:
                # increment para downwards
                para_tmp = pneighbors['next']
                pneighbors = lxml_utils.getNeighborParas(para_tmp)

            # figure out whether we matched an END style or something else
            if pneighbors['nextstyle'] and pneighbors['nextstyle'] in container_end_styles:
                logger.debug("found a container end style before section start, container start or document end")
            elif not pneighbors['nextstyle']:
                logger.warn("reached end of document before container END styled para :( logging")
                lxml_utils.logForReport(report_dict,xml_root,start_para,"container_error-eof","encountered end of document before END styled para")
            elif pneighbors['nextstyle'] and pneighbors['nextstyle'] in sectionnames.keys():
                logger.warn("reached section-start style before container END styled para :( logging")
                lxml_utils.logForReport(report_dict,xml_root,start_para,"container_error-newsection","encountered '%s' para before END styled para" % pneighbors['nextstyle'])
            else:
                logger.warn("reached container-start style before container END styled para :( logging")
                lxml_utils.logForReport(report_dict,xml_root,start_para,"container_error-newcontainer","encountered '%s' para before END styled para" % pneighbors['nextstyle'])

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
            searchstring = ".//*w:pStyle[@w:val='%s']" % style
            list_pStyles = xml_root.findall(searchstring, wordnamespaces)
            # cycle through matches, get prev styles
            for list_pStyle in list_pStyles:
                list_p = list_pStyle.getparent().getparent()
                pneighbors = lxml_utils.getNeighborParas(list_p)
                prevstyle = pneighbors['prevstyle']
                # if blank para, log as an error:
                paratxt = lxml_utils.getParaTxt(list_p)
                if not paratxt.strip():
                    lxml_utils.logForReport(report_dict,xml_root,list_p,"blank_list_para","style is %s" % style)

                # determine our style's level
                for level in li_styles_by_level:
                    if style in li_styles_by_level[level]:
                        # check preceding parastyles for all listpara instances: if not same style of matching list type/level (para or non-para), log err
                        if style in listparagraphs:
                            if prevstyle not in li_styles_by_level[level] or prevstyle not in li_styles_by_type[type]:
                                lxml_utils.logForReport(report_dict,xml_root,list_p,"list_nesting_err","'%s', preceded by: %s" % (style, prevstyle))
                        # all other checks are for list non-paragraphs
                        else:
                            # if list non-paras preceding paragraph is the same level but different type, issue warning
                            if prevstyle in li_styles_by_level[level] and prevstyle not in li_styles_by_type[type]:
                                lxml_utils.logForReport(report_dict,xml_root,list_p,"list_change_warning","'%s', preceded by: %s" % (style, prevstyle))
                            # list level 2 non-para, if preceded by list para, must be preceded by list level 1 or 2 (any)
                            elif level == "level2" and prevstyle in all_list_styles and prevstyle not in li_styles_by_level["level1"] + li_styles_by_level["level2"]:
                                lxml_utils.logForReport(report_dict,xml_root,list_p,"list_nesting_err","'%s', preceded by: %s" % (style, prevstyle))
                            # list level 3 non-para, if preceded by list para, must be preceded by list level 2 or 3 (any)
                            elif level == "level3" and prevstyle in all_list_styles and prevstyle not in li_styles_by_level["level2"] + li_styles_by_level["level3"]:
                                lxml_utils.logForReport(report_dict,xml_root,list_p,"list_nesting_err","'%s', preceded by: %s" % (style, prevstyle))

    return report_dict

def checkEndnoteFootnoteStyles(xml_root, report_dict, note_style, sectionname):
    logger.info("* * * commencing checkEndnoteFootnoteStyles function, for %s..." % sectionname)
    # first check styles of paras
    allparas = xml_root.findall(".//w:p", wordnamespaces)
    for para in allparas:
        parastyle = para.find(".//w:pStyle", wordnamespaces).get('{%s}val' % wnamespace)
        # print parastyle
        if parastyle != note_style:
            note_id = para.getparent().get('{%s}id' % wnamespace)
            lxml_utils.logForReport(report_dict,xml_root,para,"improperly_styled_%s" % sectionname,"note ref: %s" % note_id)
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


def rsuiteValidations(report_dict):
    vbastyleconfig_json = cfg.vbastyleconfig_json
    styleconfig_json = cfg.styleconfig_json
    doc_xml = cfg.doc_xml
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    endnotes_tree = etree.parse(cfg.endnotes_xml)
    endnotes_root = endnotes_tree.getroot()
    footnotes_tree = etree.parse(cfg.footnotes_xml)
    footnotes_root = footnotes_tree.getroot()
    comments_tree = etree.parse(cfg.comments_xml)
    comments_root = comments_tree.getroot()
    commentsExtended_tree = etree.parse(cfg.commentsExtended_xml)
    commentsExtended_root = commentsExtended_tree.getroot()
    commentsIds_tree = etree.parse(cfg.commentsIds_xml)
    commentsIds_root = commentsIds_tree.getroot()

    xmlfile_dict = {
        doc_root:doc_xml,
        endnotes_root:cfg.endnotes_xml,
        footnotes_root:cfg.footnotes_xml
        }
    commentfiles_dict = {
        # comments_root:cfg.comments_xml,
        # commentsExtended_root:cfg.commentsExtended_xml,
        commentsIds_root:cfg.commentsIds_xml
        }

    footnotestyle = cfg.footnotestyle
    endnotestyle = cfg.endnotestyle


    # get Section Start names & styles from vbastyleconfig_json
    #    Could pull styles from macmillan.json  with "Section-" if I don't want to use vbastyleconfig_json
    vbastyleconfig_dict = os_utils.readJSON(vbastyleconfig_json)
    sectionnames = lxml_utils.getAllSectionNamesFromVSC(vbastyleconfig_dict)
    # get Container styles (shortnames only) from styleconfig_dict
    #   Could pull them from macmillan.json as well, searching for ALL CAPS.. though would need to get container-end from style_config
    #   Also might be able to reverse engineer Container longnames
    styleconfig_dict = os_utils.readJSON(styleconfig_json)
    container_start_styles = getContainerStarts(styleconfig_dict)
    container_end_styles = [s[1:] for s in styleconfig_dict["containerendparas"]] # <--strip leading period
    li_styles_by_level, li_styles_by_type, listparagraphs, all_list_styles = getListStylenames(styleconfig_dict)

    # test / verify Container structures
    report_dict = checkContainers(report_dict, doc_root, sectionnames, container_start_styles, container_end_styles)
    # check list nesting
    report_dict = verifyListNesting(report_dict, doc_root, li_styles_by_level, li_styles_by_type, listparagraphs, all_list_styles)

    # remove blank paras from docxml, endnotes, footnotes
    for xml_root in xmlfile_dict:
        report_dict = removeBlankParas(xml_root, report_dict)

    # check footnote / endnote para styles
    report_dict = checkEndnoteFootnoteStyles(footnotes_root, report_dict, footnotestyle, "footnote")
    report_dict = checkEndnoteFootnoteStyles(endnotes_root, report_dict, endnotestyle, "endnote")
    # rm footnote / endnote leading whitespace
    report_dict = rmEndnoteFootnoteLeadingWhitespace(footnotes_root, report_dict, "footnote")
    report_dict = rmEndnoteFootnoteLeadingWhitespace(endnotes_root, report_dict, "endnote")

    # delete shapes, pictures, clip art etc
    report_dict = doc_prepare.deleteObjects(report_dict, doc_root, cfg.shape_objects, "shapes")
    # delete bookmarks:
    report_dict = doc_prepare.deleteObjects(report_dict, doc_root, cfg.bookmark_objects, "bookmarks")
    # delete comments from docxml, commentsxml, commentsIds & commentsExtended:
    report_dict = doc_prepare.deleteObjects(report_dict, doc_root, cfg.comment_objects, "comment_ranges")
    report_dict = doc_prepare.deleteObjects(report_dict, comments_root, cfg.comment_objects, "comments-comment_xml")
    report_dict = doc_prepare.deleteObjects(report_dict, commentsExtended_root, cfg.comment_objects, "commentEx-commentEx_xml")
    report_dict = doc_prepare.deleteObjects(report_dict, commentsIds_root, cfg.comment_objects, "commentcid-commentsIds_xml")


    # # # # # # Wrap up this parent function
    # write our changes back to xml files
    logger.debug("writing changes out to xml files")
    for xml_root in xmlfile_dict:
        os_utils.writeXMLtoFile(xml_root, xmlfile_dict[xml_root])
    for xml_root in commentfiles_dict:
        os_utils.writeXMLtoFile(xml_root, commentfiles_dict[xml_root])


    # # add/update para index numbers <-- commenting for rsuite_validations: its getting run in stylereporter too
    # logger.debug("Update all report_dict records with para_index-")
    # report_dict = lxml_utils.calcLocationInfoForLog(report_dict, doc_root, sectionnames, [footnotes_root, endnotes_root])

    logger.info("* * * ending rsuiteValidations function.")

    return report_dict



#---------------------  MAIN
# only run if this script is being invoked directly
if __name__ == '__main__':

    # set up debug log to console
    logging.basicConfig(level=logging.DEBUG)

    report_dict = {}
    report_dict = rsuiteValidations(report_dict)

    logger.debug("report_dict contents:  %s" % report_dict)

# should check on scope of passing and returning reort_dict. it may not need to be returned all the time. Will have to do some scope testing to test calls from other scripts & / or standalone.
# newtitlestring = "{} {}".format(newtitlestring, nexttext)
# this line caused issues with a title with ' and ; in it...  where else might I have used 'format' where I may get in trouble?
