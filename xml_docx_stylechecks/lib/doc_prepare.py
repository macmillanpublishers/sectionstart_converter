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
xmlnamespace = cfg.xmlnamespace

# initialize logger
logger = logging.getLogger(__name__)


# #---------------------  METHODS

def getParaParentofElement(element):
    tmp_element = element
    while tmp_element.tag != '{%s}p' % wnamespace and tmp_element.tag != '{%s}body' % wnamespace:
        tmp_element = tmp_element.getparent()
    if tmp_element.tag == '{%s}body' % wnamespace:
        para = None
    else:
        para = tmp_element
    return para

def rmEmptyFirstLastParas(doc_root, report_dict):
    logger.info("* * * commencing rmEmptyFirstLastParas function...")
    allparas = doc_root.findall(".//w:p", wordnamespaces)
    # handle first para
    i=0
    firstpara = allparas[0+i]
    while not lxml_utils.getParaTxt(firstpara).strip():
        # print "lastpara", lxml_utils.getParaIndex(lastpara)
        lxml_utils.logForReport(report_dict,firstpara,"removed_empty_firstlast_para","first para")
        firstpara.getparent().remove(firstpara)
        # increment firstpara & recalculate
        i+=1
        firstpara = allparas[0+i]

    # now for last para
    j=1
    lastpara = allparas[len(allparas)-j]
    # print lxml_utils.getParaIndex(lastpara)
    while not lxml_utils.getParaTxt(lastpara).strip():
        # print "lastpara", lxml_utils.getParaIndex(lastpara)
        lxml_utils.logForReport(report_dict,lastpara,"removed_empty_firstlast_para","last para")
        lastpara.getparent().remove(lastpara)
        # increment lastpara & recalculate
        j+=1
        lastpara = allparas[len(allparas)-j]

    return report_dict

def rmCharStylesFromHeads(report_dict, doc_root, nocharstyle_headingstyles):
    for headingstyle in nocharstyle_headingstyles:
        paras = lxml_utils.findParasWithStyle(headingstyle, doc_root)
        for para in paras:
            rstyle = para.find(".//*w:rStyle", wordnamespaces)
            if rstyle is not None:
                # optional; log to report_dict
                rstylename = rstyle.get('{%s}val' % wnamespace)
                lxml_utils.logForReport(report_dict,para,"rm_charstyle_from_heading","Removed '%s' charstyle from '%s' heading." % (rstylename, headingstyle))
                # delete the runstyle!
                rstyle.getparent().remove(rstyle)
    return report_dict

def replaceSoftBreak(para, report_dict):
    logger.info("* * * commencing replaceSoftBreak function...")
    softbreaks = para.findall(".//*w:br", wordnamespaces)
    if softbreaks:
        replace_string = " "
        for softbreak in softbreaks:
            # create new text element with a single space as content:
            new_run_text = etree.Element("{%s}t" % wnamespace)
            new_run_text.text = replace_string
            # insert the w:t with single space next to the br
            softbreak.addnext(new_run_text)
            # rm the break
            softbreak.getparent().remove(softbreak)
        lxml_utils.logForReport(report_dict,para,"replaced_soft_break","replaced %s softbreaks with replace_string: '%s'" % (len(softbreaks),replace_string))
    return report_dict

def concatTitleParas(titlestyle, report_dict, doc_root):
    logger.info("* * * commencing concatTitleParas function...")
    # combine runs from titleparas
    searchstring = ".//*w:pStyle[@w:val='%s']" % titlestyle
    firsttitlepara = doc_root.find(searchstring, wordnamespaces).getparent().getparent()
    # replace softbreaks in the firsttitlepara
    report_dict = replaceSoftBreak(firsttitlepara, report_dict)
    # set vars
    titlestring = lxml_utils.getParaTxt(firsttitlepara)
    newtitlestring = titlestring
    pneighbors = lxml_utils.getNeighborParas(firsttitlepara)
    while pneighbors['nextstyle'] == titlestyle:
        # replace softbreaks in this title para (this will add spaces in cases where a softbreak was used)
        report_dict = replaceSoftBreak(pneighbors['next'], report_dict)
        # set newtitlestring
        newtitlestring = "%s %s" % (newtitlestring, lxml_utils.getParaTxt(pneighbors['next']))
        # newtitlestring = "{} {}".format(newtitlestring, nexttext)  # should review why this failed with unicode
        # increment, and delete this para
        tmp_para = pneighbors['next']
        pneighbors = lxml_utils.getNeighborParas(pneighbors['next'])
        tmp_para.getparent().remove(tmp_para)
    # if we have changes in the titlestring, remove existing contents and write the new full title as a new run
    if newtitlestring != titlestring:
        lxml_utils.addRunToPara(newtitlestring, firsttitlepara, True)
        # log for report (optional)
        lxml_utils.logForReport(report_dict,pneighbors['next'],"concatenated_extra_titlepara_and_removed",newtitlestring)
    return report_dict

def removeNonISBNsfromISBNspans(report_dict, doc_root, isbnstyle, isbnregex):
    logger.info("* * * commencing removeISBNspanfromNonISBN function...")
    isbnspan_runs = lxml_utils.findRunsWithStyle(isbnstyle, doc_root)
    for run in isbnspan_runs:
        runtxt = lxml_utils.getParaTxt(run)
        result = isbnregex.findall(runtxt)
        # capture the para number before we remove or edit:
        para = getParaParentofElement(run)
        # if isbn is found but there are extra chars, we need to yank them out
        if result:
            leadingtxt = str([x[0] for x in result][0])
            isbntxt = str([x[1] for x in result][0])
            followingtxt = str([x[3] for x in result][0])
            attrib_space_key = '{%s}space' % xmlnamespace
            if leadingtxt:
                # create new run element with leading content & addprevious to our para!
                new_run = etree.Element("{%s}r" % wnamespace)
                new_run_text = etree.Element("{%s}t" % wnamespace)
                new_run_text.set(attrib_space_key,"preserve")   # without this property set leading & trailing spaces aren't displayed in new runs in Word
                new_run_text.text = leadingtxt
                new_run.append(new_run_text)
                run.addprevious(new_run)
            if followingtxt:
                # create new run element with leading content & addnext to our para!
                new_run = etree.Element("{%s}r" % wnamespace)
                new_run_text = etree.Element("{%s}t" % wnamespace)
                new_run_text.set(attrib_space_key,"preserve")
                new_run_text.text = followingtxt
                new_run.append(new_run_text)
                run.addnext(new_run)
            # recreate current run with just the isbn:
            new_run = etree.Element("{%s}r" % wnamespace)
            # create runstyle and append to run element
            new_run_props = etree.Element("{%s}rPr" % wnamespace)
            new_run_props_style = etree.Element("{%s}rStyle" % wnamespace)
            new_run_props_style.attrib["{%s}val" % wnamespace] = isbnstyle
            new_run_props.append(new_run_props_style)
            new_run.append(new_run_props)
            # create runtext and append to run
            new_run_text = etree.Element("{%s}t" % wnamespace)
            new_run_text.text = isbntxt
            new_run.append(new_run_text)
            run.addnext(new_run)
            # remove the original run element & log!
            run.getparent().remove(run)
            lxml_utils.logForReport(report_dict,para,"rmd_nonisbn_from_isbnspan","non-isbn content present in span, split into new runs")
        # if no isbn is present, let's remove the rStyle (isbn span)
        else:
            # find the runstyle element and remove it
            rstyle = run.find(".//*w:rStyle", wordnamespaces)
            rstyle.getparent().remove(rstyle)
            # optional - log to report_dict:
            lxml_utils.logForReport(report_dict,para,"rmd_nonisbn_from_isbnspan","isbn not present in this isbn span, removed the whole thing")
    return report_dict

# # looks like this is not needed
# def scanDocForISBNs(isbnregex, doc_root):
#     for para in doc_root.findall(".//w:p", wordnamespaces):
#         paratext = lxml_utils.getParaTxt(para).strip()
#         result = isbnregex.findall(paratext)
#         if result:
#             print [x[1] for x in result]
#             # to capture these across runs, we would need to take this result, split it in half... wait it could be accross more than two runs. We would need to scan for runstyles & breaks
#             # then recreate the para.?!
            # noooo... you just need to know where it lives. Can you edit the raw xml instead of parsing via etree?
            # still a mess.

# insert required section start at the beginning of the doc if it's not already present
def insertRequiredSectionStart(sectionstartstyle, doc_root, contents, report_dict):
    logger.info("* * * commencing insertRequiredSectionStart function for (%s)..." % sectionstartstyle)
    if not lxml_utils.findParasWithStyle(sectionstartstyle, doc_root):
        # get 1st para of doc:
        first_para = doc_root.find(".//w:p", wordnamespaces)
        # insert my sspara at the beginnig of the doc
        lxml_utils.insertPara(sectionstartstyle, first_para, doc_root, contents, "before")
        # log that we added this!
        lxml_utils.logForReport(report_dict,first_para.getprevious(),"added_required_section_start","added '%s' to the beginning of the manuscript" % lxml_utils.getStyleLongname(sectionstartstyle))
    return report_dict

def removeTextWithCharacterStyle(report_dict, doc_root, style):
    styled_runs = lxml_utils.findRunsWithStyle(style, doc_root)
    for run in styled_runs:
        # get para for log
        para = getParaParentofElement(run)
        text = lxml_utils.getParaTxt(run)
        # remove the run with this style
        run.getparent().remove(run)
        # optional - log to report_dict:
        lxml_utils.logForReport(report_dict,para,"rmd_styled_runs","removed text (%s) styled with '%s'" % (text, style))
    return report_dict

def insertEbookISBN(report_dict, doc_root, copyrightsection_stylename, copyrightstyles, isbn, isbnstyle):
    logger.info("* * * commencing insertEbookISBN function...")
    sectionpara = lxml_utils.findParasWithStyle(copyrightsection_stylename, doc_root)[0]
    lastpara = sectionpara
    pneighbors = lxml_utils.getNeighborParas(sectionpara)
    while pneighbors["nextstyle"] in copyrightstyles:
        lastpara = pneighbors["next"]
        # increment the loop
        pneighbors = lxml_utils.getNeighborParas(lastpara)
    # add para
    lxml_utils.insertPara(copyrightstyles[0], lastpara, doc_root, isbn, "after")
    # add runstyle to the isbn:
    new_para = lastpara.getnext()
    new_text = new_para.find(".//*w:t", wordnamespaces)
    # create runstyle and append to run element
    new_run_props = etree.Element("{%s}rPr" % wnamespace)
    new_run_props_style = etree.Element("{%s}rStyle" % wnamespace)
    new_run_props_style.attrib["{%s}val" % wnamespace] = isbnstyle
    new_run_props.append(new_run_props_style)
    new_text.addprevious(new_run_props)
    # log for report
    lxml_utils.logForReport(report_dict,lastpara.getnext(),"added_ebook_isbn","added '%s'" % isbn)
    return report_dict


def getBookInfoFromExternalLookups(bookinfo_json, config_json):
    bookinfo, import_dict = {}, {}
    if os.path.exists(bookinfo_json):
        import_dict = os_utils.readJSON(bookinfo_json)
        bookinfo["isbn"] = import_dict["isbn"]
        bookinfo["author"] = import_dict["author"]
        bookinfo["title"] = import_dict["title"]
    elif os.path.exists(config_json):
        import_dict = os_utils.readJSON(config_json)
        bookinfo["isbn"] = import_dict["ebookid"]
        bookinfo["author"] = import_dict["author"]
        bookinfo["title"] = import_dict["title"]
    else:
        bookinfo["isbn"] = ""
        bookinfo["author"] = ""
        bookinfo["title"] = ""
        logger.warn("Unable to find meta-info for this manuscript")
    return bookinfo

# inserts author & or title paras if they are not present
def insertBookinfo(report_dict, doc_root, stylename, leadingpara_style, bookinfo_item):
    logger.info("commencing insertBookinfo for: '%s'..." % stylename)
    if not lxml_utils.findParasWithStyle(stylename, doc_root):
        # get LEADINGPARASTYLE style which should already exist
        leadingpara = lxml_utils.findParasWithStyle(leadingpara_style, doc_root)[0]
        if leadingpara is not None and bookinfo_item:
            lxml_utils.insertPara(stylename, leadingpara, doc_root, bookinfo_item, "after")
            lxml_utils.logForReport(report_dict,leadingpara.getnext(),"added_required_book_info","added '%s' paragraph with content: '%s'" % (lxml_utils.getStyleLongname(stylename), bookinfo_item))
        else:
            if not bookinfo_item:
                lxml_utils.logForReport(report_dict,leadingpara.getnext(),"added_required_book_info","added '%s' paragraph with content: '%s'" % (lxml_utils.getStyleLongname(stylename), bookinfo_item))
                logger.warn("'%s' was missing from the manuscript but could not be auto-inserted because %s lookup value was empty." % (stylename, bookinfo_item))
            if leadingpara is None:
                logger.warn("Could not find required %s styled 'leading_para' to insert bookinfo field: '%s'" % (leadingpara_style, bookinfo_item))
    return report_dict

# Tring to capture shapes, inserted clipart, etc, with two types of elements, as per:
#   http://officeopenxml.com/drwOverview.php
# Re: Section breaks, if one is inserted, another one is created at the end of the body (invisible to user). Not reporting this one, bu it's getting rm'd too.
def deleteShapesAndBreaks(report_dict, doc_root, objects_to_delete):
    logger.info("* * * commencing deleteShapesAndBreaks function...")
    for shape in objects_to_delete:
        searchstring = ".//*{}".format(shape)
        for element in doc_root.findall(searchstring, wordnamespaces):
            # get para for report (before we delete theelement!):
            para = getParaParentofElement(element)
            # remove element
            element.getparent().remove(element)
            # optional - log to report_dict
            if para is not None:
                lxml_utils.logForReport(report_dict,para,"deleted_shapes_and_sectionbreaks","deleted %s item" % shape)
    return report_dict

def rmNonPrintingHeads(report_dict, doc_root, nonprintingheads):
    logger.info("* * * commencing rmNonPrintingHeads function...")
    for style in nonprintingheads:
        paras = lxml_utils.findParasWithStyle(style, doc_root)
        for para in paras:
            # get text for log
            paratxt = lxml_utils.getParaTxt(para)
            # log the para and remove
            lxml_utils.logForReport(report_dict,para,"removed_nonprintinghead_para","removed a '%s' with contents: '%s'" % (style, paratxt))
            para.getparent().remove(para)
    return report_dict

def docPrepare(report_dict):
    logger.info("* * * commencing docPrepare function...")
    # local vars
    bookinfo_json = os.path.join(cfg.tmpdir, "book_info.json")
    config_json = os.path.join(cfg.tmpdir, "config.json")
    section_start_rules_json = cfg.section_start_rules_json
    styleconfig_json = cfg.styleconfig_json
    doc_xml = cfg.doc_xml
    doc_tree = etree.parse(doc_xml)
    doc_root = doc_tree.getroot()
    isbnstyle = lxml_utils.transformStylename(cfg.isbnstyle)
    # isbnregex = re.compile(r"(97[89]((\D?\d){10}))")
    isbnregex = re.compile(r"(^.*?)(97[89](\D?\d){10})(.*?$)")

    logger.info("reading in json resource files")
    # read rules & heading-style list from JSONs
    section_start_rules = os_utils.readJSON(section_start_rules_json)
    styleconfig_dict = os_utils.readJSON(styleconfig_json)

    # set vars based on JSON imports
    headingstyles = [classname[1:] for classname in styleconfig_dict["headingparas"]]
    bookinfo = getBookInfoFromExternalLookups(bookinfo_json, config_json)

    # get Section Start names & styles from sectionstartrules
    sectionnames = lxml_utils.getAllSectionNames(section_start_rules)

    # delete shapes, pictures, clip art etc
    report_dict = deleteShapesAndBreaks(report_dict, doc_root, cfg.objects_to_delete)

    # remove character styles from headings in list
    report_dict = rmCharStylesFromHeads(report_dict, doc_root, cfg.nocharstyle_headingstyles)
    report_dict = rmCharStylesFromHeads(report_dict, doc_root, headingstyles)

    # # # setup required frontmatter
    # remove non-isbn chars from ISBN span
    report_dict = removeNonISBNsfromISBNspans(report_dict, doc_root, isbnstyle, isbnregex)
    # make sure Copyright page exists, with isbn from lookup
    report_dict = insertRequiredSectionStart(cfg.copyrightsection_stylename, doc_root, "Copyright", report_dict)
    # # rm existing styled ISBNs and append isbn from lookup to after last Copyright Page section
    if bookinfo["isbn"]:
        report_dict = removeTextWithCharacterStyle(report_dict, doc_root, isbnstyle)
        report_dict = insertEbookISBN(report_dict, doc_root, cfg.copyrightsection_stylename, cfg.copyrightstyles, bookinfo["isbn"], isbnstyle)
    else:
        logger.warn("No lookup-ISBN available, skipping ISBN cleanup & auto-insertion.")
    # make sure Titlepage exists: leaving contents empty, that will get auto-added
    report_dict = insertRequiredSectionStart(cfg.titlesection_stylename, doc_root, "", report_dict)
    # add author info to titlepage if it's not present
    report_dict = insertBookinfo(report_dict, doc_root, lxml_utils.transformStylename(cfg.authorstyle), cfg.titlesection_stylename, bookinfo["author"])
    # add title info to titlepage if it's not present
    report_dict = insertBookinfo(report_dict, doc_root, lxml_utils.transformStylename(cfg.titlestyle), cfg.titlesection_stylename, bookinfo["title"])

    # concatenate consecutive titleparas and remove softbreaks
    report_dict = concatTitleParas(lxml_utils.transformStylename(cfg.titlestyle), report_dict, doc_root)

    # # # tally and repair section start paras & their contents
    # get all Section Starts paras in the doc, add content to each para as needed:
    report_dict = lxml_utils.sectionStartTally(report_dict, sectionnames, doc_root, "insert", headingstyles)

    # remove first or last paras if they contain only white space
    #   (this has to come after sectionStartTally function, otherwise it may rip out empty Section Start para at beginning of doc)
    report_dict = rmEmptyFirstLastParas(doc_root, report_dict)

    # # # Commenting this out until we have alternate padding in place for these items.
    # # # See Asana task: https://app.asana.com/0/405970099375554/453275643539715
    # now that we captured their contents (where needed), we can get rid of Nonprinting head paras
    # report_dict = rmNonPrintingHeads(report_dict, doc_root, cfg.nonprintingheads)

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

# should check on scope of passing and returning reort_dict. it may not need to be returned all the time. Will have to do some scope testing to test calls from other scripts & / or standalone.
# newtitlestring = "{} {}".format(newtitlestring, nexttext)
# this line caused issues with a title with ' and ; in it...  where else might I have used 'format' where I may get in trouble?
