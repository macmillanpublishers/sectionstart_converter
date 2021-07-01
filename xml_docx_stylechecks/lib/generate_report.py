######### IMPORT PY LIBRARIES

import os
import sys
import logging
# make sure to install lxml: sudo pip install lxml
from lxml import etree


######### IMPORT LOCAL MODULES
if __name__ == '__main__':
    # to go up a level to read cfg (and other files) when invoking from this script (for testing).
    cfgpath = os.path.join(sys.path[0], '..', 'cfg.py')
    osutilspath = os.path.join(sys.path[0], '..', 'shared_utils', 'os_utils.py')
    # lxmlutilspath = os.path.join(sys.path[0], '..', 'shared_utils', 'lxml_utils.py')
    import imp
    cfg = imp.load_source('cfg', cfgpath)
    os_utils = imp.load_source('os_utils', osutilspath)
    lxml_utils = imp.load_source('lxml_utils', lxmlutilspath)
else:
    import cfg
    import shared_utils.os_utils as os_utils
    import shared_utils.lxml_utils as lxml_utils
import report_recipe


######### LOCAL DECLARATIONS
styles_xml = cfg.styles_xml

# Local namespace vars
wnamespace = cfg.wnamespace
w14namespace = cfg.w14namespace
wordnamespaces = cfg.wordnamespaces

# initialize logger
logger = logging.getLogger(__name__)


# #---------------------  METHODS
def buildReport(report_dict, textreport_list, scriptname, stylenamemap, recipe_item, errorlist, warninglist, notelist, validator_warnings):
    # make sure we're not supposed to skip this recipe_item as per "include_for" key
    if "include_for" in recipe_item and scriptname in recipe_item["include_for"]:
        tmptextlist = []
        tmptextlist.append("")  # start with a blank line
        # add formatted title if present
        if "title" in recipe_item and recipe_item["title"]:
            formattedtitle = "\n{:-^80}".format(" "+recipe_item["title"]+" ")
            tmptextlist.append(formattedtitle)
        # add text if present
        if "text" in recipe_item and recipe_item["text"]:
            tmptextlist.append(recipe_item["text"])
        # see if "dict_category_name" is present
        if "dict_category_name" in recipe_item and recipe_item["dict_category_name"]:
            # if the category is in report_dict and has contents, proceed, else, print an alternate & log err as needed
            if recipe_item["dict_category_name"] in report_dict and report_dict[recipe_item["dict_category_name"]]:
                if "v_warning_banner" in recipe_item and recipe_item["v_warning_banner"]:
                    if recipe_item["v_warning_banner"] not in validator_warnings:
                        validator_warnings.append("- %s" % recipe_item['v_warning_banner'])
                    # validator_warnings = True
                for item in report_dict[recipe_item["dict_category_name"]]:
                    if recipe_item["dict_category_name"] == "Macmillan_style_first_use":
                        new_section_text = recipe_item["new_section_text"].format(parent_section_start_content='"'+item['parent_section_start_content'].encode('utf-8')+'"', \
                            parent_section_start_type=lxml_utils.getStyleLongname(item['parent_section_start_type'], stylenamemap))
                        if new_section_text not in tmptextlist:
                            tmptextlist.append(new_section_text)
                    # line = "{:<40}:{:>30}".format("parent_section_start_type","'parent_section_start_content'")
                    newline = recipe_item["line_template"].format(description=item['description'].encode('utf-8'), para_string='"'+item['para_string'].encode('utf-8')+'"', \
                        parent_section_start_content='"'+item['parent_section_start_content'].encode('utf-8')+'"', \
                        parent_section_start_type=lxml_utils.getStyleLongname(item['parent_section_start_type'], stylenamemap), para_index=item['para_index'])
                    # newline = line.replace("zdescription", item['description']).replace("para_string", item['para_string']).replace("parent_section_start_content", item['parent_section_start_content'])
                    # newline = newline.replace("parent_section_start_type", lxml_utils.getStyleLongname(item['parent_section_start_type'], stylenamemap))
                    tmptextlist.append(newline)
                    if "badnews" in recipe_item and recipe_item["badnews"] == 'any':
                        # for complicated descriptions, where we are transferring 2 pieces of info, we have to split them here for the errstring:
                        if "_" in item['description']:
                            descriptionA, descriptionB=item['description'].split("_",1)
                        else:
                            descriptionA, descriptionB = "", ""
                        # now we set err strings from report_recipe for toplist items
                        new_errstring = recipe_item["errstring"].format(description=item['description'].encode('utf-8'), \
                            para_string='"'+item['para_string'].encode('utf-8')+'"', \
                            parent_section_start_content='"'+item['parent_section_start_content'].encode('utf-8')+'"', \
                            parent_section_start_type=lxml_utils.getStyleLongname(item['parent_section_start_type'], stylenamemap), \
                            para_index=item['para_index'], \
                            count=len(report_dict[recipe_item["dict_category_name"]]), \
                            section_count=sum(1 for sectiontxt in report_dict[recipe_item["dict_category_name"]] if sectiontxt['parent_section_start_content'] == item['parent_section_start_content']), \
                            notes_count=sum(1 for notestype in report_dict[recipe_item["dict_category_name"]] if notestype['parent_section_start_type'] == item['parent_section_start_type']), \
                            descriptionA=descriptionA.encode('utf-8'), \
                            descriptionB=descriptionB.encode('utf-8'), \
                            valid_file_extensions=cfg.imageholder_supported_ext)
                        if item['para_index'] == 'tablecell_para':# and not("summary" in recipe_item and recipe_item["summary"] == True):
                            if 'suppress_table_note' not in recipe_item or recipe_item['suppress_table_note'] != True:
                                new_errstring += "  (< this item is from a table)"
                        # added 'summary' key so we could specify whether to summarize warnings or notes:
                        #   default is warnings are listed singly, notes are summarized
                        alerttypes = {
                            'note': notelist,
                            'warning': warninglist,
                            'error': errorlist
                        }
                        for alertname, alertlist in alerttypes.iteritems():
                            if "badnews_type" in recipe_item:
                                if recipe_item["badnews_type"] == alertname:
                                    # 'notes' are expected to summarize by default
                                    if alertname == 'note' or ("summary" in recipe_item and recipe_item["summary"] == True):
                                        if new_errstring not in alertlist:
                                            alertlist.append(new_errstring)
                                            break
                                    # if we don't specify 'summary' in recipe at all, add it singly
                                    else:
                                        alertlist.append(new_errstring)
                                        break
                            # if badnews_type key is not present in recipe, assume its an error (not warning or Note)
                            else:
                                errorlist.append(new_errstring)
                                break
                        tmptextlist =[]
                    if "badnews" in recipe_item and recipe_item["badnews"] == 'one_allowed' and len(report_dict[recipe_item["dict_category_name"]]) > 1:
                        new_errstring = recipe_item["errstring"].format(count=len(recipe_item["dict_category_name"]))
                        if "badnews_type" in recipe_item and recipe_item["badnews_type"] == 'warning':
                            warninglist.append(new_errstring)
                        elif "badnews_type" in recipe_item and recipe_item["badnews_type"] == 'note' and new_errstring not in notelist:
                            # adding provision to conditional to prevent summary items from repeating
                            if new_errstring not in notelist:
                                notelist.append(new_errstring)
                        else:
                            errorlist.append(new_errstring)
                        tmptextlist =[]
            else:
                # if this is required content and is absent from report_dict, we have an error.
                if "required" in recipe_item and recipe_item["required"] == True:
                    errorlist.append(recipe_item["errstring"])
                # if this is suggested content and is absent from report_dict, we have a warning.
                if "suggested" in recipe_item and recipe_item["suggested"] == True:
                    warninglist.append(recipe_item["errstring"])
                # if we have alternate content, add it in (if we continue using this item)
                if "alternate_content" in recipe_item:
                    tmptextlist = []
                    tmptextlist.append("")
                    if "title" in recipe_item["alternate_content"]:
                        formattedtitle = "\n{:-^80}".format(" "+recipe_item["alternate_content"]["title"]+" ")
                        tmptextlist.append(formattedtitle)
                    # if no alt title listed, get original title
                    elif "title" in recipe_item and recipe_item["title"]:
                        formattedtitle = "\n{:-^80}".format(" "+recipe_item["title"]+" ")
                        tmptextlist.append(formattedtitle)
                    if "text" in recipe_item["alternate_content"]:
                        tmptextlist.append(recipe_item["alternate_content"]["text"])
                    elif "text" in recipe_item and recipe_item["text"]:
                        tmptextlist.append(recipe_item["text"])
                # if this is not required and there's no output from the file, reset this part of the report.
                elif "required" not in recipe_item or recipe_item["required"] != True:
                    tmptextlist =[]

        textreport_list += tmptextlist
    return validator_warnings

def addTopList(textreport_list, notice_list, headertext, notice_label, footer=False):
    if notice_list:
        newlist = []
        # add the header
        header = "\n{:-^80}".format(" %s " % headertext)
        newlist.append(header)
        # add notices with labels
        for notice_string in notice_list:
            new_errstring = "** {}: {}\n".format(notice_label, notice_string)
            newlist.append(new_errstring)
        # add footer where indicated
        if footer:
            footer = "\nIf you have any questions about how to handle these errors,\nplease contact %s.\n" % cfg.support_email_address
            newlist.append(footer)
        # add this to the top of the report
        textreport_list = newlist + textreport_list
    return textreport_list

def addErrorList(textreport_list, errorlist, warninglist):
    if errorlist or warninglist:
        tmperrorlist = []

        # define header & footer
        if errorlist:
            header = "\n{:-^80}".format(" ERRORS ")
        elif warninglist:
            header = "\n{:-^80}".format(" WARNINGS ")

        # add header to tmplist
        tmperrorlist.append(header)

        # add errors to tmplist
        for errstring in errorlist:
            new_errstring = "** ERROR: {}\n".format(errstring)
            tmperrorlist.append(new_errstring)
        for errstring in warninglist:
            new_errstring = "** WARNING: {}\n".format(errstring)
            tmperrorlist.append(new_errstring)

        # add footer to tmplist
        if errorlist:
            footer = "\nIf you have any questions about how to handle these errors,\nplease contact %s." % cfg.support_email_address
            tmperrorlist.append(footer)

        # add this to list for output
        textreport_list = tmperrorlist + textreport_list
    return textreport_list

def addBanner(textreport_list, errorlist, warninglist, validator_warnings, scriptname):
    if scriptname == "converter":
        banner = report_recipe.getBanners()['converter'].format(helpurl=cfg.helpurl)
    elif scriptname == "reporter" or scriptname == "rsuitevalidate":
        if errorlist: #or warninglist:
            banner = report_recipe.getBanners()['reporter_err']
        else:
            banner = report_recipe.getBanners()['reporter_noerr']
    elif scriptname == "validator":
        if validator_warnings:
            vwarningstring = '\n'.join(validator_warnings)
            banner = report_recipe.getBanners()['validator_err'].format(helpurl=cfg.helpurl, v_warning_banner=vwarningstring)
        else:
            banner = report_recipe.getBanners()['validator_noerr']
    textreport_list.insert(0,banner)

def generateReport(report_dict, outputtxt_path, scriptname):
    logger.info("* * * commencing buildreport function...")
    # # this reports the name of ths script that called this function, capturing so we can customize report output per product
    # invokedby_scriptpath = inspect.stack()[1][2]
    # invokedby_script = os.path.splitext(os.path.basename(invokedby_scriptpath))[0]

    # using this shortname:longname 'stylenamemap' to prevent repeat lookups in styles.xml
    stylenamemap = {}
    textreport = ""
    textreport_list = []
    errorlist = []
    warninglist = []
    notelist = []
    validator_warnings = []

    # get the report recipe
    recipe = report_recipe.getReportRecipe(cfg.titlestyle, cfg.authorstyle, cfg.isbnstyle, cfg.logostyle, cfg.booksection_stylename, cfg.notessection_stylename)

    # build our style report as a list of strings
    for item in sorted(recipe):
        validator_warnings = buildReport(report_dict, textreport_list, scriptname, stylenamemap, recipe[item], errorlist, warninglist, notelist, validator_warnings)

    # add Error List, errheader & footer, Warning List, Notice List& footer
    textreport_list = addTopList(textreport_list, notelist, "PROCESSING NOTES", "NOTE")
    textreport_list = addTopList(textreport_list, warninglist, "WARNINGS", "WARNING")
    textreport_list = addTopList(textreport_list, errorlist, "ERRORS", "ERROR", True)
    # textreport_list = addErrorList(textreport_list, errorlist, warninglist)

    # add success/fail banner based script & presence of alerts
    addBanner(textreport_list, errorlist, warninglist, validator_warnings, scriptname)

    # print report to console
    for line in textreport_list:    # for debug
        print line

    # create email_ready string
    email_text = "\n".join(textreport_list)

    # write report to file
    os_utils.writeListToFileByLine(textreport_list, outputtxt_path)

    logger.info("* * * ending buildReport function.")
    return email_text

#---------------------  MAIN
# # only run if this script is being invoked directly
# if __name__ == '__main__':
#
#     # set up debug log to console
#     logging.basicConfig(level=logging.DEBUG)
#
#     report_dict = {}
#     # for debug:
#     txtfile = '/Users/matthew.retzer/Documents/programming_projects/1710_1_makereport/etstB.txt'
#     report_dict = os_utils.readJSON('/Users/matthew.retzer/Documents/programming_projects/1708_2_python_ssconvertertests/tmpdir_validatepy/stylereport.json')
#
#     generateReport(report_dict, txtfile)
#
#     # concerns / next steps.  Will have to look at if we can scrape authors titles and isbns pre combination.
#     # should add illustration holder spans too
#     # character styles in use are not in order - check Erica's?. Changed them to separate category.
#     # could have also sorte so they were mixed in with orig category.
