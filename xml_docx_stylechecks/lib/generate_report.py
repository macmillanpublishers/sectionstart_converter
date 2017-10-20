######### IMPORT PY LIBRARIES

import os
# import shutil
# import re
# import uuid
# import json
import sys
import logging
import inspect
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
    # lxml_utils = imp.load_source('lxml_utils', lxmlutilspath)
else:
    import cfg
    import shared_utils.os_utils as os_utils
    # import shared_utils.lxml_utils as lxml_utils
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
# lookup longname of style in styles.xml of file.
#  save looked up values in a dict to speed up repeat lookups
def getStyleLongname(styleshortname, stylenamemap):
    styles_tree = etree.parse(styles_xml)
    styles_root = styles_tree.getroot()
    if styleshortname in stylenamemap:
        stylelongname = stylenamemap[styleshortname]
        # print "in the map!"
    else:
        # print "not in tht emap!"
        searchstring = ".//w:style[@w:styleId='%s']/w:name" % styleshortname
        stylematch = styles_root.find(searchstring, wordnamespaces)
        # get fullname value and test against Macmillan style list
        stylelongname = stylematch.get('{%s}val' % wnamespace)
        stylenamemap[styleshortname] = stylelongname
    return stylelongname

def buildReport(report_dict, textreport_list, invokedby_script, stylenamemap, recipe_item, errorlist):
    # make sure we're not supposed to skip this recipe_item as per "exclude_from" key
    if not("exclude_from" in recipe_item and invokedby_script in recipe_item["exclude_from"]):
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
                for item in report_dict[recipe_item["dict_category_name"]]:
                    # line = "{:<40}:{:>30}".format("parent_section_start_type","'parent_section_start_content'")
                    newline = recipe_item["line_template"].format(description=item['description'], para_string='"'+item['para_string'].encode('utf-8')+'"', parent_section_start_content='"'+item['parent_section_start_content'].encode('utf-8')+'"', parent_section_start_type=getStyleLongname(item['parent_section_start_type'], stylenamemap), para_index=item['para_index'])
                    # newline = line.replace("zdescription", item['description']).replace("para_string", item['para_string']).replace("parent_section_start_content", item['parent_section_start_content'])
                    # newline = newline.replace("parent_section_start_type", getStyleLongname(item['parent_section_start_type'], stylenamemap))
                    tmptextlist.append(newline)
                    if "badnews" in recipe_item and recipe_item["badnews"] == True:
                        new_errstring = recipe_item["errstring"].format(description=item['description'], para_string='"'+item['para_string'].encode('utf-8')+'"', parent_section_start_content='"'+item['parent_section_start_content'].encode('utf-8')+'"', parent_section_start_type=getStyleLongname(item['parent_section_start_type'], stylenamemap),  para_index=item['para_index'])
                        # new_errstring = "** ERROR: {}\n".format(new_errstring)
                        errorlist.append(new_errstring)
                        tmptextlist =[]
            else:
                # if this is required content and is absent from report_dict, we have an error.
                if "required" in recipe_item and recipe_item["required"] == True:
                    errorlist.append(recipe_item["errstring"])
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
    # return errorlist

def addErrorList(textreport_list, errorlist):
    if errorlist:
        tmperrorlist = []
        # add header to tmplist
        errheader = "\n{:-^80}".format(" ERRORS ")
        tmperrorlist.append(errheader)
        # add errors to tmplist
        for errstring in errorlist:
            new_errstring = "** ERROR: {}\n".format(errstring)
            tmperrorlist.append(new_errstring)

        # add footer with our contact info to tmplist
        errfooter = "\nIf you have any questions about how to handle these errors,\nplease contact workflows@macmillan.com."
        tmperrorlist.append(errfooter)

        # add this to list for output
        textreport_list = tmperrorlist + textreport_list
    return textreport_list

def addReportSuccessFailBanner(textreport_list, errorlist):
    # could set this  width value globallty.
    banner = "{:^80}\n{:^80}"
    if errorlist:
        banner = banner.format("OOPS!", "Problems were found with the styles in your document.")
        banner = banner + "\n"
    else:
        banner = banner.format("CONGRATULATIONS! YOU PASSED!", "But you're not done yet. Please check the info listed below.")
    textreport_list.insert(0,banner)

def generateReport(report_dict, outputtxt_path):
    logger.info("* * * commencing buildreport function...")
    # this reports the name of ths script that called this function, capturing so we can customize report output per product
    invokedby_scriptpath = inspect.stack()[1][1]
    invokedby_script = os.path.splitext(os.path.basename(invokedby_scriptpath))[0]

    # using this shortname:longname 'stylenamemap' to prevent repeat lookups in styles.xml
    stylenamemap = {}
    textreport = ""
    textreport_list = []
    errorlist = []

    # get the report recipe
    recipe = report_recipe.getReportRecipe(cfg.titlestyle, cfg.authorstyle, cfg.isbnstyle)

    # build our style report as a list of strings
    for item in sorted(recipe):
        buildReport(report_dict, textreport_list, invokedby_script, stylenamemap, recipe[item], errorlist)

    # add Error List, errheader & footer to
    textreport_list = addErrorList(textreport_list, errorlist)

    # add success/fail banner based on whether errorlist is empty
    addReportSuccessFailBanner(textreport_list, errorlist)

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
