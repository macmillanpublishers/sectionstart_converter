import sys
import os
import zipfile
import shutil
import re
import getpass
import platform
import json
import logging.config
import time


# #	# # # # # # ARGS
### Arg 1 - Filename
script_name = os.path.basename(sys.argv[0]).replace("_main.py","")
inputfile = sys.argv[1]

# strip out surrounding double quotes if passed from batch file.
if inputfile[0] == '"':
    inputfile = inputfile[1:]
if inputfile[-1:] == '"':
    inputfile = inputfile[:-1]
# get just basename
original_inputfilename = os.path.basename(inputfile)
# separate filename and extension
original_inputfilename_noext, inputfile_ext = os.path.splitext(original_inputfilename)
# clean out non-alphanumeric chars
inputfilename_noext = re.sub('\W','',original_inputfilename_noext)
inputfilename = inputfilename_noext + inputfile_ext

### Arg 2 - processwatch file for standalones, or alternate logfile if validator (embedded run)
#   Could replace existing logging or could try to add a handler on the fly to log to both places
if sys.argv[2:]:
    if script_name == "validator":
        validator_logfile = sys.argv[2]
    else:
        processwatch_file = sys.argv[2]
else:
    validator_logfile = ''
    processwatch_file = ''


# # # # # # # # ENV
loglevel = "INFO"		# global setting for logging. Options: DEBUG, INFO, WARN, ERROR, CRITICAL.  See defineLogger() below for more info
hostOS = platform.system()
currentuser = getpass.getuser()
# the path of this file: setting '__location__' allows this relative path to adhere to this file, even when invoked from a different path:
# 	https://stackoverflow.com/questions/4060221/how-to-reliably-open-a-file-in-the-same-directory-as-a-python-script
__location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))


# # # # # # # # CONFIGURE BASED ON ENVIRONMENT:
# The parent tmpdir needs to be manually set.  Everything else here is dynamically spun off based on
# 	which script is invoked and the location of the dropbox folder
main_tmpdir = os.path.join("S:",os.sep,"pythonxml_tmp") # debug-change for prod
scripts_dir = ""	# we use realtive paths based on __location__ of this file (cfg.py) unless scripts_dir has a value

# # # # # # # # PATHS
### Top level Folders
# dropbox folder (for in and out folders)
if hostOS == "Windows":
    dropboxfolder = os.path.join("C:",os.sep,"Users",currentuser,"Dropbox (Macmillan Publishers)")
else:
    dropboxfolder = os.path.join(os.sep,"Users",currentuser,"Dropbox (Macmillan Publishers)")
    main_tmpdir = os.path.join(os.sep,"Users",currentuser,"Documents","programming_projects","tmpdir") # debug, for testing on MacOS
# tmpfolder and outfolder
if script_name == "validator":
    tmpdir = os.path.dirname(inputfile)
    this_outfolder = tmpdir
else:
    tmpdir = os.path.join(main_tmpdir,"%s_%s" % (inputfilename_noext, time.strftime("%y%m%d-%H%M%S")))
    # in_folder = os.path.join(dropboxfolder, "stylecheck", script_name, "IN")
    out_folder = os.path.join(dropboxfolder, "stylecheck", script_name, "OUT")
    this_outfolder = os.path.join(out_folder, inputfilename_noext)
# log folder
logdir = os.path.join(dropboxfolder, "bookmaker_logs", "stylecheck")

### Files
newdocxfile = os.path.join(this_outfolder,"{}_converted.docx".format(inputfilename_noext))  	# the rebuilt docx post-converter or validator
stylereport_txt = os.path.join(this_outfolder,"{}_StyleReport.txt".format(inputfilename_noext))
workingfile = os.path.join(tmpdir, inputfilename)
ziproot = os.path.join(tmpdir, "{}_unzipped".format(inputfilename_noext))		# the location where we unzip the input file
template_ziproot = os.path.join(tmpdir, "macmillan_template_unzipped")
stylereport_json = os.path.join(tmpdir, "stylereport.json")
alerts_json = os.path.join(tmpdir, "alerts.json")

### Resources in other Repos
macmillan_template_name = "macmillan.dotx"
if scripts_dir:
    macmillan_template = os.path.join(scripts_dir, "Word-template_assets","StyleTemplate_auto-generate","macmillan.dotx")
    macmillanstyles_json = os.path.join(scripts_dir, "Word-template_assets","StyleTemplate_auto-generate","macmillan.json")
    vbastyleconfig_json = os.path.join(scripts_dir, "Word-template_assets","StyleTemplate_auto-generate","vba_style_config.json")
    section_start_rules_json = os.path.join(scripts_dir, "bookmaker_validator","section_start_rules.json")
    styleconfig_json = os.path.join(scripts_dir, "htmlmaker_js","style_config.json")
    smtp_txt = os.path.join(scripts_dir, "bookmaker_authkeys","smtp.txt")
    db_access_token_txt = os.path.join(scripts_dir, "bookmaker_authkeys","access_token.txt")
else:
    macmillan_template = os.path.join(__location__,'..','..',"Word-template_assets","StyleTemplate_auto-generate","macmillan.dotx")
    macmillanstyles_json = os.path.join(__location__,'..','..',"Word-template_assets","StyleTemplate_auto-generate","macmillan.json")
    vbastyleconfig_json = os.path.join(__location__,'..','..',"Word-template_assets","StyleTemplate_auto-generate","vba_style_config.json")
    section_start_rules_json = os.path.join(__location__,'..','..',"bookmaker_validator","section_start_rules.json")
    styleconfig_json = os.path.join(__location__,'..','..',"htmlmaker_js","style_config.json")
    smtp_txt = os.path.join(__location__,'..','..',"bookmaker_authkeys","smtp.txt")
    db_access_token_txt = os.path.join(__location__,'..','..',"bookmaker_authkeys","access_token.txt")

# # # # # # # # RELATIVE PATHS for unzipping and zipping docx files
### xml filepaths relative to ziproot
docxml_relpath = os.path.join("word","document.xml")
stylesxml_relpath = os.path.join("word","styles.xml")
settingsxml_relpath = os.path.join("word","settings.xml")  # for rsid index
custompropsxml_relpath = os.path.join("docProps","custom.xml")  # for version document property
numberingxml_relpath = os.path.join("word","numbering.xml")  # for replacing or preserving wholesale
rels_relpath = os.path.join("_rels",".rels")
contenttypes_relpath = os.path.join(".","[Content_Types].xml")

# Template dirs & files
template_customprops_xml = os.path.join(template_ziproot, custompropsxml_relpath)
template_styles_xml = os.path.join(template_ziproot, stylesxml_relpath)
template_numbering_xml = os.path.join(template_ziproot, numberingxml_relpath)
template_rels_file = os.path.join(template_ziproot, rels_relpath)
template_contenttypes_xml = os.path.join(template_ziproot, contenttypes_relpath)

# doc files
doc_xml = os.path.join(ziproot, docxml_relpath)
numbering_xml = os.path.join(ziproot, numberingxml_relpath)
styles_xml = os.path.join(ziproot, stylesxml_relpath)
settings_xml = os.path.join(ziproot, settingsxml_relpath)
customprops_xml = os.path.join(ziproot, custompropsxml_relpath)
rels_file = os.path.join(ziproot, rels_relpath)
contenttypes_xml = os.path.join(ziproot, contenttypes_relpath)


# # # # # # # GLOBAL VARS
# alert email address:
alert_email_address = "workflows@macmillan.com"
support_email_address = "workflows@macmillan.com"
from_email_address = "workflows@macmillan.com"
# The first document version in history with section starts
sectionstart_versionstring = '4.7.0'
# TitlepageTitle style
titlestyle = "Titlepage Book Title (tit)"
chapnumstyle = "Chap Number (cn)"
chaptitlestyle = "Chap Title (ct)"
partnumstyle = "Part Number (pn)"
parttitlestyle = "Part Title (pt)"
isbnstyle = "span ISBN (isbn)"
authorstyle = "Titlepage Author Name (au)"
illustrationholder_style = "Illustration holder (ill)"
titlesection_stylename = "Section-Titlepagesti"
copyrightsection_stylename = "Section-Copyrightscr"

autonumber_sections = {"Section-Chapter (scp)":"arabic", "Section-Part (spt)":"roman", "Section-Appendix (sap)":"alpha"}
# the first 3 are shapes, the 4th is a section break
objects_to_delete = ["mc:AlternateContent", "w:pict", "w:drawing", "w:sectPr"]
nocharstyle_headingstyles = ["FMHeadfmh", "BMHeadbmh", "ChapNumbercn", "PartNumberpn"]
nonprintingheads = ["ChapTitleNonprintingctnp", "BMHeadNonprintingbmhnp", "FMHeadNonprintingfmhnp"]
copyrightstyles = ["CopyrightTextdoublespacecrtxd", "CopyrightTextsinglespacecrtx"]

# Word namespace vars
wnamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
w14namespace = 'http://schemas.microsoft.com/office/word/2010/wordml'
vtnamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
mcnamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlnamespace = "http://www.w3.org/XML/1998/namespace"
wordnamespaces = {'w': wnamespace, 'w14': w14namespace, 'vt': vtnamespace, 'mc': mcnamespace}


# # # # # # # LOGGING SETUP via dictConfig
# loglevel can be globally set at the top of this script.
# FYI the loglevel for a handler will be whichever setting is more restrictive: handler setting or logger setting
# So to log DEBUG in console and INFO for file, set logger & stream handler to "DEBUG" and file handler to "INFO"
def defineLogger(logfile, loglevel):
    logging.config.dictConfig({
        'version': 1,
        'disable_existing_loggers': False,
        'formatters': {
            'console': {
                'format': '[%(levelname)s] %(name)s.%(funcName)s : %(message)s'
            },
            'file': {
                'format': '%(asctime)s [%(levelname)s] %(name)s.%(funcName)s : %(message)s',
                'datefmt': '%y-%m-%d %H:%M:%S'
            # },
            # 'warnings': {
            # 	'format': '%(message)s'
            }
        },
        'handlers': {
            'stream': {
                'class': 'logging.StreamHandler',
                'formatter': 'console',
                # 'level': 'DEBUG'
            },
            'file':{
                'class': 'logging.FileHandler',
                'formatter': 'file',
                'filename': logfile
                # 'level': 'DEBUG'
            # },
            # 'secondfile':{
            # 	'class':'logging.FileHandler',
            # 	'formatter':'warnings',
            # 	'filename' : warnings_json
            }
        },
        'loggers': {
            '': {
                'handlers': ['stream', 'file'],
                'level': loglevel,
                'propagate': True
            # },
            # 'w_logger':{
            #      'handlers': ['stream', 'secondfile'],
            #     'level': loglevel,
            #     'propagate': True
            }
        }
    })


# class StructuredMessage(object):
#     def __init__(self, message, **kwargs):
#         self.message = message
#         self.kwargs = kwargs

#     def __str__(self):
#         return '%s >>> %s' % (self.message, json.dumps(self.kwargs))

# _ = StructuredMessage   # optional, to improve readability


# TODO:
# 2) the validator
# for stylereporter: got a unicode encoding for special char in authorname
# commenting out pagebreak logging to json, reinstate depending on macro / pagenumber count speed is good (also custom styles revert)
