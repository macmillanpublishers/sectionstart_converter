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
# get project_dir (eg, 'stylecheck/converter', 'stylecheck_stg/reporter')
project_dir = os.path.dirname(os.path.dirname(inputfile))
# separate filename and extension
original_inputfilename_noext, inputfile_ext = os.path.splitext(original_inputfilename)
# clean out non-alphanumeric chars
inputfilename_noext = re.sub('\W','',original_inputfilename_noext)
# cut down extraordinarily long filenames to 47 char, + timestamp
if len(inputfilename_noext) > 60:
    inputfilename_noext = "%s_%s" % (str[:47],time.strftime("%y%m%d%H%M%S"))
inputfilename = inputfilename_noext + inputfile_ext

### Arg 2 - processwatch file for standalones, or alternate logfile if validator (embedded run)
#   Could replace existing logging or could try to add a handler on the fly to log to both places
validator_logfile = ''
processwatch_file = ''
if sys.argv[2:]:
    if script_name.startswith("validator"):
    # if script_name == "validator":
        validator_logfile = sys.argv[2]
    else:
        processwatch_file = sys.argv[2]


# # # # # # # # ENV
loglevel = "INFO"		# global setting for logging. Options: DEBUG, INFO, WARN, ERROR, CRITICAL.  See defineLogger() below for more info
hostOS = platform.system()
currentuser = getpass.getuser()
# the path of this file: setting '__location__' allows this relative path to adhere to this file, even when invoked from a different path:
# 	https://stackoverflow.com/questions/4060221/how-to-reliably-open-a-file-in-the-same-directory-as-a-python-script
__location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
staging_file = os.path.join("C:",os.sep,"staging.txt")


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
    main_tmpdir = os.path.join(os.sep,"Users",currentuser,"stylecheck_tmp") # debug, for testing on MacOS
    staging_file = os.path.join(os.sep,"Users",currentuser,"staging.txt")
# tmpfolder and outfolder
if script_name.startswith("validator"):
# if script_name == "validator":
    tmpdir = os.path.dirname(inputfile)
    this_outfolder = tmpdir
else:
    tmpdir = os.path.join(main_tmpdir,"%s_%s_%s" % (script_name[0], inputfilename_noext, time.strftime("%y%m%d-%H%M%S")))
    # tmpdir = os.path.join(main_tmpdir,"%s_%s" % (inputfilename_noext, 'debug'))     # for debug
    # in_folder = os.path.join(project_dir, 'IN')
    out_folder = os.path.join(project_dir, 'OUT')
    this_outfolder = os.path.join(out_folder, inputfilename_noext)
# log folder (differs on staging server)
if os.path.exists(staging_file):
    logdir = os.path.join(dropboxfolder, "bookmaker_logs", "stylecheck_stg")
else:
    logdir = os.path.join(dropboxfolder, "bookmaker_logs", "stylecheck")


### Files
newdocxfile = os.path.join(this_outfolder,"{}_converted.docx".format(inputfilename_noext))  	# the rebuilt docx post-converter or validator
stylereport_txt = os.path.join(this_outfolder,"{}_StyleReport.txt".format(inputfilename_noext))
# if script_name == "validator":
if script_name.startswith("validator"):
    stylereport_txt = os.path.join(this_outfolder,"{}_ValidationReport.txt".format(inputfilename_noext))
workingfile = os.path.join(tmpdir, inputfilename)
ziproot = os.path.join(tmpdir, "{}_unzipped".format(inputfilename_noext))		# the location where we unzip the input file
template_ziproot = os.path.join(tmpdir, "macmillan_template_unzipped")
stylereport_json = os.path.join(tmpdir, "stylereport.json")
alerts_json = os.path.join(tmpdir, "alerts.json")
isbn_check_json = os.path.join(tmpdir, "isbn_check.json")

### Resources in other Repos
# use static paths if scripts_dir exists, else try to use relative paths
if scripts_dir:
    scripts_dir_path = scripts_dir
else:
    scripts_dir_path = os.path.join(__location__,'..','..')

# rsuite versus macmillan template paths. For now mocking up a separate repo locally for rsuite
if script_name.startswith("rsuite"):
    templatefiles_path = os.path.join(scripts_dir_path,"RSuite_Word-template","StyleTemplate_auto-generate")
    template_name = "Rsuite"
else:
    templatefiles_path = os.path.join(scripts_dir_path,"Word-template_assets","StyleTemplate_auto-generate")
    template_name = "macmillan"

# paths
section_start_rules_json = os.path.join(scripts_dir_path, "bookmaker_validator","section_start_rules.json")
smtp_txt = os.path.join(scripts_dir_path, "bookmaker_authkeys","smtp.txt")
db_access_token_txt = os.path.join(scripts_dir_path, "bookmaker_authkeys","access_token.txt")
macmillan_template = os.path.join(templatefiles_path, "%s.dotx" % template_name)
macmillanstyles_json = os.path.join(templatefiles_path, "%s.json" % template_name)
vbastyleconfig_json = os.path.join(templatefiles_path, "vba_style_config.json")
styleconfig_json = os.path.join(templatefiles_path, "style_config.json")


# # # # # # # # RELATIVE PATHS for unzipping and zipping docx files
### xml filepaths relative to ziproot
docxml_relpath = os.path.join("word","document.xml")
stylesxml_relpath = os.path.join("word","styles.xml")
settingsxml_relpath = os.path.join("word","settings.xml")  # for rsid index
custompropsxml_relpath = os.path.join("docProps","custom.xml")  # for version document property
numberingxml_relpath = os.path.join("word","numbering.xml")  # for replacing or preserving wholesale
rels_relpath = os.path.join("_rels",".rels")
contenttypes_relpath = os.path.join(".","[Content_Types].xml")
endnotesxml_relpath = os.path.join("word","endnotes.xml")
footnotesxml_relpath = os.path.join("word","footnotes.xml")


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
endnotes_xml = os.path.join(ziproot, endnotesxml_relpath)
footnotes_xml = os.path.join(ziproot, footnotesxml_relpath)
commentsIds_xml = os.path.join(ziproot, "word", "commentsIds.xml")
comments_xml = os.path.join(ziproot, "word", "comments.xml")
commentsExtended_xml = os.path.join(ziproot, "word", "commentsExtended.xml")


# # # # # # # GLOBAL VARS
# alert email address:
alert_email_address = "Publishing Workflows <workflows@macmillan.com>"
support_email_address = "workflows@macmillan.com" # if the display name is present it comes out weird in user-messaging.. and not required for emails via smtplib
from_email_address = "Publishing Workflows <workflows@macmillan.com>"
always_bcc_address = "Workflows Notifications <wfnotifications@macmillan.com>"
helpurl = "https://confluence.macmillan.com/x/U4AYB#Stylecheck-ConverterandStylecheck-Reporter-ReviewingyourStylecheckReport"
# The first key version of a template-type
if script_name.startswith("rsuite"):
    templateversion_cutoff = '6.0.0'
else:
    templateversion_cutoff = '4.7.0'
# regex for finding ISBNS
isbnregex = re.compile(r"(97[89](\D?\d){10})")
isbnspanregex = re.compile(r"(^.*?)(97[89](\D?\d){10})(.*?$)")
# Hardcoded stylenames
if script_name.startswith("rsuite"):
    # RSuite hardcoded stylenames (can I get these from styleconfig? in some cases)
    titlesection_stylename = "Section-TitlepageSTI"
    booksection_stylename = "Section-Book (BOOK)"
    copyrightsection_stylename = "Section-CopyrightSCR"
    titlestyle = "Title (Ttl)"
    isbnstyle = "cs-isbn (isbn)"
    authorstyle = "Author1 (Au1)"
    logostyle = "Logo-Placement (Logo)"
    imageholder_style = "Image-Placement (Img)"
    inline_imageholder_style = "cs-image-placement (cimg)"
    footnotestyle = "FootnoteText" #/ "Footnote Text"
    endnotestyle = "EndnoteText" #/ "Endnote Text"
    spacebreakstyles = ['SeparatorSep','Blank-Space-BreakBsbrk','Ornamental-Space-BreakOsbrk']
    # for some reason the long-stylenames for these references are lowercase?
    valid_native_word_styles = ['Hyperlink', 'footnote reference', 'endnote reference', 'annotation reference']
else:
    titlestyle = "Titlepage Book Title (tit)"
    chapnumstyle = "Chap Number (cn)"
    chaptitlestyle = "Chap Title (ct)"
    partnumstyle = "Part Number (pn)"
    parttitlestyle = "Part Title (pt)"
    isbnstyle = "span ISBN (isbn)"
    authorstyle = "Titlepage Author Name (au)"
    logostyle = "Titlepage Logo (logo)"
    hyperlinkstyle = "span hyperlink (url)"
    imageholder_style = "Illustration holder (ill)"
    inline_imageholder_style = "span illustration holder (illi)"
    titlesection_stylename = "Section-Titlepagesti"
    copyrightsection_stylename = "Section-Copyrightscr"
    # staticstyle groups (section start)
    valid_native_word_styles = ['endnote reference', 'annotation reference']
    nocharstyle_headingstyles = ["FMHeadfmh", "BMHeadbmh", "ChapNumbercn", "PartNumberpn"]
    nonprintingheads = ["ChapTitleNonprintingctnp", "BMHeadNonprintingbmhnp", "FMHeadNonprintingfmhnp"]
    copyrightstyles = ["CopyrightTextdoublespacecrtxd", "CopyrightTextsinglespacecrtx"]
    autonumber_sections = {"Section-Chapter (scp)":"arabic", "Section-Part (spt)":"roman", "Section-Appendix (sap)":"alpha"}

# objects for deletion
shape_objects = ["mc:AlternateContent", "w:drawing", "w:pict"]
section_break = ["w:sectPr"]
bookmark_objects = ["w:bookmarkStart", "w:bookmarkEnd"]
comment_objects = ["w:commentRangeStart","w:commentRangeEnd","w:commentReference","w:comment","w15:commentEx", "w16cid:commentId"]

# Word namespace vars
wnamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
w14namespace = 'http://schemas.microsoft.com/office/word/2010/wordml'
w15namespace = 'http://schemas.microsoft.com/office/word/2012/wordml'
w16cidnamespace = 'http://schemas.microsoft.com/office/word/2016/wordml/cid'
vtnamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
mcnamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlnamespace = "http://www.w3.org/XML/1998/namespace"
wordnamespaces = {'w': wnamespace, 'w14': w14namespace, 'vt': vtnamespace, 'mc': mcnamespace, 'w15': w15namespace, "w16cid": w16cidnamespace}

# track changes elements:
collapse_trackchange_tags = ["ins", "moveTo"]
del_trackchange_tags = ["cellDel",
"cellIns",
"cellMerge",
"customXmlDelRangeEnd",
"customXmlDelRangeStart",
"customXmlInsRangeEnd",
"customXmlInsRangeStart",
"del",
"delInstrText",
"delText",
"moveFromRangeEnd",
"moveFromRangeStart",
"moveFrom",
"moveToRangeStart",
"moveToRangeEnd",
"numberingChange",
"pPrChange",
"rPrChange",
"sectPrChange",
"tblGridChange",
"tblPrChange",
"tblPrExChange",
"tcPrChange",
"trPrChange"]

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
