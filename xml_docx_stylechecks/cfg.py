import sys
import os
import zipfile
import shutil
import re
import getpass
import platform
import json
import logging.config

# import shared_utils.os_utils as os_utils

# NOTE: The tmpdir needs to be manually set.  Everything else here is dynamically spun off based on
# 	which script is invoked and the location of the dropbox folder
# 	I guess we could dynamically set tmpdir it if we used standard sys tmp locations that are garbage collected

# #	# # # # # # ARGS
script_name = os.path.basename(sys.argv[0]).replace("_main.py","")
inputfile = sys.argv[1]
inputfilename = os.path.basename(inputfile)
inputfilename_noext = os.path.splitext(inputfilename)[0]
# so we can log to the validator logfile if we need to. Could replace or could try to add a handler on the fly
# to log to both places
if script_name == "validator" and sys.argv[2:]:
	validator_logfile = sys.argv[2]
else:
	validator_logfile = ''

# # # # # # # # ENV
loglevel = "INFO"		# global setting for logging. Options: DEBUG, INFO, WARN, ERROR, CRITICAL.  See defineLogger() below for more info
hostOS = platform.system()
currentuser = getpass.getuser()
scripts_dir = ""	# we use realtive paths based on location of this file (cfg.py) unless scripts_dir has a value
# the path of this file: setting '__location__' allows this relative path to adhere to this file, even when invoked from a different path:
# 	https://stackoverflow.com/questions/4060221/how-to-reliably-open-a-file-in-the-same-directory-as-a-python-script
__location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))


# # # # # # # # PATHS
### Folders
if script_name == "validator":
	tmpdir = os.path.dirname(inputfile)
else:
	tmpdir = os.path.join(os.sep,"Users",currentuser,"Documents","programming_projects","1708_2_python_ssconvertertests","tmpdir",inputfilename_noext)
if hostOS == "Windows":
	dropboxfolder = os.path.join("C:",os.sep,"Users",currentuser,"Dropbox (Macmillan Publishers)")
else:
	dropboxfolder = os.path.join(os.sep,"Users",currentuser,"Dropbox (Macmillan Publishers)")
logdir = os.path.join(dropboxfolder, "bookmaker_logs", "stylecheck")
# in_folder = os.path.join(dropboxfolder, "stylecheck", script_name, "IN")  # not necessary?
out_folder = os.path.join(dropboxfolder, "stylecheck", script_name, "OUT")
this_outfolder = os.path.join(out_folder, inputfilename_noext)

### Files
workingfile = os.path.join(tmpdir, inputfilename)
# originalfile_copy = os.path.join(tmp_original_dir, inputfilename)
ziproot = os.path.join(tmpdir, "{}_unzipped".format(inputfilename_noext))		# the location where we unzip the input file
template_ziproot = os.path.join(tmpdir, "macmillan_template_unzipped")
newdocxfile = os.path.join(this_outfolder,"{}_converted.docx".format(inputfilename_noext))  	# the rebuilt docx post-converter or validator
stylereport_json = os.path.join(tmpdir, "stylereport.json")
warnings_json = os.path.join(tmpdir, "warnings.json")

### Resources in other Repos
macmillan_template_name = "macmillan.dotx"
if scripts_dir:
	macmillan_template = os.path.join(scripts_dir, "Word-template_assets","StyleTemplate_auto-generate","macmillan.dotx")
	macmillanstyles_json = os.path.join(scripts_dir, "Word-template_assets","StyleTemplate_auto-generate","macmillan.json")
	vbastyleconfig_json = os.path.join(scripts_dir, "Word-template_assets","StyleTemplate_auto-generate","vba_style_config.json")
	section_start_rules_json = os.path.join(scripts_dir, "bookmaker_validator","section_start_rules.json")
	styleconfig_json = os.path.join(scripts_dir, "htmlmaker_js","style_config.json")
else:
	macmillan_template = os.path.join(__location__,'..','..',"Word-template_assets","StyleTemplate_auto-generate","macmillan.dotx")
	macmillanstyles_json = os.path.join(__location__,'..','..',"Word-template_assets","StyleTemplate_auto-generate","macmillan.json")
	vbastyleconfig_json = os.path.join(__location__,'..','..',"Word-template_assets","StyleTemplate_auto-generate","vba_style_config.json")
	section_start_rules_json = os.path.join(__location__,'..','..',"bookmaker_validator","section_start_rules.json")
	styleconfig_json = os.path.join(__location__,'..','..',"htmlmaker_js","style_config.json")


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
# The first document version in history with section starts
sectionstart_versionstring = '4.7.0'
# TitlepageTitle style
titlestylename = "Titlepage Book Title (tit)"
chapnumstyle = "Chap Number (cn)"
chaptitlestyle = "Chap Title (ct)"
partnumstyle = "Part Number (pn)"
parttitlestyle = "Part Title (pt)"
autonumber_sections = {"Section-Chapter (scp)":"arabic", "Section-Part (spt)":"roman", "Section-Appendix (sap)":"alpha"}

# Word namespace vars
wnamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
w14namespace = 'http://schemas.microsoft.com/office/word/2010/wordml'
vtnamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
wordnamespaces = {'w': wnamespace, 'w14': w14namespace, 'vt': vtnamespace}


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
# standardize headings and sections
# 1) look at flow control for crashes, figure out how we finish the process and return what we need to
	# (nested tries?)  Try inserting junk at top level locations too.
	# for now surface errors in text files, we'll add that for emails later
# 2) the validator
# possibly consolidate setup and cleanup scripts in lib as bundled calls
# for stylereporter: got a unicode encoding for special char in authorname
# commenting out pagebreak logging to json, reinstate depending on macro / pagenumber count speed is good (also custom styles revert)
# python logger is not the easiest way to create a json of only warnings; should probably just setup a separate function for that.
