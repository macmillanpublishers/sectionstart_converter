######### IMPORT SOME STANDARD PY LIBRARIES

import sys
import os
import shutil
import re
import json
import logging
import time
from lxml import etree


# initialize logger
logger = logging.getLogger(__name__)


######### METHODS
# trying to make this flexible so "dest" param can take a file or dir as an argument, and create dir if it does not exist.
def moveFile(pathtofile, dest):
    # Create dest directory if it does not exist (guessing whether we have a file or dir as dest, based on period in name)
    if "." not in dest:
        if not os.path.isdir(dest):
            os.makedirs(dest)
    else:
        if not os.path.isdir(os.path.dirname(dest)):
            os.makedirs(os.path.dirname(dest))
    # Move the file
    try:
        shutil.move(pathtofile, dest)
    except Exception, e:
        logger.error('Failed to move file, exiting', exc_info=True)
        sys.exit(1)

def copyFiletoFile(pathtofile, dest_file):
    if not os.path.isdir(os.path.dirname(dest_file)):
        os.makedirs(os.path.dirname(dest_file))
    try:
        shutil.copyfile(pathtofile, dest_file)
    except Exception, e:
        logger.error('Failed copyfile, exiting', exc_info=True)
        sys.exit(1)

def copyDir(pathtodir, dest_dir):
    if os.path.isdir(os.path.dirname(dest_dir)):
        dest_dir="%s_%s" % (dest_dir, time.strftime("%y%m%d-%H%M%S"))
    try:
        shutil.copytree(pathtodir, dest_dir)
    except Exception, e:
        logger.error('Failed copydir, exiting', exc_info=True)
        sys.exit(1)

def incrementToUniquePath(init_path):
    # check if init_path folder exists, if so, increment until we find one avail.
    new_path = init_path
    while os.path.exists(new_path):
        current_name = new_path.rpartition("_")[0]
        current_number = new_path.rpartition("_")[2]
        new_int = int(current_number) + 1
        new_path = "%s_%02d" % (current_name, new_int)
    return new_path

# def setupTmpfolder(tmpfolderpath):
#     if os.path.isdir(tmpfolderpath):
#         # tmpfolderpath = incrementToUniquePath("%s_01" % tmpfolderpath)
#         tmpfolderpath = "%s_%s" % (tmpfolderpath, time.strftime("%y%m%d-%H%M%S"))
#     # make new dir
#     if not os.path.isdir(tmpfolderpath):
#         os.makedirs(tmpfolderpath)
#     return tmpfolderpath

def setupOutfolder(outfolderpath):
    # make the dir if it does not  exist
    if not os.path.isdir(outfolderpath):
        os.makedirs(outfolderpath)
    # if dir is not empty (not including invisible files), move outfiles to 'previous_runs'
    elif [f for f in os.listdir(outfolderpath) if not f.startswith('.')]:
        prevrun_folder = os.path.join(outfolderpath, 'previous_runs', 'run_01')
        prevrun_folder = incrementToUniquePath(prevrun_folder)
        # make new dir
        if not os.path.isdir(prevrun_folder):
            os.makedirs(prevrun_folder)
        # move files over
        for f in os.listdir(outfolderpath):
            if not f == "previous_runs":
                moveFile(os.path.join(outfolderpath, f), prevrun_folder)

# if current os_object exists at path, rm or fail
def readJSON(filename):
    try:
        with open(filename) as json_data:
            d = json.load(json_data)
            logger.debug("reading in json file %s" % filename)
            return d
    except Exception, e:
        logger.error('Failed read JSON file, exiting', exc_info=True)
        sys.exit(1)

def rm_existing_os_object(path, obj_name):
    if os.path.exists(path):
        logger.debug("os_object '%s' exists, removing" % obj_name)
        try:
            if os.path.isdir(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
        except Exception, e:
            logger.error('Failed remove os_object, exiting', exc_info=True)
            sys.exit(1)

def writeXMLtoFile(root, filename):
    try:
        newfile = open(filename, 'w')
        with newfile as f:
            f.write(etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes"))
            f.close()
        logger.info("wrote xml to file '%s'" % filename)
    except Exception, e:
        logger.error('Failed write xml to file, exiting', exc_info=True)
        sys.exit(1)

def dumpJSON(dictname, filename):
    # # write json to console:
    # json.dump(dictname, sys.stdout, sort_keys=True, indent=4)        # debug
    # # and write json to file
    try:
        with open(filename, 'w') as outfile:
            json.dump(dictname, outfile, sort_keys=True, indent=4)
        logger.info("wrote dict to json file '%s'" % filename)
    except Exception, e:
        logger.error('Failed write JSON file, exiting', exc_info=True)
        sys.exit(1)

# expecting alert_type of "error", "warning", or "notice", but will accept anything.
def logAlerttoJSON(alerts_json, alert_category, new_errtext):
    if os.path.exists(alerts_json):
        alerts_dict = readJSON(alerts_json)
    else:
        alerts_dict = {}
    if alert_category in alerts_dict:
        alerts_dict[alert_category].append(new_errtext)
    else:
        alerts_dict[alert_category] = []
        alerts_dict[alert_category].append(new_errtext)
    dumpJSON(alerts_dict, alerts_json)

# https://stackoverflow.com/questions/10821083/writing-nicely-formatted-text-in-python - we'll see if we need to write line by line
def writeListToFileByLine(text, filename):
    try:
        newfile = open(filename, 'w')
        with newfile as f:
            # for line in text:
            #     f.write("%s\n" % line)
            for item in text:
                print>>f, item
                # f.write(line)
            f.close()
        logger.info("wrote text to file '%s'" % filename)
    except Exception, e:
        logger.error('Failed write text to file, exiting', exc_info=True)
        sys.exit(1)

def writeAlertstoTxtfile(alerts_json, this_outfolder):
    alerts_dict = readJSON(alerts_json)
    alerttxt_list = []
    if os.path.exists(alerts_json):
        # get all the alert text in a list
        for alert_category, alerts in sorted(alerts_dict.iteritems()):
            alerttxt_list.append("{}(s):".format(alert_category.upper()))
            for alert in alerts:
                alerttxt_list.append("- {}".format(alert))
            alerttxt_list.append("")
        # if we found any alerts at all
        if alerttxt_list:
            # figure out appropriate filename
            if "error" in alerts_dict:
                alertfile = os.path.join(this_outfolder, "ERROR.txt")
            elif "warning" in alerts_dict:
                alertfile = os.path.join(this_outfolder, "WARNING.txt")
            else:
                alertfile = os.path.join(this_outfolder, "NOTICE.txt")
            # write our file
            writeListToFileByLine(alerttxt_list, alertfile)
    return alerttxt_list
