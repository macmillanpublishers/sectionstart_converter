######### IMPORT PY LIBRARIES
import os
import shutil
import re
import uuid
import json
import sys
import collections
import logging
import dropbox
import textwrap


######### IMPORT LOCAL MODULES
if __name__ == '__main__':
    # to go up a level to read cfg (and other files) when invoking from this script (for testing).
    cfgpath = os.path.join(sys.path[0], '..', 'cfg.py')
    osutilspath = os.path.join(sys.path[0], '..', 'shared_utils', 'os_utils.py')
    unzipDOCXpath = os.path.join(sys.path[0], '..', 'shared_utils', 'unzipDOCX.py')
    sendmailpath = os.path.join(sys.path[0], '..', 'shared_utils', 'sendmail.py')
    import imp
    cfg = imp.load_source('cfg', cfgpath)
    import generate_report  # this is in the same dir so needs no direction for relative import
    import usertext_templates
    os_utils = imp.load_source('os_utils', osutilspath)
    unzipDOCX = imp.load_source('unzipDOCX', unzipDOCXpath)
    sendmail = imp.load_source('sendmail', sendmailpath)
else:
    import cfg
    import lib.generate_report as generate_report
    import lib.usertext_templates as usertext_templates
    import shared_utils.os_utils as os_utils
    import shared_utils.unzipDOCX as unzipDOCX
    import shared_utils.sendmail as sendmail


######### LOCAL DECLARATIONS
# initialize logger
logger = logging.getLogger(__name__)
# get db access token
with open(cfg.db_access_token_txt) as f:
    db_access_token = f.readline()
dropboxfolder = cfg.dropboxfolder


#---------------------  METHODS
def getSubmitterViaAPI(inputfile):
    logger.info("Retrieve submitter info via Dropbox api...")
    submitter_email = ""
    display_name = ""
    try:
        dropbox_relpath = inputfile.replace(dropboxfolder,"")
        dbx = dropbox.Dropbox(db_access_token)
        submitter = (dbx.files_get_metadata(dropbox_relpath).sharing_info.modified_by)
        display_name = dbx.users_get_account(submitter).name.display_name
        submitter_email = dbx.users_get_account(submitter).email
    except:
        logger.exception("ERROR with Dropbox api:")
    finally:
        return submitter_email, display_name

def setupFolders(tmpdir, inputfile, inputfilename, this_outfolder, inputfilename_noext):
    logger.info("Create tmpdir, create & cleanup project outfolder")

    # create new tmpdir, reset value for working file
    tmpdir = os_utils.setupTmpfolder(tmpdir)
    workingfile = os.path.join(tmpdir, inputfilename)
    os_utils.setupOutfolder(this_outfolder)

    # move inputfile to tmpdir as workingfile
    logger.info('Moving input file ({}) and template to tmpdir'.format(inputfilename))
    # os_utils.movefile(inputfile, workingfile)			# for production
    os_utils.copyFiletoFile(inputfile, workingfile)		# debug/testing only

    ziproot = os.path.join(tmpdir, "{}_unzipped".format(inputfilename_noext))		# the location where we unzip the input file
    template_ziproot = os.path.join(tmpdir, "macmillan_template_unzipped")
    stylereport_json = os.path.join(tmpdir, "stylereport.json")
    alerts_json = os.path.join(tmpdir, "alerts.json")

    return tmpdir, workingfile, ziproot, template_ziproot, stylereport_json, alerts_json

def copyTemplateandUnzipFiles(macmillan_template, tmpdir, workingfile, ziproot, template_ziproot):
    # move template to the tmpdir
    os_utils.copyFiletoFile(macmillan_template, os.path.join(tmpdir, os.path.basename(macmillan_template)))

    ### unzip the manuscript to ziproot, template to template_ziproot
    os_utils.rm_existing_os_object(ziproot, 'ziproot')
    os_utils.rm_existing_os_object(ziproot, 'template_ziproot')
    unzipDOCX.unzipDOCX(workingfile, ziproot)
    unzipDOCX.unzipDOCX(macmillan_template, template_ziproot)

def returnOriginal(this_outfolder, workingfile, inputfilename):
    # Return original file to user
    logger.info("Copying original file to outfolder/original_file dir")
    if not os.path.isdir(os.path.join(this_outfolder, "original_file")):
        os.makedirs(os.path.join(this_outfolder, "original_file"))
    os_utils.copyFiletoFile(workingfile, os.path.join(this_outfolder, "original_file", inputfilename))

def emailStyleReport(submitter_email, display_name, report_string, stylereport_txt, alerttxt_list, inputfilename, scriptname):
    logger.info("Putting together email to submitter... ")
    # adding this var so we know whether to re:email user if processing error comes up
    report_emailed = False
    # salutation = "Hello %s,\n\n" % display_name.split()[0]
    # contactus = "If you are unsure how to go about fixing these errors, check our Confluence page (), or email '%s' to reach out to the workflows team!\n" % cfg.support_email_address
    if os.path.exists(stylereport_txt):
        subject = usertext_templates.subjects()["success"].format(inputfilename=inputfilename)
        if alerttxt_list:
            # alert_intro = "Stylecheck-%s has successfully run on your file, '%s', with the following Warning(s) &/or Notice(s):\n\n" % (scriptname, inputfilename)
            alert_text = "\n".join(alerttxt_list)
            # preheader = salutation + alert_intro + contactus + alert_text
            bodytxt = usertext_templates.emailtxt()["success_with_alerts"].format(firstname=display_name.split()[0], scriptname=scriptname, inputfilename=inputfilename,
                report_string=report_string, helpurl="", support_email_address=cfg.support_email_address, alert_text=alert_text)
        else:
            bodytxt = usertext_templates.emailtxt()["success"].format(firstname=display_name.split()[0], scriptname=scriptname, inputfilename=inputfilename,
                report_string=report_string, helpurl="", support_email_address=cfg.support_email_address)

        # send our email!
        try:
            sendmail.sendMail([submitter_email], subject, bodytxt, [], [stylereport_txt])
            report_emailed = True
        except:
            raise
        # header = preheader + "\n\n________ STYLREPORT FOR '%s': _________\n\n" % inputfilename
        # bodytxt = header + report_string
    elif alerttxt_list:
        alert_text = "\n".join(alerttxt_list)
        subject = usertext_templates.subjects()["err"].format(inputfilename=inputfilename, scriptname=scriptname)
        bodytxt = usertext_templates.emailtxt()["error"].format(firstname=display_name.split()[0], scriptname=scriptname, inputfilename=inputfilename,
            report_string=report_string, helpurl="", support_email_address=cfg.support_email_address, alert_text=alert_text)

        # send our email!
        try:
            sendmail.sendMail([submitter_email], subject, bodytxt)
            report_emailed = True
        except:
            raise
        # alert_text = "\n".join(alerttxt_list)
        # err_intro = "There was a problem running Stylecheck-%s on your file '%s'.\n Please review any errors listed below for more information:\n\n" % (scriptname, inputfilename)
        # bodytxt = salutation + err_intro + alert_text + "\n\n" + contactus
    else:
        logger.warn("no style report or alerts fouund, so no email to send.")

    return report_emailed# exit function before mailing


def cleanupforReporterOrConverter(scriptname, this_outfolder, workingfile, inputfilename, report_dict, stylereport_txt, alerts_json, tmpdir, submitter_email, display_name):
    logger.info("Running cleanup, 'cleanupforReporterOrConverter'...")

    # 1 return original_file to outfolder
    returnOriginal(this_outfolder, workingfile, inputfilename)

    # 2 write our alertfile.txt if necessary
    if os.path.exists(alerts_json):
        logger.debug("Writing alerts.txt to outfolder")
        alerttxt_list = os_utils.writeAlertstoTxtfile(alerts_json, this_outfolder)
    else:
        logger.debug("Skipping write alerts.txt to outfolder (no alerts.json)")
        alerttxt_list=[]

    # 3 if report_dict has contents, write stylereport file & send email!:
    if report_dict:
        logger.debug("Writing stylereport.txt to outfolder")
        report_string = generate_report.generateReport(report_dict, stylereport_txt)
    else:
        logger.debug("Skipping write stylereport.txt to outfolder (empty report_dict)")
        report_string = ""

    # 4 and send stylereport and/or alerts as mail
    logger.debug("emailing stylereport &/or alerts ")
    report_emailed = emailStyleReport(submitter_email, display_name, report_string, stylereport_txt, alerttxt_list, inputfilename, scriptname)

    # 5 Rm tmpdir
    logger.debug("deleting tmp folder")
    # os_utils.rm_existing_os_object(tmpdir, 'tmpdir')		# comment out for testing / debug

    return report_emailed

def cleanupforValidator(this_outfolder, workingfile, inputfilename, report_dict, stylereport_txt, alerts_json):
    logger.info("Running cleanup, 'cleanupforValidator'...")

    # 1 if report_dict has contents, write stylereport file:
    logger.debug("Writing stylereport.txt to outfolder")
    if report_dict:
        generate_report.generateReport(report_dict, stylereport_txt)
        # and send stylereport as mail

    # 2 write our alertfile.txt if necessary
    if os.path.exists(alerts_json):
        logger.debug("Writing alerts.txt to outfolder")
        alerttxt_list = os_utils.writeAlertstoTxtfile(alerts_json, this_outfolder)
    else:
        logger.debug("Skipping write alerts.txt to outfolder (no alerts.json)")
        alerttxt_list=[]


def sendAlertEmail(scriptname, logfile, inputfilename, errs_duringcleanup=[]):
    logger.info("running 'sendAlertEmail' function...")
    try:
        # if we encountered error during exception handling we'll try to send another mail,
        #   this time without attachment, and let us know that some cleanup / alerts may need ot be done manually
        if errs_duringcleanup:
            logger.debug("this alert is for an error during exception cleanup")
            cleanup_errs = '\n'.join(errs_duringcleanup)
            subject = "Error cleaning up after Exception: running %s on %s" % (scriptname, inputfilename)
            bodytxt = textwrap.dedent("""\
            Error encountered cleaning up after an Exception: running '%s' on %s.
            The following items failed during 'cleanupException' method:
            %s
            """ % (scriptname, inputfilename, cleanup_errs))
            # send the mail!
            sendmail.sendMail(cfg.alert_email_address, subject, bodytxt)
        # just a normal alert email, including attached logfile
        else:
            subject = usertext_templates.subjects()["err"].format(inputfilename=inputfilename, scriptname=scriptname)
            bodytxt = "Error encountered while running '%s' on %s.\n\nSee attached logfile for details." % (scriptname, inputfilename)
            # send the mail!
            sendmail.sendMail(cfg.alert_email_address, subject, bodytxt, None, [logfile])
    except:
        logger.exception("ERROR sendAlertEmail function :(")
        raise   # is this necessary?

def cleanupException(this_outfolder, workingfile, inputfilename, alerts_json, tmpdir, logdir, inputfilename_noext, scriptname, logfile, report_emailed, submitter_email, display_name):
    logger.warn("POST-ERROR: Running cleanup, 'cleanupException'...")
    # setting defaults in case we encounter an exception during cleanup
    errs_duringcleanup = []

    # 1 send email to workflows
    try:
        logger.info("trying: emailing workflows to notify re: error")
        sendAlertEmail(scriptname, logfile, inputfilename)
    except:
        logger.exception("* while trying to send mail Traceback:")
        errs_duringcleanup.append("-email alert to workflows")

    # 2 write error to alerts json, write alertfile
    logger.info("trying: Writing error to alerts.json, and posting alerts.txt to outfolder")
    try:
        errstring = usertext_templates.alerts()["processing_alert"].format(scriptname=scriptname, support_email_address=cfg.support_email_address)
        os_utils.logAlerttoJSON(cfg.alerts_json, "error", errstring)
        alerttxt_list = os_utils.writeAlertstoTxtfile(alerts_json, this_outfolder)
    except:
        logger.exception("* writing alert to json and posting alertfile Traceback:")
        errs_duringcleanup.append("-write error alert to json, dump json alerts to errfile in OUT folder")

    # 3 save a copy of tmpdir to logdir for troubleshooting (since it will be deleted)
    logger.info("trying: Backing up tmpdir to logfolder")
    try:
        os_utils.copyDir(tmpdir, os.path.join(logdir, "tmpdir_%s" % inputfilename_noext))
    except:
        logger.exception("* backing up tmpdir to logfile Traceback:")
        errs_duringcleanup.append("-back up tmpdir to logfolder")

    # these two items only apply to converter and reporter
    if scriptname != "validator":
        # 4 return original_file to outfolder
        logger.info("trying: return original file to OUT folder")
        try:
            returnOriginal(this_outfolder, workingfile, inputfilename)
        except:
            logger.exception("* returning original to outfolder Traceback:")
            errs_duringcleanup.append("-returning original file to OUT folder")

        # email submitter
        if report_emailed == False and submitter_email:
            logger.info("trying: notify submitter")
            try:
                subject = usertext_templates.subjects()["err"].format(inputfilename=inputfilename, scriptname=scriptname)
                alert_text = usertext_templates.alerts()["processing_alert"].format(scriptname=scriptname, support_email_address=cfg.support_email_address)
                bodytxt = usertext_templates.emailtxt()["success"].format(firstname=display_name.split()[0], scriptname=scriptname, inputfilename=inputfilename,
                    helpurl="", support_email_address=cfg.support_email_address, alert_text=alert_text)
                sendmail.sendMail([submitter_email], subject, bodytxt)
            except:
                logger.exception("* returning original to outfolder Traceback")
                errs_duringcleanup.append("-sending err email to submitter")
        else:
            logger.info("skipping: notify submitter (report_email already sent)")

        # 5 Rm tmpdir to avoid interfering with next run
        logger.info("trying: delete tmp folder")
        try:
            os_utils.rm_existing_os_object(tmpdir, 'tmpdir')		# comment out for testing / debug
        except:
            logger.exception("* deleting tmp folder Traceback:")
            errs_duringcleanup.append("-deleting tmp folder")

    # try once more to send alert if we encountered cleanup errors
    if errs_duringcleanup:
        logger.error("ERRORS DURING CLEANUP: %s" % errs_duringcleanup)
        sendAlertEmail(scriptname, logfile, inputfilename, errs_duringcleanup)

# # only run if this script is being invoked directly
# if __name__ == '__main__':
#
#     submitter_email, display_name = getSubmitterViaAPI(cfg.inputfile)
#
#     print submitter_email, display_name
