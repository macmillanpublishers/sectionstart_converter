######### IMPORT PY LIBRARIES
import os
import shutil
import re
import uuid
import json
import sys
import collections
import logging
import textwrap
import time
import imp


######### IMPORT LOCAL MODULES
if __name__ == '__main__':
    # to go up a level to read cfg (and other files) when invoking from this script (for testing).
    cfgpath = os.path.join(sys.path[0], '..', 'cfg.py')
    osutilspath = os.path.join(sys.path[0], '..', 'shared_utils', 'os_utils.py')
    apiPOSTpath = os.path.join(sys.path[0], '..', 'shared_utils', 'api_POST_to_camel.py')
    unzipDOCXpath = os.path.join(sys.path[0], '..', 'shared_utils', 'unzipDOCX.py')
    sendmailpath = os.path.join(sys.path[0], '..', 'shared_utils', 'sendmail.py')
    cfg = imp.load_source('cfg', cfgpath)
    import generate_report  # this is in the same dir so needs no direction for relative import
    import usertext_templates
    os_utils = imp.load_source('os_utils', osutilspath)
    apiPOST = imp.load_source('os_utils', apiPOSTpath)
    unzipDOCX = imp.load_source('unzipDOCX', unzipDOCXpath)
    sendmail = imp.load_source('sendmail', sendmailpath)
else:
    import cfg
    import lib.generate_report as generate_report
    import lib.usertext_templates as usertext_templates
    import shared_utils.api_POST_to_camel as apiPOST
    import shared_utils.os_utils as os_utils
    import shared_utils.unzipDOCX as unzipDOCX
    import shared_utils.sendmail as sendmail


######### LOCAL DECLARATIONS
# initialize logger
logger = logging.getLogger(__name__)
processwatch_file = cfg.processwatch_file
alerts_json = cfg.alerts_json


#---------------------  METHODS
# alert level should be 'error', 'warning' or 'notice'
def setAlert(alertlevel, ut_template_name, format_dict={}, suppress_logger=False):
    logger.debug('logging alert: "{}" with alerttext template "{}" and format params: {}'.format(alertlevel, ut_template_name, format_dict))
    alert_string = usertext_templates.alerts()[ut_template_name].format(**format_dict)
    os_utils.logAlerttoJSON(alerts_json, alertlevel, alert_string)
    if suppress_logger == False:
        if alertlevel == 'error':
            logger.error("* {}".format(alert_string))
        else:
            logger.warn("* {}".format(alert_string))

def setupforReporterOrConverter(inputfile, inputfilename, workingfile, this_outfolder, inputfile_ext):
    # get submitter name, email
    submitter_email, display_name = getSubmitterViaAPI(inputfile)
    logger.info("Submitter name:'%s', email: '%s'" % (submitter_email, display_name))

    # move inputfile to tmpdir as workingfile
    logger.info('Moving input file ({}) and template to tmpdir'.format(inputfilename))
    if cfg.leave_infile == False:
        os_utils.moveFile(inputfile, workingfile)   # cleanup, for production
    elif cfg.leave_infile == True:
        os_utils.copyFiletoFile(inputfile, workingfile) # debug/local testing
        logger.info("* retaining infile at initial path (running in 'local' mode).")

    # cleanup outfolder (archive existing)
    logger.info("Cleaning up existing outfolder")
    os_utils.setupOutfolder(this_outfolder)

    if inputfile_ext != ".docx":
        logger.warning("This file is not a .docx :(")
        notdocx = True
    else:
        notdocx = False

    return submitter_email, display_name, notdocx

def copyTemplateandUnzipFiles(macmillan_template, tmpdir, workingfile, ziproot, template_ziproot):
    setup_ok = True
    # move template to the tmpdir
    os_utils.copyFiletoFile(macmillan_template, os.path.join(tmpdir, os.path.basename(macmillan_template)))

    # remove existing alerts_json (really only matters for testing)
    os_utils.rm_existing_os_object(cfg.alerts_json, 'alerts.json')
    os_utils.rm_existing_os_object(os.path.join(tmpdir, cfg.err_fname), cfg.err_fname)
    os_utils.rm_existing_os_object(os.path.join(tmpdir, cfg.warn_fname), cfg.warn_fname)
    os_utils.rm_existing_os_object(os.path.join(tmpdir, cfg.notice_fname), cfg.notice_fname)

    # if inputfilename had unsupported chars, create a sanitized-fname workingfile in tmpdir
    if cfg.inputfile != workingfile:
        os_utils.copyFiletoFile(cfg.inputfile, workingfile)

    ### unzip the manuscript to ziproot, template to template_ziproot
    os_utils.rm_existing_os_object(ziproot, 'ziproot')
    os_utils.rm_existing_os_object(ziproot, 'template_ziproot')
    # putting a try block here so we can raise proper alert as needed
    try:
        errdict = {'no_filelist': 'cannot unzip; no document namelists',
                    'not_docfile': 'cannot unzip; not a Word doctype'}
        unzipDOCX.unzipDOCX(workingfile, ziproot, errdict)
        unzipDOCX.unzipDOCX(macmillan_template, template_ziproot, errdict)
    except Exception as e:
        setup_ok = False
        if str(e) == errdict['no_filelist']:
            setAlert('error', 'no_unzip_files', {})
        # this bit should currently be prescreened via th api, shouldn't come up
        elif str(e) == errdict['not_docfile']:
            setAlert('error', 'notdocx', {}, True)
        else:
            raise
    return setup_ok

def returnOriginal(this_outfolder, workingfile, inputfilename):
    # Return original file to user
    logger.info("Copying original file to outfolder/original_file dir")
    if not os.path.isdir(os.path.join(this_outfolder, "original_file")):
        os.makedirs(os.path.join(this_outfolder, "original_file"))
    os_utils.copyFiletoFile(workingfile, os.path.join(this_outfolder, "original_file", inputfilename))

def emailStyleReport(submitter_email, display_name, report_string, stylereport_txt, alerttxt_list, inputfilename, scriptname, newdocxfile):
    logger.info("Putting together email to submitter... ")
    # adding this var so we know whether to re:email user if processing error comes up
    report_emailed = False
    # set display_name for greeting
    if display_name:
        firstname=display_name.split()[0]
        to_string = "%s <%s>" % (display_name, submitter_email)
    else:
        firstname="Sir or Madam"
        to_string = submitter_email
    # Build email via this path if we have a style_report
    if os.path.exists(stylereport_txt):
        subject = usertext_templates.subjects()["success"].format(inputfilename=inputfilename)
        if scriptname == 'converter':
            converter_txt = usertext_templates.emailtxt()["converter_txt"]
        else:
            converter_txt = ""
        # if we have alerts / warnings /notices, include them
        if alerttxt_list:
            alert_text = "\n".join(alerttxt_list)
            bodytxt = usertext_templates.emailtxt()["success_with_alerts"].format(firstname=firstname, scriptname=scriptname.title(), inputfilename=inputfilename,
                report_string=report_string, helpurl=cfg.helpurl, support_email_address=cfg.support_email_address, alert_text=alert_text, converter_txt=converter_txt)
            htmltxt = usertext_templates.emailtxt()["success_with_alerts_html"].format(firstname=firstname, scriptname=scriptname.title(), inputfilename=inputfilename,
                report_string=report_string, helpurl=cfg.helpurl, support_email_address=cfg.support_email_address, alert_text=alert_text, converter_txt=converter_txt)
        # no alerts, printing just the report
        else:
            bodytxt = usertext_templates.emailtxt()["success"].format(firstname=firstname, scriptname=scriptname.title(), inputfilename=inputfilename,
                report_string=report_string, helpurl=cfg.helpurl, support_email_address=cfg.support_email_address, converter_txt=converter_txt)
            htmltxt = usertext_templates.emailtxt()["success_html"].format(firstname=firstname, scriptname=scriptname.title(), inputfilename=inputfilename,
                report_string=report_string, helpurl=cfg.helpurl, support_email_address=cfg.support_email_address, converter_txt=converter_txt)
        # send our email!
        try:
            if os.path.exists(newdocxfile):
                # sendmail.sendMail([to_string], subject, bodytxt, [], [stylereport_txt, newdocxfile])
                sendmail.sendMail([to_string], subject, bodytxt, [], [stylereport_txt, newdocxfile], htmltxt)
            else:
                # sendmail.sendMail([to_string], subject, bodytxt, [], [stylereport_txt])
                sendmail.sendMail([to_string], subject, bodytxt, [], [stylereport_txt], htmltxt)
            report_emailed = True
        except:
            raise
    # Build email via this path if we have NO style_report but YES alerts
    elif alerttxt_list:
        alert_text = "\n".join(alerttxt_list)
        subject = usertext_templates.subjects()["err"].format(inputfilename=inputfilename, scriptname=scriptname.title())
        bodytxt = usertext_templates.emailtxt()["error"].format(firstname=firstname, scriptname=scriptname.title(), inputfilename=inputfilename,
            report_string=report_string, helpurl=cfg.helpurl, support_email_address=cfg.support_email_address, alert_text=alert_text)

        # send our email!
        try:
            sendmail.sendMail([to_string], subject, bodytxt)
            report_emailed = True
        except:
            raise
    # nothing to send, skipping
    else:
        logger.warn("no style report or alerts found, so no email to send.")
    return report_emailed

def postFilesToOutfolder(stylereport_txt, newdocxfile, alertfile):
    logger.info("posting files to outfolder for 'direct' run, via camel api...")
    # set api url
    dest_folder = os.path.basename(cfg.this_outfolder)
    posturldict = os_utils.readJSON(cfg.post_urls_json)
    api_POSTurl = posturldict['rsvalidate']
    if os.path.exists(cfg.staging_file):
        api_POSTurl = posturldict['rsvalidate_stg']
    api_POSTurl = '{}?folder={}'.format(api_POSTurl, dest_folder)
    # send files
    try:
        apipost_result, api_success = {}, True
        if os.path.exists(stylereport_txt):
            apipost_result['stylereport'] = apiPOST.apiPOST(stylereport_txt, api_POSTurl)
        if os.path.exists(newdocxfile):
            apipost_result['newdocxfile'] = apiPOST.apiPOST(newdocxfile, api_POSTurl)
        if os.path.exists(alertfile):
            apipost_result['alertfile'] = apiPOST.apiPOST(alertfile, api_POSTurl)
        for k, v in apipost_result.items():
            if v != 'Success':
                api_success = False
        if api_success == False:
            raise ValueError("apipost_result(s) include non-success status: {}".format(apipost_result))
        return api_success
    except:
        raise

def cleanupforReporterOrConverter(scriptname, this_outfolder, workingfile, inputfilename, report_dict, stylereport_txt, alerts_json, tmpdir, submitter_email, display_name, original_inputfilename, newdocxfile=""):
    logger.info("Running cleanup, 'cleanupforReporterOrConverter'...")

    # 1 return original_file to outfolder
    if cfg.runtype != 'direct':
        returnOriginal(this_outfolder, workingfile, original_inputfilename)

    # 2 write our alertfile.txt if necessary
    if os.path.exists(alerts_json):
        logger.debug("Writing alerts.txt to outfolder")
        alerttxt_list, alertfile = os_utils.writeAlertstoTxtfile(alerts_json, this_outfolder, cfg.err_fname, cfg.warn_fname, cfg.notice_fname)
    else:
        logger.debug("Skipping write alerts.txt to outfolder (no alerts.json)")
        alerttxt_list, alertfile=[], ''

    # 3 if report_dict has contents, write stylereport file & send email!:
    if report_dict:
        logger.debug("Writing stylereport.txt to outfolder")
        report_string = generate_report.generateReport(report_dict, stylereport_txt, scriptname)
    else:
        logger.debug("Skipping write stylereport.txt to outfolder (empty report_dict)")
        report_string = ""

    # 4 and send stylereport and/or alerts as mail
    if not os.environ.get('TRANSFORM_TEST_FLAG'):
        logger.debug("emailing stylereport &/or alerts ")
        report_emailed = emailStyleReport(submitter_email, display_name, report_string, stylereport_txt, alerttxt_list, inputfilename, scriptname, newdocxfile)
    else:
        report_emailed = True

    # 4.5 if this is a 'direct' run, sendfiles to true outfolder via api
    if cfg.runtype == 'direct' and not os.environ.get('TRANSFORM_TEST_FLAG') and cfg.disable_POST == False:
        logger.debug("sending files to outfolder for direct run")
        api_success = postFilesToOutfolder(stylereport_txt, newdocxfile, alertfile)
    else:
        api_success = True
        if cfg.disable_POST == True:
            logger.info("* skipping POST finished files via api (running in 'local' mode).")

    # 5 Rm tmpdir
    if cfg.preserve_tmpdir == False or cfg.runtype != 'direct':    # leave tmpdir for debug/testing
        logger.info("deleting tmp folder")
        os_utils.rm_existing_os_object(tmpdir, 'tmpdir')
    elif cfg.preserve_tmpdir == True:
        logger.info("* skipping tmpdir cleanup (running in 'local' mode).")

    # 6 Rm processwatch_file
    logger.debug("deleting processwatch_file")
    os_utils.rm_existing_os_object(processwatch_file, 'processwatch_file')
    return report_emailed

def cleanupforValidator(this_outfolder, workingfile, inputfilename, report_dict, stylereport_txt, alerts_json, scriptname):
    logger.info("Running cleanup, 'cleanupforValidator'...")

    # 1 if report_dict has contents, write stylereport file:
    logger.debug("Writing stylereport.txt to outfolder")
    if report_dict:
        generate_report.generateReport(report_dict, stylereport_txt, scriptname)
        # and send stylereport as mail

    # 2 write our alertfile.txt if necessary
    if os.path.exists(alerts_json):
        logger.debug("Writing alerts.txt to outfolder")
        alerttxt_list, alertfile = os_utils.writeAlertstoTxtfile(alerts_json, this_outfolder, cfg.err_fname, cfg.warn_fname, cfg.notice_fname)
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
            subject = "Error cleaning up after Exception: Stylecheck-%s, file: '%s'" % (scriptname, inputfilename)
            bodytxt = textwrap.dedent("""\
            Date/Time: %s

            Error encountered cleaning up after an Exception: running '%s' on %s.
            The following item(s) failed during 'cleanupException' method:
            %s
            """ % (time.strftime("%y%m%d-%H%M%S"), scriptname, inputfilename, cleanup_errs))
            # send the mail!
            sendmail.sendMail([cfg.alert_email_address], subject, bodytxt)
        # just a normal alert email, including attached logfile
        else:
            subject = usertext_templates.subjects()["err"].format(inputfilename=inputfilename, scriptname=scriptname.title())
            bodytxt = "Date/Time: %s\n\nError encountered while running '%s' on file '%s'.\n\nSee attached logfile for details." % (time.strftime("%y%m%d-%H%M%S"), scriptname, inputfilename)
            # send the mail!
            sendmail.sendMail([cfg.alert_email_address], subject, bodytxt, None, [logfile])
    except:
        logger.error("ERROR sendAlertEmail function :(")
        raise

def cleanupException(this_outfolder, workingfile, inputfilename, alerts_json, tmpdir, logdir, inputfilename_noext, scriptname, logfile, report_emailed, submitter_email, display_name, original_inputfilename):
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
        setAlert('error', 'processing_alert', {'scriptname':scriptname.title(), 'support_email_address':cfg.support_email_address}, True)
        alerttxt_list, alertfile = os_utils.writeAlertstoTxtfile(alerts_json, this_outfolder, cfg.err_fname, cfg.warn_fname, cfg.notice_fname)
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
    if not scriptname.startswith("validator"):
        # 4 return original_file to outfolder
        if cfg.runtype != 'direct':
            logger.info("trying: return original file to OUT folder")
            try:
                returnOriginal(this_outfolder, workingfile, original_inputfilename)
            except:
                logger.exception("* returning original to outfolder Traceback:")
                errs_duringcleanup.append("-returning original file to OUT folder")

        # email submitter
        if report_emailed == False and submitter_email:
            logger.info("trying: notify submitter")
            try:
                if display_name:
                    firstname=display_name.split()[0]
                else:
                    firstname="Sir or Madam"
                subject = usertext_templates.subjects()["err"].format(inputfilename=inputfilename, scriptname=scriptname.title())
                alert_text = usertext_templates.alerts()["processing_alert"].format(scriptname=scriptname.title(), support_email_address=cfg.support_email_address)
                bodytxt = usertext_templates.emailtxt()["processing_error"].format(firstname=firstname, scriptname=scriptname.title(), inputfilename=inputfilename,
                    helpurl=cfg.helpurl, support_email_address=cfg.support_email_address, alert_text=alert_text)
                sendmail.sendMail([submitter_email], subject, bodytxt)
            except:
                logger.exception("* returning original to outfolder Traceback")
                errs_duringcleanup.append("-sending err email to submitter")
        else:
            logger.info("skipping: notify submitter (report_email already sent)")

        # # # \/ this is moot for direct runs as currently constituted
        # # 5 Rm tmpdir to avoid interfering with next run
        # logger.info("trying: delete tmp folder")
        # try:
        #     if cfg.preserve_tmpdir == False:    # leave tmpdir for debug/testing
        #         os_utils.rm_existing_os_object(tmpdir, 'tmpdir')
        # except:
        #     logger.exception("* deleting tmp folder Traceback:")
        #     errs_duringcleanup.append("-deleting tmp folder")

        # 6 Rm processwatch_file
        logger.debug("deleting processwatch_file")
        try:
            os_utils.rm_existing_os_object(processwatch_file, 'processwatch_file')
        except:
            logger.exception("* deleting processwatch_file Traceback:")
            errs_duringcleanup.append("-deleting processwatch_file")

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
