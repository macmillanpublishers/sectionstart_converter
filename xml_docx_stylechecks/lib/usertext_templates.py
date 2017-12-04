######### IMPORT PY LIBRARIES
import logging
import textwrap

# initialize logger
logger = logging.getLogger(__name__)


# #---------------------  METHODS
# This method defines what goes in the StyleReport txt and mail outputs, in what order, + formatting.
def emailtxt():
    templates = {
    	"success_html": textwrap.dedent("""\
            <html>
            <head></head>
            <body>
            <p>Hello {firstname},</p>
            <p>Stylecheck-{scriptname} has processed your file, '{inputfilename}'!</p>
            <p>You can view the StyleReport in this email (below), or download the attached StyleReport.txt file if you prefer.<br/>
            {converter_txt}</p>
            <p>For help interpreting any errors in the report, take a look at <a href="{helpurl}">this page</a> on Confluence, or email {support_email_address} to reach out to the workflows team!</p>
            <p>&nbsp;</p>
            <p>Report for '{inputfilename}':</p>
            <hr/>
            <font size=2em>
            <pre>
            {report_string}
            </pre></font>
            </body>
            </html>
            """),
    	"success_with_alerts_html": textwrap.dedent("""\
            <html>
            <head></head>
            <body>
            <p>Hello {firstname},</p>
            <p>Stylecheck-{scriptname} has processed your file, '{inputfilename}', with the below Warning(s) &/or Notice(s):</p>
            <p>You can view the StyleReport in this email (below), or download the attached StyleReport.txt file if you prefer.<br/>
            {converter_txt}</p>
            <p>For help interpreting any errors, take a look at <a href="{helpurl}">this page</a> on Confluence, or email {support_email_address} to reach out to the workflows team!</p>
            <p>&nbsp;</p>
            <p>Warning(s) / Notice(s):<br/>
            --------------------------------------<p>
            <font size=2em>
            <pre>{alert_text}</pre></font>
            <p>--------------------------------------</p>
            <p>&nbsp;</p>
            <p>Report for '{inputfilename}':</p>
            <hr/>
            <font size=2em>
            <pre>
            {report_string}
            </pre></font>
            </body>
            </html>
            """),
    	"success": textwrap.dedent("""\
            Hello {firstname},

            Stylecheck-{scriptname} has processed your file, '{inputfilename}'!

            Please download and view the attached StyleReport.txt file to view info on your file.
            {converter_txt}

            For help interpreting any errors, try the guide on this Confluence page: {helpurl}, or email {support_email_address} to reach out to the workflows team!
            """),
    	"success_with_alerts": textwrap.dedent("""\
            Hello {firstname},

            Stylecheck-{scriptname} has processed your file, '{inputfilename}', with the below Warning(s) &/or Notice(s):

            Please download and view the attached StyleReport.txt file to view info on your file.
            {converter_txt}

            --------------------------------------
            {alert_text}
            --------------------------------------

            For help interpreting any errors, try the guide on this Confluence page: {helpurl}, or email '{support_email_address}' to reach out to the workflows team!
            """),
    	"error": textwrap.dedent("""\
            Hello {firstname},

            Stylecheck-{scriptname} could not process your file: '{inputfilename}'.

            Please review Error(s) listed below:

            --------------------------------------
            {alert_text}
            --------------------------------------

            If you are unsure how to go about fixing the above item(s), try the guide on this Confluence page: {helpurl}, or email '{support_email_address}' to reach out to the workflows team!
            """),
    	"processing_error": textwrap.dedent("""\
            Hello {firstname},

            Stylecheck-{scriptname} could not process your file: '{inputfilename}'.

            Please review Error(s) listed below:

            --------------------------------------
            {alert_text}
            --------------------------------------
            """),
    	"converter_txt": "You should find the 'converted' version of your .docx attached as well (if not, check the Stylecheck-Converter OUT folder)."
    }
    return templates

def subjects():
    subjects = {
    	"success": "StyleReport for '{inputfilename}'",
    	"err": "Error running Stylecheck-{scriptname} for '{inputfilename}'"
    }
    return subjects

def alerts():
    alerts = {
        # Error - self explanatory
        "notdocx": "This file is not a '.docx'. Only .docx files can be run through Stylecheck-{scriptname}.",
        # Error - self explanatory. The extra percent sign is to escape the othe r% (it's a wildcard in python)
        "notstyled": "This .docx has less than 50% of paragraphs styled with Macmillan styles, so cannot be processed.",
        # Error - self explanatory
    	"protected": "This .docx has '{protection}' protection enabled. Please disable protection and try again!",
        # Error - for 'Reporter' only
    	"r_err_oldtemplate": "You must attach the most recent macmillan style template before running the Style Report: (this .docx's version: {current_version}, current version: {template_version})",
        # Error - Converter only.
        "c_has_template": "This document already has the most recent style template attached, if you think this does need conversion, contact {support_email_address}.",
        # Warning / Notice: unaccepted_tcs
        "c_unaccepted_tcs": "We found un-reviewed tracked changes in this document. We went ahead and inserted section-starts, but if things look significantly off, accept/reject tracked changes and run converter again!",
        # Warning: unaccepted_tcs
        "v_unaccepted_tcs": "We found un-reviewed tracked changes in this document. All tracked-changes were accepted",
        # Warning: unaccepted_tcs
        "r_unaccepted_tcs": "We found un-reviewed tracked changes in this document. Accepting or rejecting all pending Tracked Changes helps ensure the accuracy of the StyleReport.",
        # Notice: trackchange_enabled
        "trackchange_enabled": "'Track Changes' feature is currently enabled for this document.",
        # Notice - Converter only
        "c_newertemplate_avail": "There is a newer version of the macmillan style template available (this .docx's version: {current_version}, template version: {template_version})",
        # Notice - Validator only
        "v_newertemplate_avail": "There was a newer version of the macmillan style template available, attached during processing (this .docx's version: {current_version}, template version: {template_version})",
        # Fatal error (untrapped crash)
        "processing_alert": textwrap.dedent("""\
            An error was encountered while running 'Stylecheck-{scriptname}'. The workflows team has been notified of this error.
            If you don't hear from us within 2 hours, please email {support_email_address} for assistance.""")
    }
    return alerts

# #---------------------  MAIN
# # # only run if this script is being invoked directly (for testing)
# if __name__ == '__main__':
    # report_string = "Teststring pt1\n\nsome more text\tafterthought\n"
    # templates = usertext_templates()
    # print templates["report_success"].format(report_string=report_string, firstname="Jerff", scriptname="converter", inputfilename="test.docx", helpurl="www.Confluence.com", support_email_address="help@helpless.net")
