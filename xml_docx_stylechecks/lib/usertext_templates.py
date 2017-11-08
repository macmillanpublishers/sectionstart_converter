######### IMPORT PY LIBRARIES
import logging
import textwrap

# initialize logger
logger = logging.getLogger(__name__)


# #---------------------  METHODS
# This method defines what goes in the StyleReport txt and mail outputs, in what order, + formatting.
# See the commented "SAMPLE RECIPE ENTRY" below for details on each field.  All fields should be optional,
#   though text, title or dict_category_name must be present for something to print
def emailtxt():
    templates = {
    	"success_html": textwrap.dedent("""\
            <html>
            <head></head>
            <body>
            <p>Hello {firstname},</p>

            <p>Stylecheck-{scriptname} has successfully run on your file, '{inputfilename}'!</p>
            <p>You can view the StyleReport below in this email, or download the attached file for a more formatted version.</p>
            <p>For help interpreting any errors, check <a href="{helpurl}">this Confluence page</a>, or email {support_email_address} to reach out to the workflows team!</p>


            <p>REPORT FOR {inputfilename}:</p>
            <hr/>
            <pre>
            {report_string}
            </pre>
            </body>
            </html>
            """),
    	"success": textwrap.dedent("""\
            Hello {firstname},

            Stylecheck-{scriptname} has successfully run on your file, '{inputfilename}'!
            You can view the StyleReport below in this email, or download the attached file for a more formatted version.
            For help interpreting any errors, try the guide on this Confluence page: {helpurl}, or email {support_email_address} to reach out to the workflows team!


            REPORT FOR {inputfilename}:
            ----------------------------------------------------------------------------------------------------------------------------------------------------------------

            {report_string}
            """),
    	"success_with_alerts": textwrap.dedent("""\
            Hello {firstname},

            "Stylecheck-{scriptname} has successfully run on your file, '{inputfilename}', with the following Warning(s) &/or Notice(s):

            --------------------------------------
            {alert_text}
            --------------------------------------


            Below is the stylereport output
            For help interpreting any errors, try the guide on this Confluence page: {helpurl}, or email '{support_email_address}' to reach out to the workflows team!


            REPORT FOR {inputfilename}:
            ----------------------------------------------------------------------------------------------------------------------------------------------------------------

            {report_string}
            """),
    	"error": textwrap.dedent("""\
            Hello {firstname},

            There was a problem running Stylecheck-{scriptname} on your file '{inputfilename}'.

            Please review Error(s) listed below:

            --------------------------------------
            {alert_text}
            --------------------------------------

            If you are unsure how to go about fixing these errors, try the guide on this Confluence page: {helpurl}, or email '{support_email_address}' to reach out to the workflows team!
            """),
    	"processing_error": textwrap.dedent("""\
            Hello {firstname},

            There was a problem running Stylecheck-{scriptname} on your file '{inputfilename}'.

            Please review Error(s) listed below:

            --------------------------------------
            {alert_text}
            --------------------------------------
            """)
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
        "notdocx": "This file is not a '.docx'. Only .docx files can be run through Stylecheck-{scriptname}.",
        "notstyled": "This .docx has {percent_styled} percent of paragraphs styled with Macmillan styles",
    	"protected": "This .docx has protection enabled.",
    	"r_err_oldtemplate": "You must attach the newest version of the macmillan style template before running the Style Report: (this .docx's version: {current_version}, template version: {template_version})",
        "v_has_template": "This document already has a template attached with section_start styles.",
        "v_newertemplate_avail": "Newer available version of the macmillan style template (this .docx's version: {current_version}, template version: {template_version})",
        "processing_alert": textwrap.dedent("""\
            An error was encountered while running '{scriptname}'_main.py. The workflows team has been notified of this error.
            If you don't hear from us within 2 hours, please email {support_email_address} for assistance.""")
    }
    return alerts

# #---------------------  MAIN
# # # only run if this script is being invoked directly (for testing)
# if __name__ == '__main__':
    # report_string = "Teststring pt1\n\nsome more text\tafterthought\n"
    # templates = usertext_templates()
    # print templates["report_success"].format(report_string=report_string, firstname="Jerff", scriptname="converter", inputfilename="test.docx", helpurl="www.Confluence.com", support_email_address="help@helpless.net")
