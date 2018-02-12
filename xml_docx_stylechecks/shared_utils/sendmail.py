import smtplib
import os
import sys
import logging
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders


# initialize logger
logger = logging.getLogger(__name__)


#---------------------  METHODS
# Note: the to-address, cc-address and attachments need to be list objects.
# cc_addresses and attachments are optional arguments
def sendMailBasic(port, smtp_address, from_email_address, always_bcc_address, to_addr_list, subject, bodytxt, cc_addr_list, attachfile_list, htmltxt=""):
    try:
        if htmltxt:
            msg = MIMEMultipart('related')
        else:
            msg = MIMEMultipart()
        msg['From'] = from_email_address
        msg['To'] = ','.join(to_addr_list)
        msg['Subject'] = subject
        if cc_addr_list:
            msg['Cc'] = ','.join(cc_addr_list)
			# the to_addr_list is used inthe sendmail cmd below and includes all recipients (including cc)
            to_addr_list = to_addr_list + cc_addr_list
            # add bcc:
        if always_bcc_address:
            to_addr_list = to_addr_list + [always_bcc_address]
        # setup handling for html-email with alternative text
        if htmltxt:
            msgAlternative = MIMEMultipart('alternative')
            msg.attach(msgAlternative)
            msgAlternative.attach(MIMEText(bodytxt, 'plain'))
            msgAlternative.attach(MIMEText(htmltxt, 'html'))
        else:
            msg.attach(MIMEText(bodytxt, 'plain'))

        if attachfile_list:
            for attachfile in attachfile_list:
                filename = "%s" % os.path.basename(attachfile)
                attachment = open(attachfile, "rb")
                part = MIMEBase('application', 'octet-stream')
                part.set_payload((attachment).read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
                msg.attach(part)

        server = smtplib.SMTP(smtp_address, port)
        text = msg.as_string()
        server.sendmail(from_email_address, to_addr_list, text)
        server.quit()
    except:
        logger.exception("MAILER ERROR ------------------ :")
        raise

def sendMail(to_addr_list, subject, bodytxt, cc_addr_list=None, attachfile_list=None, htmltxt=""):
    # moving common dependencies for this file for converter/reporter/validator sute of scripts into an outer method...
    #   so I can reuse this script for the independent process_watcher.
    try:
        ######### IMPORT LOCAL MODULES
        if __name__ == '__main__':
        	# to go up a level to read cfg when invoking from this script (for testing).
        	import imp
        	parentpath = os.path.join(sys.path[0], '..', 'cfg.py')
        	cfg = imp.load_source('cfg', parentpath)
        else:
        	import cfg

        ######### LOCAL DECLARATIONS
        with open(cfg.smtp_txt) as f:
            smtp_address = f.readline().strip()
        port = 25
        from_email_address = cfg.from_email_address
        always_bcc_address = cfg.always_bcc_address

        sendMailBasic(port, smtp_address, from_email_address, always_bcc_address, to_addr_list, subject, bodytxt, cc_addr_list, attachfile_list, htmltxt)

    except smtplib.SMTPConnectError:
        errstring = "Email send fail: 'SMTPConnectError' -- Email subject: '%s'" % subject
        logger.warn(errstring)
        logger.info("Email info: '%s', '%s'\n  '%s'" % (to_addr_list, subject, bodytxt))
        import os_utils as os_utils
        # write err to file
        os_utils.logAlerttoJSON(cfg.alerts_json, 'warning', errstring)
        alerttxt_list = os_utils.writeAlertstoTxtfile(cfg.alerts_json, cfg.this_outfolder)
        # print "LOG THIS EMAIL!: ",to_addr_list, subject, bodytxt # debug only

#---------------------  MAIN
# only run if this script is being invoked directly
# if __name__ == '__main__':
    # # set up debug log to console
    # logging.basicConfig(level=logging.DEBUG)
    #
    # # from_addr = "workflows@macmillan.com"
    # to_addr_list = ["your email address here"]
    # subject = "Test email"
    # bodytxt = "Did this work?\n\n\t(I hope?)"
    # cc_addr_list = ["cc address 1", "cc address 2"]
    #
    # sendmail(to_addr_list, subject, bodytxt)
    #
    # sendmail(to_addr_list, subject, bodytxt, cc_addr_list, [cfg.inputfile])
