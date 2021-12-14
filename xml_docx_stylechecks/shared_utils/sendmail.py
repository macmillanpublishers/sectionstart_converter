import smtplib
import os
import sys
import logging
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders
import socket
from decorators import retry

# initialize logger
logger = logging.getLogger(__name__)
seconds_smtp_timeout = 30

#---------------------  METHODS
# Note: the to-address, cc-address and attachments need to be list objects.
# cc_addresses and attachments are optional arguments
@retry()
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
        server = smtplib.SMTP(smtp_address, port, timeout=seconds_smtp_timeout)
        text = msg.as_string()
        server.sendmail(from_email_address, to_addr_list, text)
        server.quit()
    except socket.gaierror, socket.timeout:
        exc_type, value, traceback = sys.exc_info()
        logger.error("'sendMailBasic' exception, type: '{}': {}. Reraising".format(exc_type.__name__, value))
        raise
    except:
        exc_type, value, traceback = sys.exc_info()
        logger.error("Unexpected exception with 'sendMailBasic', type: '{}': {}. Reraising".format(exc_type.__name__, value))
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

        # send mails unless we disabled via cfg.disable_sendmail for local run or testing
        if cfg.disable_sendmail == False:
            with open(cfg.smtp_txt) as f:
                smtp_address = f.readline().strip()
            port = 25
            sendMailBasic(port, smtp_address, cfg.from_email_address, cfg.always_bcc_address, to_addr_list, subject, bodytxt, cc_addr_list, attachfile_list, htmltxt)
        elif cfg.disable_sendmail == True:
            logger.info("* skipping 'sendmail' because we are in 'local' mode.")
            if cfg.loglevel != 'DEBUG':
                logger.info("* (set loglevel to DEBUG to see suppressed email addressees, subject, body, attachment list)")
        logger.debug("    email to: {}".format(to_addr_list))
        logger.debug("    email cc: {}".format(cc_addr_list))
        logger.debug("    email subject: {}".format(subject))
        logger.debug("    email content: {}".format(bodytxt))

    except:
        exc_type, value, traceback = sys.exc_info()
        logger.error("Exception sending email: {}, {}. Reraising\nset log level to DEBUG to see traceback, or review reraised traceback in Main below"
            .format(exc_type.__name__, value))
        logger.debug("(Traceback for debug)", exc_info=True)
        logger.info("  Details of unsent email:")
        logger.info("    email to: {}".format(to_addr_list))
        logger.info("    email cc: {}".format(cc_addr_list))
        logger.info("    email subject: {}".format(subject))
        logger.info("    email content: {}".format(bodytxt))
        raise

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
    # sendMail(to_addr_list, subject, bodytxt)
    #
    # sendMail(to_addr_list, subject, bodytxt, cc_addr_list, [cfg.inputfile])
