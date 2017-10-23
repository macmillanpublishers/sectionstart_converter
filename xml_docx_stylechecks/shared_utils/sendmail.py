import smtplib
import os
import sys
import logging
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders

######### IMPORT LOCAL MODULES
if __name__ == '__main__':
	# to go up a level to read cfg when invoking from this script (for testing).
	import imp
	parentpath = os.path.join(sys.path[0], '..', 'cfg.py')
	cfg = imp.load_source('cfg', parentpath)
else:
	import cfg

# initialize logger
logger = logging.getLogger(__name__)

######### LOCAL DECLARATIONS
with open(cfg.smtp_txt) as f:
    smtp_address = f.readline()
port = 25
from_email_address = cfg.from_email_address

#---------------------  METHODS
# Note: the to-address, cc-address and attachments need to be list objects.
# cc_addresses and attachments are optional arguments
def sendMail(to_addr_list, subject, bodytxt, cc_addr_list=None, attachfile_list=None):
    try:
        # print "EMAIL!: ",to_addr_list, subject, bodytxt # debug only
        msg = MIMEMultipart()
        msg['From'] = from_email_address
        msg['To'] = ','.join(to_addr_list)
        msg['Subject'] = subject
        if cc_addr_list:
            msg['Cc'] = ','.join(cc_addr_list)
			# the to_addr_list is used inthe sendmail cmd below and includes all recipients (including cc)
            to_addr_list = to_addr_list + cc_addr_list
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

#---------------------  MAIN
# only run if this script is being invoked directly
if __name__ == '__main__':
    # set up debug log to console
    logging.basicConfig(level=logging.DEBUG)

    # from_addr = "workflows@macmillan.com"
    to_addr_list = ["your email address here"]
    subject = "Test email"
    bodytxt = "Did this work?\n\n\t(I hope?)"
    cc_addr_list = ["cc address 1", "cc address 2"]

    sendmail(to_addr_list, subject, bodytxt)

    sendmail(to_addr_list, subject, bodytxt, cc_addr_list, [cfg.inputfile])
