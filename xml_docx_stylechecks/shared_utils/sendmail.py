import smtplib
import os
import sys
import logging
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
# from email.MIMEBase import MIMEBase
# from email import encoders

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
username = "username"
password = "password"

#---------------------  METHODS
def sendmail(from_addr, to_addr, subject, bodytxt):
    try:
        msg = MIMEMultipart()
        msg['From'] = from_addr
        msg['To'] = to_addr
        msg['Subject'] = subject
        msg.attach(MIMEText(bodytxt, 'plain'))

        server = smtplib.SMTP(smtp_address, port)
        # server.starttls()
        # server.login(username, password)
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()
    except:
        logger.exception("ERROR ------------------ :")

#---------------------  MAIN
# only run if this script is being invoked directly
if __name__ == '__main__':
    # set up debug log to console
    logging.basicConfig(level=logging.DEBUG)

    from_addr = "workflows@macmillan.com"
    to_addr = "matthew.retzer@macmillan.com"
    subject = "Big test"
    bodytxt = "Did this work???!?!??\n\n\t(I hope?)"

    sendmail(from_addr, to_addr, subject, bodytxt)
