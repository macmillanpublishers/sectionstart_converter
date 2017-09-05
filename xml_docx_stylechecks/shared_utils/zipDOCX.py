import os
import zipfile
import shutil
import re
import logging
import xml.etree.ElementTree as ET
from sys import argv


# # # # # initialize logger
logger = logging.getLogger(__name__)


# # # # # METHODS
def getZipFiles(ziproot):
	# get a list of all files in the zip_root with relative paths
	try:
		zipmanifest = []
		for path, subdirs, files in os.walk(ziproot):
			for name in files:
				# exclude unwanted files here
				if name != ".DS_Store":
					fullpath = os.path.join(path, name)
					# get the path of file (relative to ziproot) and add to zip manifest
					zipmanifest.append(os.path.relpath(fullpath, ziproot))
		return zipmanifest
	except Exception, e:  	
		logger.error('Failed to get zip manifest', exc_info=True)
 		sys.exit(1)


def zipFiles(ziproot, zipmanifest, finaldocx):
	# Create the new zip file and add all the files into the archive
	# (the compression parameter here is optional, but without it files are ~5x larger;
	# 	this seems to roughly match what MS is using)
	try:
		with zipfile.ZipFile(finaldocx, "w", compression=zipfile.ZIP_DEFLATED) as docx:
			for filename in zipmanifest:
				docx.write(os.path.join(ziproot,filename), filename)
	except Exception, e:  	
		logger.error('Failed to zip-up .docx', exc_info=True)
 		sys.exit(1)


# combining the prior 2 methods into a single call
def zipDOCX(ziproot, finaldocx):
	logger.info("zipping '%s' into '%s'" % (ziproot, finaldocx))
	zipmanifest = getZipFiles(ziproot)
	zipFiles(ziproot, zipmanifest, finaldocx)	


# # # # # RUN
# for running this script as a standalone:
if __name__ == '__main__':
	# handle args as a standalone
	ziproot = argv[1]
	finaldocx = argv[2]

	# set up debug log to console
	logging.basicConfig(level=logging.DEBUG)

	zipDOCX(ziproot, finaldocx)