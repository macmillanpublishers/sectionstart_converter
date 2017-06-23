from sys import argv

ziproot = argv[1]
finaldocx = argv[2]
print ziproot
print finaldocx

import os
import zipfile
import shutil
import re
import xml.etree.ElementTree as ET

def get_zip_files(self):
	# should add a test for if file exists,return empty manifest, raise alert. Or just try / except?
	# get a list of all files in the zip_root with relative paths
	zipmanifest = []
	for path, subdirs, files in os.walk(self):
		for name in files:
			# exclude unwanted files here
			if name != ".DS_Store":
				fullpath = os.path.join(path, name)
				# get the path of file (relative to ziproot) and add to zip manifest
				zipmanifest.append(os.path.relpath(fullpath, self))

	return zipmanifest	


def zip_docx(ziproot, zipmanifest, finaldocx):
	# Create the new zip file and add all the files into the archive
	# (the compression parameter here is optional, but without it files are ~5x larger;
	# 	this seems to roughly match what MS is using)
	with zipfile.ZipFile(finaldocx, "w", compression=zipfile.ZIP_DEFLATED) as docx:
		for filename in zipmanifest:
			docx.write(os.path.join(ziproot,filename), filename)


zipmanifest = get_zip_files( ziproot )

zip_docx(ziproot, zipmanifest, finaldocx)
