from sys import argv
# make sure to insall lxml: sudo pip install lxml
from lxml import etree
# from lxml.etree import QName
# from lxml.builder import E
# from lxml.builder import ElementMaker

if __name__ == '__main__':
	ziproot = argv[1]

import os
import shutil
import re
import uuid
import json
import sys

# ---------------------- LOCAL DECLARATIONS

xmlfile = 'word/document.xml'
docxml = os.path.join(ziproot, xmlfile)
wnamespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
w14namespace = 'http://schemas.microsoft.com/office/word/2010/wordml'
namespaces = {'w': wnamespace, 'w14': w14namespace}
tree = etree.parse(docxml)
root = tree.getroot()
pagebreakstring = ".//*w:br[@w:type='page']"
changelog = []
ziproot_basename = os.path.basename(ziproot)
changelogfile = "%s_edits.json" % ziproot


#---------------------  METHODS

# return all text from a paragraph
# a parent function to search for a para with a certain style, then call two other functions
# def deletePageBreakAndInsertPara(searchstyle, insertstyle):
def macmillanStyleCheck():
	searchstyle = "Dedicationded"
	searchstring = ".//*w:pStyle[@w:val='Dedicationded']" #% searchstyle
	print len(root.findall(searchstring, namespaces))
	for parastyle in root.findall(searchstring, namespaces):
		print parastyle
		# get parent paragraph (up two levels from a paragraph style, past the pPr)
		# para = parastyle.getparent().getparent()

		# # delete preceding manual pagebreaks!
		# deletePrecedingPageBreak(para)

		# # insert a para styled with 'replacestyle' before para with 'searchstyle'
		# insertStyledParaBefore(para,searchstyle,insertstyle,"Heading test")

print "starting"

macmillanStyleCheck()

print "ending"


