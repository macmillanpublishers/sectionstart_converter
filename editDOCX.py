from sys import argv
# make sure to insall lxml: sudo pip install lxml
from lxml import etree
# from lxml.etree import QName
# from lxml.builder import E
# from lxml.builder import ElementMaker

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
changelogjson = ""


#---------------------  METHODS

# return all text from a paragraph
def paraTxt(para):
    try:
        paratext = "".join([x for x in para.itertext()])
    except:
        paratext = "n-a"

    return paratext

# return a dict of neighboring para elements and their text
def getNeighborParas(para):
    pneighbors = {}
    try:
        pneighbors['prev'] = para.getprevious()
        len(pneighbors['prev'].tag)
        pneighbors['prevtext'] = paraTxt(pneighbors['prev'])
    except:
        pneighbors['prev'] = ""
        pneighbors['prevtext'] = ""
    try:
        pneighbors['next'] = para.getnext()
        len(pneighbors['next'].tag)
        pneighbors['nexttext'] = paraTxt(pneighbors['next'])
    except:
        pneighbors['next'] = ""
        pneighbors['nexttext'] = ""

    return pneighbors

# return the w14:paraId attribute's value for a paragraph
def getParaId(para):
    if len(para):
        attrib_id_key = '{%s}paraId' % w14namespace
        para_id = para.get(attrib_id_key)
    else:
        para_id = 'n-a' # this could happen if we are trying to get id of a prev or next para that does not exist

    return para_id

# return the index value of a paragraph (within the body/root)
def getParaIndex(para):
    if len(para):
        para_index = para.getparent().index(para)
    else:
        para_index = 'n-a' # this could happen if we are trying to get id of a prev or next para that does not exist

    return para_index

# a method to be called every time we make a change in the xml, to log the paragraph id and the change we made
def trackEdit(para,para_action,description):
    # 'para_action' parameter values must be one of the following: 'insert','remove', or 'edit'
    # if a paragraph was removed, the paraid value will be calculated for the previous para element
    para_action_values = ['insert','remove','edit']
    if para_action in para_action_values:
        newedit_dict = {}
        if para_action == 'remove':
            pneighbors = getNeighborParas(para)
            newedit_dict['para_id'] = getParaId(pneighbors['prev'])
        else:
            newedit_dict['para_id'] = getParaId(para)
        newedit_dict['para_action'] = para_action
        newedit_dict['description'] = description
        changelog.append(newedit_dict.copy())
    else:
        print "unnacceptable value entered for changelog: " + "--" + paraid + "--" + para_action + "--" + description

# check paragraph directly preceding the selected (passed) para to get rid of pagebreaks/pb-paras where approrpiate
def deletePrecedingPageBreak(para):
    pneighbors = getNeighborParas(para)
    if len(pneighbors['prev']):
        # find all pagebreaks in the preceding paragraph
        breaks=pneighbors['prev'].findall(pagebreakstring, namespaces)
        if len(breaks) == 1 and not pneighbors['prevtext'].strip():  # we need the strip.. apparently a pb carries some whitespace value
            print "empty pb preceding, deleting"
            # remove pagebreak para
            # log removals before doing them, or the para_id / position refs are unavailable!
            trackEdit(pneighbors['prev'],'remove','removed a page break para preceding section start')
            pneighbors['prev'].getparent().remove(pneighbors['prev'])
        elif len(breaks) > 0 and pneighbors['prevtext'].strip():
            # could remove the last pb anyways, here, consolidate with next case
            # or just remove the text and the pb
            print "don't delete, pb has text!", pneighbors['prevtext']
        elif len(breaks) > 1 and not pneighbors['prevtext'].strip():
            print "more than one pagebreak in the preceding para... will remove the last one"
            # remove last pagebreak para from the preceding paragraph
            breaks[len(breaks)-1].getparent().remove(breaks[len(breaks)-1])
            # log it!:
            trackEdit(pneighbors['prev'],'edit','removed last manual page break from para preceding section start')
        elif len(breaks) == 0:
            print "preceding para is not a pagebreak, skipping delete pb"

# change a style value to a new one
def changeStyle(parastyle,replacestyle,searchstyle,para):
    attrib_style_key = '{%s}val' % wnamespace
    parastyle.set(attrib_style_key, replacestyle)
    trackEdit(para,'edit','changed paragraph style from "%s" to "%s"' % (searchstyle,replacestyle))

# generate ar random id
def generate_id():
    idbase = uuid.uuid4().hex
    idshort = idbase[:8]
    idupper = idshort.upper()

    return str(idupper)

# take a random id and make sure it is unique in the document, otherwise generate a new one, forever
def generate_para_id():
    iduniq = generate_id()
    idsearchstring = './/*w:p[@w14:paraId="%s"]' % iduniq
    while len(root.findall(idsearchstring, namespaces)) > 0:
        print iduniq + " already exists, generating another id"
        iduniq = generate_id()
        idsearchstring = './/*w:p[@w14:paraId="%s"]' % iduniq

    return str(iduniq)

# REVISIT rsids: may or may not be necessary
# see below url; indicates that we should check for uniqueness AND add generated rsids to settings.xml file.
# might be useful for tracking our own edits as part of a single batch of revisions, may not be necessary at all
# http://baxincc.cc/questions/757810/how-to-generate-rsid-attributes-correctly-in-word-docx-files-using-apache-poi
# def generate_rsid(counter):
#     idbase = uuid.uuid4().hex
#     idshort = idbase[:8]
#     iduniq = idshort + str(counter)
#     iduniq = "00" + iduniq[-6:]
#     iduniq = iduniq.upper()
#     return str(iduniq)

# create new paragraph and subelements and insert them prior to seleceted (passed) para
# text content for the inserted para is optional
# Should revisit this using lxml builder
def insertStyledParaBefore(para,searchstyle,insertstyle,contents=''):
    # create new para element
    new_para_id = generate_para_id()
    new_para = etree.Element("{%s}p" % wnamespace)
    new_para.attrib["{%s}paraId" % w14namespace] = new_para_id

    # create new para properties element
    new_para_props = etree.Element("{%s}pPr" % wnamespace)
    new_para_props_style = etree.Element("{%s}pStyle" % wnamespace)
    new_para_props_style.attrib["{%s}val" % wnamespace] = insertstyle

    # append props element to para element
    new_para_props.append(new_para_props_style)
    new_para.append(new_para_props)

    if contents:
        # if text included, create run and text elements, add text, and append to para
        new_para_run = etree.Element("{%s}r" % wnamespace)
        new_para_run_text = etree.Element("{%s}t" % wnamespace)
        new_para_run_text.text = contents
        new_para_run.append(new_para_run_text)
        new_para.append(new_para_run)
        logtext = 'inserted paragraph with style "%s" and text "%s"' % (insertstyle,contents)
    else:
        logtext = 'inserted paragraph with style "%s"' % (insertstyle)

    # append insert new paragraph before the selected para element
    para.addprevious(new_para)

    # log what we did with thte paragraph number
    trackEdit(new_para,'insert',logtext)

# a parent function to search for a para with a certain style, then call two other functions
def deletePageBreakAndInsertPara(searchstyle, insertstyle):
    searchstring = ".//*w:pStyle[@w:val='%s']" % searchstyle
    for parastyle in root.findall(searchstring, namespaces):
        # get parent paragraph (up two levels from a paragraph style, past the pPr)
        para = parastyle.getparent().getparent()

        # delete preceding manual pagebreaks!
        deletePrecedingPageBreak(para)

        # insert a para styled with 'replacestyle' before para with 'searchstyle'
        insertStyledParaBefore(para,searchstyle,insertstyle,"Heading test")

# a parent function to search for a para with a certain style, then call another function
def changeParaStyle(searchstyle, replacestyle):
    searchstring = ".//*w:pStyle[@w:val='%s']" % searchstyle
    for parastyle in root.findall(searchstring, namespaces):
        # get parent paragraph (up two levels from a paragraph style, past the pPr)
        para = parastyle.getparent().getparent()

        # change a paragraph style!
        changeStyle(parastyle,replacestyle,searchstyle,para)

# once all changes havebeen made, call this to add paragraph index numbers to the changelog dicts
def calcParaIndexForEditLog(changelog):
    # changelog is a list of dicts. We are iterating through:
    # for each dict...
    for change in changelog:
        for key in change.keys():
            # for dict key para_id...
            if key == "para_id":
                # seach for the value
                searchstring = ".//*w:p[@w14:paraId='%s']" % change[key]
                para = root.find(searchstring, namespaces)
                # and set value of a new dict item to the paragraph index
                change['final_para_index'] = getParaIndex(para)

# (over)write updated xml back to the file from whence it came
def writeXMLtoFile(file):
    newfile = open(file, 'w')
    with newfile as f:
        f.write(etree.tostring(root, xml_declaration=True, encoding="utf-8", standalone="yes"))
        f.close()

#--------------------- RUN

# test function 1
searchstyle = "TitlepageBookTitletit"
replacestyle = "Section-Titlepagesti"
changeParaStyle(searchstyle, replacestyle)

# test function 2
searchstyle = "Dedicationded"
insertstyle = "Section-Dedicationsde"
deletePageBreakAndInsertPara(searchstyle, insertstyle)

# add final paragraph index numbers to our changelog
calcParaIndexForEditLog(changelog)

# overwrite content to document.xml file
writeXMLtoFile(docxml)

# dump our changelog to json outfile
json.dump(changelog, sys.stdout)
