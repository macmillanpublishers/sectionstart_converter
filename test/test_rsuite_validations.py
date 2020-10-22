# -*- coding: utf-8 -*-
import unittest
# from mock import patch
import sys, os, copy, re
from lxml import etree, objectify
import logging

# key local paths
mainproject_path = os.path.join(sys.path[0],'xml_docx_stylechecks')
testfiles_basepath = os.path.join(sys.path[0], 'test', 'files_for_test')
rsuite_template_path = os.path.join(sys.path[0], '..', 'RSuite_Word-template', 'StyleTemplate_auto-generate', 'RSuite.dotx')

# append main project path to system path for below imports to work
sys.path.append(mainproject_path)

# import functions for tests below
import xml_docx_stylechecks.lib.doc_prepare as doc_prepare
import xml_docx_stylechecks.lib.rsuite_validations as rsuite_validations
import xml_docx_stylechecks.lib.stylereports as stylereports
import xml_docx_stylechecks.shared_utils.os_utils as os_utils
import xml_docx_stylechecks.cfg as cfg
import xml_docx_stylechecks.shared_utils.lxml_utils as lxml_utils
import xml_docx_stylechecks.shared_utils.check_docx as check_docx
import xml_docx_stylechecks.shared_utils.unzipDOCX as unzipDOCX

# # # # # # Set testing env variable:
os.environ["TEST_FLAG"] = 'true'
### other defs
notes_nsmap = {'w': cfg.wnamespace}

# # # # # # LOCAL FUNCTIONS
# return xml root node of xml file
def getRoot(xmlfile):
    xml_tree = etree.parse(xmlfile)
    xml_root = xml_tree.getroot()
    return xml_root

# this function helps with comparing xmldata that was prettified or manually prepared
def normalizeXML(xmldata):
    # convert passed xml object to string as needed
    if not isinstance(xmldata, basestring):
        xmldata = etree.tostring(xmldata)
    # 'objectify' xml string, which helps normalize xml
    object = objectify.fromstring(xmldata)
    # convert back to string
    xml_string_raw = etree.tostring(object)
    # remove newline chars and their trailing whitespace
    xml_string = re.sub(r'\n\s*', '', xml_string_raw, flags=re.MULTILINE)
    return xml_string

def appendRuntoXMLpara(para, rstylename, runtxt):
    # create run with text
    run = etree.Element("{%s}r" % cfg.wnamespace)
    run_text = etree.Element("{%s}t" % cfg.wnamespace)
    run_text.text = runtxt
    run.append(run_text)
    if rstylename:
        # create new run properties element
        run_props = etree.Element("{%s}rPr" % cfg.wnamespace)
        run_props_style = etree.Element("{%s}rStyle" % cfg.wnamespace)
        run_props_style.attrib["{%s}val" % cfg.wnamespace] = rstylename
        # append props element to run element
        run_props.append(run_props_style)
        run.append(run_props)
    para.append(run)
    return para

def createXMLroot(ns_map=cfg.wordnamespaces):
    root = etree.Element("{%s}document" % cfg.wnamespace, nsmap = ns_map)
    return root

def createMiscElement(element_name, namespace, attribute_name='', attr_val='', attr_namespace=''):
    # have to triple-curly brace curly braces in format string for a single set to be escaped!
    misc_el = etree.Element("{{{}}}{}".format(namespace, element_name))
    # body = etree.Element("{%s}body" % cfg.wnamespace)
    if attribute_name:
        misc_el.attrib["{{{}}}{}".format(attr_namespace, attribute_name)] = attr_val
    return misc_el

def createRun(runtxt, rstylename=''):
    # create run
    run = etree.Element("{%s}r" % cfg.wnamespace)
    if runtxt:
        run_text = etree.Element("{%s}t" % cfg.wnamespace)
        run_text.text = runtxt
        run.append(run_text)
    if rstylename:
        # create new run properties element
        run_props = etree.Element("{%s}rPr" % cfg.wnamespace)
        run_props_style = etree.Element("{%s}rStyle" % cfg.wnamespace)
        run_props_style.attrib["{%s}val" % cfg.wnamespace] = rstylename
        # append props element to run element
        run_props.append(run_props_style)
        run.append(run_props)
    return run

def createPara(para_id, pstylename='', runtxt='', rstylename=''):
    # create para
    new_para = etree.Element("{%s}p" % cfg.wnamespace)
    new_para.attrib["{%s}paraId" % cfg.w14namespace] = para_id
    # if parastyle specified, add it here
    if pstylename:
        # create new para properties element
        new_para_props = etree.Element("{%s}pPr" % cfg.wnamespace)
        new_para_props_style = etree.Element("{%s}pStyle" % cfg.wnamespace)
        new_para_props_style.attrib["{%s}val" % cfg.wnamespace] = pstylename
        # append props element to para element
        new_para_props.append(new_para_props_style)
        new_para.append(new_para_props)
    if runtxt or rstylename:
        run = createRun(runtxt, rstylename='')
        new_para.append(run)
    return new_para

def createTableWithPara(paratxt, parastyle, paraId="p-id"):
    tbl = createMiscElement('tbl', cfg.wnamespace)
    tr = createMiscElement('tr', cfg.wnamespace)
    tc = createMiscElement('tc', cfg.wnamespace)
    para = createPara(paraId, pstylename=parastyle, runtxt=paratxt)
    tc.append(para)
    tr.append(tc)
    tbl.append(tr)
    return tbl, para

# another way to spin up basic xml tree quickly/reproducably without going to file
#   to add more paras run again with root & para_id specified
def createXML_paraWithRun(pstylename, rstylename, runtxt, root=None, para_id='test', ns_map=cfg.wordnamespaces):
    if root is None:
        root = createXMLroot(ns_map)
        body = createMiscElement('body', cfg.wnamespace)
        root.append(body)
    else:
        body = root.find(".//{%s}body" % cfg.wnamespace)
    # create para with run
    para = createPara(para_id, pstylename, runtxt, rstylename)
    body.append(para)
    return root, para

# this function helps with comparing xmldata that was prettified or manually prepared:
# reads xml from file and passes to function above
def normalizeXMLfile(xmlfile):
    with open(xmlfile,'r') as f:
        filecontents = f.read()
    xml_string = normalizeXML(filecontents)
    return xml_string

class Tests(unittest.TestCase):
    def setUp(self):
        # build a basic xml object with lxml.objectify
        myE = objectify.ElementMaker(annotate=False, namespace=cfg.wnamespace, nsmap=cfg.wordnamespaces)
        root = myE.root( myE.body() )
        para_el = objectify.SubElement(root, "{%s}p" % cfg.wnamespace)
        para_el.attrib["{%s}paraId" % cfg.w14namespace] = "test_id"
        self.expected_root = copy.deepcopy(root)
        bad_sub_el = objectify.SubElement(para_el, "unwanted_object")
        self.testroot = root
        # # \/ option to use existing function to help build sxml structure from scratch:
        # lxml_utils.insertPara("Teststyle", existing_para, doc_root, contents, insert_before_or_after)

        ### \/ useful for troubleshooting, when diff-ing xml outputs
        # test_xml = os.path.join(testfiles_basepath, sys._getframe().f_code.co_name, 'expectedxml', 'testing.xml')
        # os_utils.writeXMLtoFile(xml_root, test_xml)

        # unzip current template to files_for_test dir
        self.template_ziproot = os.path.join(testfiles_basepath, 'template_root')
        unzipDOCX.unzipDOCX(rsuite_template_path, self.template_ziproot)
        self.maxDiff = None

    # can `pip mock` to use mock lib, then use "patch" to replace a globally scoped a value for a given test/module, as a decorator
    # @patch('xml_docx_stylechecks.lib.doc_prepare.docroot' = {}), xml_root = {})
    def test_deleteObjects_fromFile(self):#, xml_root):
        # get the bad xml, save a copy for compare
        badxml_root = getRoot(os.path.join(testfiles_basepath, "test_deleteObjects", 'badxml', 'document.xml'))
        # get string of expected xml from known-good file
        expected_xml = os.path.join(testfiles_basepath, "test_deleteObjects", 'expectedxml', 'document.xml')
        # run the function
        report_dict, xml_root = doc_prepare.deleteObjects({}, badxml_root, ['mc:AlternateContent', 'w:drawing'], "shapes")

        # # # ASSERTION:  compare report_dict output, xml_strings, with expected
        self.assertEqual(report_dict, {'deleted_objects-shapes': [{'description': 'deleted shapes of type mc:AlternateContent','para_id': '19599D3D'},{'description': 'deleted shapes of type w:drawing','para_id': '3BED26FA'}]})
        self.assertEqual(normalizeXML(xml_root), normalizeXMLfile(expected_xml))

    def test_deleteObjects_targetNotPresent(self):
        # get string of expected xml from known-good file
        expectedxml_root = getRoot(os.path.join(testfiles_basepath, "test_deleteObjects", 'expectedxml', 'document.xml'))
        # run the function again on file without target object
        report_dict__secondrun, xml_root__secondrun = doc_prepare.deleteObjects({}, expectedxml_root, ['mc:AlternateContent', 'w:drawing'], "shapes")

        # # # ASSERTION:  assert for run on file sans target element (should be no change)
        self.assertEqual(report_dict__secondrun, {})
        self.assertEqual(normalizeXML(xml_root__secondrun), normalizeXML(expectedxml_root))

    def test_deleteObjects_fromNode(self):
        # run the function again, on small, constructed xml object
        report_dict__basic, xml_root__basic = doc_prepare.deleteObjects({}, self.testroot, ['unwanted_object'], "bad_object")

        # # # ASSERTION:  assert for basic run on basic constructed xml object
        self.assertEqual(report_dict__basic, {'deleted_objects-bad_object': [{'description': 'deleted bad_object of type unwanted_object','para_id': 'test_id'}]})
        self.assertEqual(normalizeXML(xml_root__basic), normalizeXML(self.expected_root))

    def test_checkFilename_allbadchars(self):
        filename = "a!@#$''[|]/{ ,.<>\"%^'&}*()-\"\\_=+1"
        badchar_array = check_docx.checkFilename(filename)
        badchars = filename.replace('a', '').replace('1', '').replace('-', '').replace('_', '')
        self.assertEqual(list(badchars), badchar_array)

    def test_checkFilename_nobadchars(self):
        filename = "Just_a_string-978010928909"
        badchar_array = check_docx.checkFilename("Just_a_string-978010928909")
        self.assertEqual([], badchar_array)

    def test_macmillanStyleCount(self):
        # unzip our test docx. Testdoc includes key para-style-types:
        #       - RSuite-styled paras
        #       - built-in (Word) styled paras, including 'Normal (Web)'
        #       - para with old Macmillan-style:
        #       - acceptable built-in styles (e.g. 'Footnote Text')
        #       - styled, unstyled and 'Normal' styled table cells (which should be ignored)
        testdocx_root = os.path.join(testfiles_basepath, 'test_checkdocx', 'stylecount')
        unzipDOCX.unzipDOCX('{}{}'.format(testdocx_root,'.docx'), testdocx_root)
        # set paths & run function
        template_styles_xml = os.path.join(self.template_ziproot, 'word', 'styles.xml')
        doc_xml = os.path.join(testdocx_root, 'word', 'document.xml')
        percent_styled, macmillan_styled_paras, total_paras = check_docx.macmillanStyleCount(doc_xml, template_styles_xml)
        # ASSERTION
        self.assertEqual(total_paras, 8)
        self.assertEqual(macmillan_styled_paras, 5)

    def test_validateImageHolders_badchar_basename(self):
        fullstylename = 'Image-Placement (Img)'
        badfilename = "file,.!@#*<: name-_3.jpg"
        # setup
        root, para = createXML_paraWithRun(fullstylename, '', badfilename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, badfilename)

        expected_rd = {'image_holder_badchar': \
            [{'description': "{}_{}".format(fullstylename, badfilename), \
                'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_validateImageHolders_wrongext(self):
        fullstylename = 'Image-Placement (Img)'
        bad_ext = '.pkg'
        filebasename = "filename-_3"
        badfilename = '{}{}'.format(filebasename, bad_ext)
        # setup
        root, para = createXML_paraWithRun(fullstylename, '', badfilename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, badfilename)

        expected_rd = {'image_holder_ext_error': \
            [{'description': "{}_{}".format(fullstylename, badfilename), \
                'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_validateImageHolders_noext(self):
        fullstylename = 'Image-Placement (Img)'
        badfilename = "filename-_3"
        # setup
        root, para = createXML_paraWithRun(fullstylename, '', badfilename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, badfilename)

        expected_rd = {'image_holder_ext_error': \
            [{'description': "{}_{}".format(fullstylename, badfilename), \
                'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_validateImageHolders_unicodechar(self):
        fullstylename = 'Image-Placement (Img)'
        filename = u'—[(—)-].jpg'
        # setup
        root, para = createXML_paraWithRun(fullstylename, '', filename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, filename)

        expected_rd = {'image_holder_badchar': \
            [{'description': u'Image-Placement (Img)_\u2014[(\u2014)-].jpg', \
                'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_validateImageHolders_inlineholder(self):
        fullstylename = 'cs-image-placement (cimg)'
        badfilename = "file,![ ]@#*<: name-_3.jpg"#"file,.!@#*<: name-_3.jpg"
        # setup
        root, para = createXML_paraWithRun('test', fullstylename, badfilename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, badfilename)

        expected_rd = {'image_holder_badchar': \
            [{'description': "{}_{}".format(fullstylename, badfilename), \
                'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_validateImageHolders_badcharANDnoext(self):
        fullstylename = 'Image-Placement (Img)'
        badfilename = "file,!@#*<: name-_3"
        # setup
        root, para = createXML_paraWithRun(fullstylename, '', badfilename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, badfilename)

        expected_rd = {'image_holder_ext_error': \
            [{'description': "{}_{}".format(fullstylename, badfilename), \
                'para_id': 'test'}], \
            'image_holder_badchar': \
            [{'description': "{}_{}".format(fullstylename, badfilename), \
                'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_validateImageHolders_noproblems(self):
        fullstylename = 'Image-Placement (Img)'
        filename = "filename-_3.jpg"
        # setup
        root, para = createXML_paraWithRun(fullstylename, '', filename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, filename)

        expected_rd = {}
        self.assertEqual(report_dict, expected_rd)

    def test_logTextOfRunsWithStyle(self):
        runstylename = 'test-style'
        run1txt, run2txt, run3txt = "How are ", "you today ", " the end "
        interloper_el = etree.Element("{%s}proofErr" % cfg.wnamespace) #< these occur between runs in real docx
        interloper_el2 = etree.Element("{%s}proofErr" % cfg.wnamespace) #< these occur between runs in real docx
        interloper_el3 = etree.Element("{%s}proofErr" % cfg.wnamespace) #< these occur between runs in real docx

        # setup
        root, para = createXML_paraWithRun("Pteststyle", '', 'leading non sequitur: ')
        para.insert(0, interloper_el)
        para = appendRuntoXMLpara(para, runstylename, run1txt)
        para.append(interloper_el2)
        para = appendRuntoXMLpara(para, runstylename, run2txt)
        para.append(interloper_el3)
        para = appendRuntoXMLpara(para, '', ' , trailing non sequitur.')
        para = appendRuntoXMLpara(para, runstylename, run3txt)
        # run function
        report_dict = stylereports.logTextOfRunsWithStyle({}, root, runstylename, 'demo_report_category')
        expected_rd = {'demo_report_category': [{'description': '{}'.format(run1txt + run2txt), 'para_id': 'test'}, \
        {'description': run3txt, 'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_checkEndnoteFootnoteStyles(self):
        note_sectionname = "Endnotes"
        note_style = cfg.endnotestyle
        unstyled_id = 'p_id2'
        bad_pStyle = "MainHead"
        bad_id = 'p_id3'
        # build xml with 3 types of paras (good, unstyled, and wrong-styled)
        root = createXMLroot()
        endnote = createMiscElement('endnote', cfg.wnamespace, 'id', '7', cfg.wnamespace)
        goodpara = createPara('p_id1', note_style, 'I am a styled para.')
        unstyled_para = createPara(unstyled_id, '', 'Pstyle-less Para', 'demo_runstyle')
        badstyled_para = createPara(bad_id, bad_pStyle, 'Bad-styled para')
        # put items together:
        root.append(endnote)
        endnote.append(goodpara)
        endnote.append(unstyled_para)
        endnote.append(badstyled_para)
        # add a separator type endnote with non-styled child para (should be ignored):
        separator_endnote = createMiscElement('endnote', cfg.wnamespace, 'type', 'continuationSeparator', cfg.wnamespace)
        separator_enpara = createPara('p_id4', '', 'I am separator placeholder para.')
        root.append(separator_endnote)
        separator_endnote.append(separator_enpara)
        # run function
        report_dict = rsuite_validations.checkEndnoteFootnoteStyles(root, {}, note_style, note_sectionname)
        expected_rd = {'improperly_styled_{}'.format(note_sectionname): \
            [{'para_id': unstyled_id, 'description': 'Normal'}, \
            {'para_id': bad_id, 'description': bad_pStyle}]}
        self.assertEqual(report_dict, expected_rd)

    # a duplicate of above test, but with a more accurate nsmap (without w14 ns)
    def test_checkEndnoteFootnoteStyles_noteNsmap(self):
        note_sectionname = "Endnotes"
        note_style = cfg.endnotestyle
        unstyled_id = 'p_id2b'
        bad_pStyle = "MainHead"
        bad_id = 'p_id3b'
        # build xml with 3 types of paras (good, unstyled, and wrong-styled)
        root = createXMLroot(notes_nsmap)
        endnote = createMiscElement('endnote', cfg.wnamespace, 'id', '7', cfg.wnamespace)
        goodpara = createPara('p_id1', note_style, 'I am a styled para.')
        unstyled_para = createPara(unstyled_id, '', 'Pstyle-less Para', 'demo_runstyle')
        badstyled_para = createPara(bad_id, bad_pStyle, 'Bad-styled para')
        # put items together:
        root.append(endnote)
        endnote.append(goodpara)
        endnote.append(unstyled_para)
        endnote.append(badstyled_para)
        # add a separator type endnote with non-styled child para (should be ignored):
        separator_endnote = createMiscElement('endnote', cfg.wnamespace, 'type', 'continuationSeparator', cfg.wnamespace)
        separator_enpara = createPara('p_id4', '', 'I am separator placeholder para.')
        root.append(separator_endnote)
        separator_endnote.append(separator_enpara)
        # run function
        report_dict = rsuite_validations.checkEndnoteFootnoteStyles(root, {}, note_style, note_sectionname)
        expected_rd = {'improperly_styled_{}'.format(note_sectionname): \
            [{'para_id': unstyled_id, 'description': 'Normal'}, \
            {'para_id': bad_id, 'description': bad_pStyle}]}
        self.assertEqual(report_dict, expected_rd)

    def test_handleBlankParasInNotes_multiblanks(self):
        note_name = 'endnote'
        note_stylename = 'EndnoteTxt'
        note_section = 'Endnotes'
        noteref_stylename = 'EndnoteReference'
        para_id = 'p_id-1'

        # build after xml; first build elements
        dummy_note_para = createPara(para_id, note_stylename, '', noteref_stylename)
        noteref_el = createMiscElement('endnoteRef', cfg.wnamespace)
        after_endnote = createMiscElement(note_name, cfg.wnamespace, 'id', '2', cfg.wnamespace)
        after_root = createXMLroot()
        text_run = createRun("[no text]")
        # put everything together
        dummy_note_para[1].append(noteref_el)
        dummy_note_para.append(text_run)
        after_endnote.append(dummy_note_para)
        after_root.append(after_endnote)

        # now build test root
        test_root = createXMLroot()
        test_endnote = createMiscElement(note_name, cfg.wnamespace, 'id', '2', cfg.wnamespace)
        blankpara1 = createPara('bp_id1', note_stylename)
        blankpara2 = createPara('bp_id2', note_stylename)
        # etree.SubElement(test_endnote, blankpara1)
        test_endnote.append(blankpara1)
        test_endnote.append(blankpara2)
        test_root.append(test_endnote)

        # run function
        report_dict = rsuite_validations.handleBlankParasInNotes({}, test_root, note_stylename, noteref_stylename, note_name, note_section)
        expected_rd = {'found_empty_note': [{'description': "endnote",
                        'para_id': 'p_id-1'}],
                        'removed_blank_para': [{'description': 'excess blank para in empty endnote',
                        'para_id': 'bp_id2'}]}
        # assert!
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(etree.tostring(test_root), etree.tostring(after_root))

    def test_handleBlankParasInNotes_mixedblanks(self):
        note_name = 'endnote'
        note_stylename = 'EndnoteTxt'
        note_section = 'Endnotes'
        noteref_stylename = 'EndnoteReference'
        para_id = 'p_id-1'
        para_id2 = 'p_id-2'
        para_id3 = 'p_id-3'
        para_id4 = 'p_id-4'
        # build test xml
        # finish building after_root:
        after_content_p1 = createPara(para_id2, '', "I am a paragraph with content")
        after_content_p2 = createPara(para_id4, note_stylename, "I am too, I have some words")
        after_endnote = createMiscElement(note_name, cfg.wnamespace, 'id', '3', cfg.wnamespace)
        after_root = createXMLroot()
        after_endnote.append(after_content_p1)
        after_endnote.append(after_content_p2)
        after_root.append(after_endnote)
        #   now test root
        test_root = createXMLroot()
        test_endnote = createMiscElement(note_name, cfg.wnamespace, 'id', '3', cfg.wnamespace)
        blankpara1 = createPara(para_id, note_stylename)
        test_content_p1 = copy.deepcopy(after_content_p1)
        blankpara2 = createPara(para_id3)
        test_content_p2 = copy.deepcopy(after_content_p2)
        test_endnote.append(blankpara1)
        test_endnote.append(test_content_p1)
        test_endnote.append(blankpara2)
        test_endnote.append(test_content_p2)
        test_root.append(test_endnote)

        # run function
        report_dict = rsuite_validations.handleBlankParasInNotes({}, test_root, note_stylename, noteref_stylename, note_name, note_section)
        expected_rd = {'removed_blank_para': \
            [{'description': 'blank para in endnote note with other text; note_id: 3', \
            'para_id': para_id}, \
            {'description': 'blank para in endnote note with other text; note_id: 3', \
            'para_id': para_id3}]}
        # assert!
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(etree.tostring(test_root), etree.tostring(after_root))

    def test_handleBlankParasInNotes_blanktablepara(self):
        note_name = 'endnote'
        note_stylename = 'EndnoteTxt'
        note_section = 'Endnotes'
        noteref_stylename = 'EndnoteReference'
        para_id = 'table_p1'

        # build test xml
        empty_note = createMiscElement(note_name, cfg.wnamespace, 'id', '1', cfg.wnamespace)
        test_root = createXMLroot()
        test_root.append(empty_note)
        tbl, blanktblpara = createTableWithPara('', 'BodyTextTxt', para_id)
        empty_note.append(tbl)
        expected_root = copy.deepcopy(test_root)

        # run function
        report_dict = rsuite_validations.handleBlankParasInNotes({}, test_root, note_stylename, noteref_stylename, note_name, note_section)

        # assert!
        expected_rd = {'table_blank_para_notes': [{'description': 'blank para in table cell in endnote',
                        'para_id': para_id}]}
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(etree.tostring(test_root), etree.tostring(expected_root))

    def test_handleBlankParasInNotes_blankparaAndblanktablepara(self):
        note_name = 'endnote'
        note_stylename = 'EndnoteTxt'
        note_section = 'Endnotes'
        noteref_stylename = 'EndnoteReference'
        para_id = 'p_id-1'

        # build test xml
        empty_note = createMiscElement(note_name, cfg.wnamespace, 'id', '1', cfg.wnamespace)
        test_root = createXMLroot()
        test_root.append(empty_note)
        tbl, blanktblpara = createTableWithPara('', 'BodyTextTxt', 'table_p1')
        empty_note.append(tbl)
        expected_root = copy.deepcopy(test_root)
        blankpara = createPara('bp_id1', note_stylename)
        empty_note.append(blankpara)

        # run function
        report_dict = rsuite_validations.handleBlankParasInNotes({}, test_root, note_stylename, noteref_stylename, note_name, note_section)
        expected_rd = {'removed_blank_para': [{'description': 'excess blank para in empty endnote',
                        'para_id': 'bp_id1'}],
                        'table_blank_para_notes': [{'description': 'blank para in table cell in endnote',
                        'para_id': 'table_p1'}]}
        # assert!
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(etree.tostring(test_root), etree.tostring(expected_root))

    def test_handleBlankParasInNotes_noparas_footnotes(self):
        note_name = 'footnote'
        note_stylename = 'FootnoteTxt'
        noteref_stylename = 'FootnoteReference'
        note_section = 'Footnotes'
        para_id = 'p_id-1'

        # build after xml; first build elements
        dummy_note_para = createPara(para_id, note_stylename, '', noteref_stylename)
        noteref_el = createMiscElement('endnoteRef', cfg.wnamespace)
        after_endnote = createMiscElement(note_name, cfg.wnamespace, 'id', '1', cfg.wnamespace)
        after_root = createXMLroot()
        text_run = createRun("[no text]")
        # put everything together
        dummy_note_para[1].append(noteref_el)
        dummy_note_para.append(text_run)
        after_endnote.append(dummy_note_para)
        after_root.append(after_endnote)

        # build test xml
        empty_note = createMiscElement(note_name, cfg.wnamespace, 'id', '1', cfg.wnamespace)
        test_root = createXMLroot()
        test_root.append(empty_note)

        # run function
        report_dict = rsuite_validations.handleBlankParasInNotes({}, test_root, note_stylename, noteref_stylename, note_name, note_section)
        expected_rd = {'found_empty_note': \
            [{'description': 'footnote', \
            'para_id': para_id}]}
        # assert!
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(etree.tostring(test_root), etree.tostring(after_root))

    def test_handleBlankParasInNotes_separator(self):
        note_name = 'endnote'
        # build test xml: cseparator
        root = createXMLroot()
        cseparator_note = createMiscElement(note_name, cfg.wnamespace, 'type', 'continuationSeparator', cfg.wnamespace)
        cseparator_note.attrib["{{{}}}{}".format(cfg.wnamespace, 'id')] = '0'
        cseparator_para = createPara('p_id-1')
        cseparator_run = createRun('')
        cseparator = createMiscElement('continuationSeparator', cfg.wnamespace)
        cseparator_run.append(cseparator)
        cseparator_para.append(cseparator_run)
        cseparator_note.append(cseparator_para)
        root.append(cseparator_note)
        # build separator:
        separator_note = createMiscElement(note_name, cfg.wnamespace, 'type', 'separator', cfg.wnamespace)
        separator_note.attrib["{{{}}}{}".format(cfg.wnamespace, 'id')] = '0'
        separator_para = createPara('p_id-1')
        separator_run = createRun('')
        separator = createMiscElement('separator', cfg.wnamespace)
        separator_run.append(separator)
        separator_para.append(separator_run)
        separator_note.append(separator_para)
        root.append(separator_note)
        expected_root = copy.deepcopy(root)
        # run the function
        report_dict = rsuite_validations.handleBlankParasInNotes({}, root, 'note_stylename', 'noteref_stylename', note_name, 'note_section')
        # assert!
        self.assertEqual(report_dict, {})
        self.assertEqual(etree.tostring(expected_root), etree.tostring(root))

    # testing targeting at new functionality: pulling info from footnotes/endnotes_xml
    # note: this is a little more of an integration test, b/c a more convoluted function
    def test_calcLocationInfoForLog(self):
        logging.basicConfig(level=logging.ERROR)
        # ^ without this we get a notice about logger handlers being unavailable.
        #   Can change loglevel to see mssges
        sectionnames = ["Section-Test (TEST)"]
        # create main root mockup with no preceding section para, and one with section preceding
        root, para = createXML_paraWithRun('BodyTextTxt', '', "I'm a para with no section", None, 'p_id0')
        root, para = createXML_paraWithRun(sectionnames[0], '', "Section Heading", root, 'p_id1')
        root, para = createXML_paraWithRun('BodyTextTxt', '', "I'm a normal para", root, 'p_id2')
        # create alt root mockup with endnote, child para
        altroot = createXMLroot()
        endnote = createMiscElement('endnote', cfg.wnamespace, 'id', '1', cfg.wnamespace)
        enpara = createPara('en_p_id1', 'EndnoteText', 'I am endnote txt.')
        altroot.append(endnote)
        endnote.append(enpara)
        # create initial report_dict. Adding a paragraph ref, 'p_id3',which does not exist,
        #   since some paras get deleted prior to this function
        report_dict = {'test_category': \
            [{'para_id': 'p_id0', 'description': 'mainxml'}, \
            {'para_id': 'p_id2', 'description': 'mainxml'}, \
            {'para_id': 'en_p_id1', 'description': 'altxml'}, \
            {'para_id': 'p_id3', 'description': 'mainxml'}]}
        # run function
        report_dict = lxml_utils.calcLocationInfoForLog(report_dict, root, sectionnames, {'Endnotes':altroot})
        #assertion
        expected_rd =  {'test_category': [{'description': 'mainxml',
                'para_id': 'p_id0',
                'para_index': 0,
                'para_string': "I'm a para with no section",
                'parent_section_start_content': '',
                'parent_section_start_type': 'n-a'},
            {'description': 'mainxml',
                'para_id': 'p_id2',
                'para_index': 2,
                'para_string': "I'm a normal para",
                'parent_section_start_content': 'Section Heading',
                'parent_section_start_type': 'Section-Test (TEST)'},
            {'description': 'altxml',
                'note-or-comment_id': '1',
                'para_id': 'en_p_id1',
                'para_index': 'n-a',
                'para_string': 'I am endnote txt.',
                'parent_section_start_content': 'n-a',
                'parent_section_start_type': 'Endnotes'},
            {'description': 'mainxml',
                'para_id': 'p_id3',
                'para_index': 'n-a',
                'para_string': 'n-a',
                'parent_section_start_content': '',
                'parent_section_start_type': 'n-a'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_calcLocationInfoForLog_tablepara(self):
        logging.basicConfig(level=logging.ERROR)
        # ^ without this we get a notice about logger handlers being unavailable.
        #   Can change loglevel to see mssges
        sectionnames = ["Section-Test (TEST)"]
        para_id = 'table_p_id1'
        # create alt root mockup with endnote, child para
        root, para = createXML_paraWithRun(sectionnames[0], '', "Section Heading", None, 'p_id1')
        table = createMiscElement('table', cfg.wnamespace)
        tr = createMiscElement('tr', cfg.wnamespace)
        tc = createMiscElement('tc', cfg.wnamespace)
        para = createPara(para_id, 'BodyTextTxt', 'I am a para in a table')
        tc.append(para)
        tr.append(tc)
        table.append(tr)
        root.append(table)

        # create initial report_dict. Adding a paragraph ref, 'p_id3',which does not exist,
        #   since some paras get deleted prior to this function
        report_dict = {'test_category': \
            [{'para_id': para_id, 'description': 'tablecellpara'}]}

        # run function
        report_dict = lxml_utils.calcLocationInfoForLog(report_dict, root, sectionnames)
        #assertion
        expected_rd =  {'test_category': [{'description': 'tablecellpara',
                'para_id': para_id,
                'para_index': 'tablecell_para',
                'para_string': "I am a para in a table",
                'parent_section_start_content': 'Section Heading',
                'parent_section_start_type': 'Section-Test (TEST)'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_checkNamespace(self):
        test_nsmap = {'w': cfg.wnamespace, 'tst': 'test_ns_value', 'w14': 'diff_ns_value'}
        good_ns = 'w'       # defined in our own wordnamespaces, and in target xml_root
        bad_ns = 'bad'     # not defined either place
        # good_ns2 = 'w14'    # defined both places, but different ns values? Do we want this or want to prevent?
        root = createXMLroot(test_nsmap)

        # run tests
        bool_good = lxml_utils.checkNamespace(root, good_ns)
        bool_bad = lxml_utils.checkNamespace(root, bad_ns)
        # bool_good2 = lxml_utils.checkNamespace(root, good_ns2)

        #assertions
        self.assertEqual(bool_good, True)
        self.assertEqual(bool_bad, False)
        # self.assertEqual(bool_good2, True)

    def test_addNamespace(self):
        # setup intitial nsmap, ns we are adding, and expected-final nsmap
        test_nsmap = {'w': cfg.wnamespace, 'tst': 'test_ns_value'}#, 'w14': 'old'}
        final_nsmap = test_nsmap.copy()
        new_nsprefix = 'w14'
        new_nsuri = cfg.w14namespace
        new_nsmap = {new_nsprefix: new_nsuri}
        final_nsmap[new_nsprefix] = new_nsuri

        # create dummy root
        root = createXMLroot(test_nsmap)
        # run our transform
        lxml_utils.addNamespace(root, new_nsprefix, new_nsuri)
        #assertions
        self.assertEqual(root.nsmap, final_nsmap)

    def test_verifyOrAddNamespace(self):
        # setup intitial nsmap, ns we are adding, and expected-final nsmap
        test_nsmap_good = {'w': cfg.wnamespace, 'tst': 'test_ns_value', 'w14': cfg.w14namespace}
        test_nsmap_bad = {'w': cfg.wnamespace, 'tst': 'test_ns_value'}
        new_nsprefix = 'w14'
        new_nsuri = cfg.w14namespace

        # create dummy roots
        root_good, para = createXML_paraWithRun('BodyTextTxt', '', "I'm a normal para", None, 'p_id2', test_nsmap_good)
        root_bad, para2 = createXML_paraWithRun('BodyTextTxt', '', "I'm a normal para", None, 'p_id2', test_nsmap_bad)
        root_good_before = copy.deepcopy(root_good)
        root_bad_before = copy.deepcopy(root_bad)

        # run our transform
        lxml_utils.verifyOrAddNamespace(root_bad, new_nsprefix, new_nsuri)
        lxml_utils.verifyOrAddNamespace(root_good, new_nsprefix, new_nsuri)

        #assertions
        self.assertEqual(root_good.nsmap, root_good_before.nsmap)
        self.assertEqual(etree.tostring(root_good), etree.tostring(root_good_before))
        self.assertEqual(root_good.nsmap, root_bad.nsmap)
        self.assertEqual(etree.tostring(root_good), etree.tostring(root_bad))
        self.assertNotEqual(root_bad_before.nsmap, root_bad.nsmap)
        self.assertNotEqual(etree.tostring(root_bad_before), etree.tostring(root_bad))

    def test_handleBlankParasInTables_soloblankpara(self):
        # create root with soloblanktablepara, table
        root, para = createXML_paraWithRun('BodyTextTxt', '', '', None)
        tbl, tblpara = createTableWithPara('', 'BodyTextTxt', 'p-id')
        para.addnext(tbl)
        before_root = copy.deepcopy(root)

        # run our transform
        report_dict, solopara_bool = rsuite_validations.handleBlankParasInTables({}, root, tblpara)

        #assertions
        expected_rd = {'table_blank_para': [{'description': 'blank para found in table cell',
                        'para_id': 'p-id'}]}
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(solopara_bool, True)

    ## test for setting: "removing_excess_tbl_blankparas" to "True" in "handleBlankParasInTables"
    # def test_handleBlankParasInTables_multiblankparas(self):
    #     # create root with blank para preceding other para
    #     root, para = createXML_paraWithRun('BodyTextTxt', '', '', None)
    #     tbl, tblpara = createTableWithPara('A', 'BodyTextTxt', 'p-id')
    #     xtrapara = createPara('p-id2', '', 'I have text.')
    #     tblpara.addnext(xtrapara)
    #     para.addnext(tbl)
    #     before_root = copy.deepcopy(root)
    #
    #     # run our transform
    #     report_dict, solopara_bool = rsuite_validations.handleBlankParasInTables({}, root, tblpara)
    #
    #     #assertions
    #     expected_rd = {}
    #     self.assertEqual(report_dict, expected_rd)
    #     self.assertEqual(solopara_bool, False)

    # this blank para should be removed and reported
    def test_removeBlankParas_blanktxtpara(self):
        testsections = {'Section1':'Section 1', 'Section2':'Section 2'}
        testcontainers = ['Excerpt1','Excerpt2']
        test_ends = ['END','END2']
        test_breaks = ['break1','break2']
        # create root and init (Section) para
        root, para = createXML_paraWithRun(list(testsections.keys())[0], '', 'Section1!', None)
        # append subsequent paras
        root, goodcontainer_p = createXML_paraWithRun(testcontainers[0], '', 'Container1Starter', root, 'goodcontainer_p')
        root, goodbreak_p = createXML_paraWithRun(test_breaks[1], '', '-break text-', root, 'goodbreak_p')
        # dupe root so we can mock up ideal outcome, by not adding badtxt to dupe, but everything else to both
        root_dupe = copy.deepcopy(root)
        root, badtxt_p = createXML_paraWithRun('BodyTextTxt', '', '   ', root, 'badtxt_p')
        root, goodend_p = createXML_paraWithRun(test_ends[0], '', 'C. End', root, 'goodend_p')
        root_dupe, goodend_p = createXML_paraWithRun(test_ends[0], '', 'C. End', root_dupe, 'goodend_p')

        # run our transform
        report_dict = rsuite_validations.removeBlankParas({}, root, testsections, testcontainers, test_ends, test_breaks)

        #assertions
        expected_rd = {'removed_blank_para': [{'description': 'removed BodyTextTxt-styled para',
                        'para_id': 'badtxt_p'}]}
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(etree.tostring(root), etree.tostring(root_dupe))

    # This blank container para should be removed, and reported in two separate categories
    def test_removeBlankParas_blankcontainerpara(self):
        testsections = {'Section1':'Section 1', 'Section2':'Section 2'}
        testcontainers = ['Excerpt1','Excerpt2']
        test_ends = ['END','END2']
        test_breaks = ['break1','break2']
        # create root and init (Section) para
        root, para = createXML_paraWithRun(list(testsections.keys())[0], '', 'Section1!', None)
        # append subsequent paras
        # dupe root so we can mock up ideal outcome, by not adding badtxt to dupe, but everything else to both
        root_dupe = copy.deepcopy(root)
        root, badcontainer_p = createXML_paraWithRun(testcontainers[0], '', '', root, 'badcontainer_p')
        root, goodbreak_p = createXML_paraWithRun(test_breaks[1], '', '-break text-', root, 'goodbreak_p')
        root, goodtxt_p = createXML_paraWithRun('BodyTextTxt', '', 'I have text', root, 'badtxt_p')
        root, goodend_p = createXML_paraWithRun(test_ends[0], '', 'C. End', root, 'goodend_p')
        root_dupe, goodbreak_p = createXML_paraWithRun(test_breaks[1], '', '-break text-', root, 'goodbreak_p')
        root_dupe, goodtxt_p = createXML_paraWithRun('BodyTextTxt', '', 'I have text', root, 'goodtxt_p')
        root_dupe, goodend_p = createXML_paraWithRun(test_ends[0], '', 'C. End', root_dupe, 'goodend_p')
        # with patch("xml_docx_stylechecks.shared_utils.lxml_utils.getStyleLongname") as getlongname_mock:
        #     getlongname_mock.return_value = "foo"
        # test = lxml_utils.getStyleLongname('sdsd')

        # run our transform
        report_dict = rsuite_validations.removeBlankParas({}, root, testsections, testcontainers, test_ends, test_breaks)

        #assertions
        expected_rd = {'removed_blank_para': [{'description': 'removed Excerpt1-styled para',
                        'para_id': 'badcontainer_p'}],
                        'removed_container_blank_para': [{'description': 'Excerpt1_\'Section2: "Section1!"\'',
                        'para_id': 'badcontainer_p'}]}
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(etree.tostring(root), etree.tostring(root_dupe))

    # This blank section & container paras should be removed, and reported in two separate categories
    #   Also notable, with no parent section; container description should contain default value
    def test_removeBlankParas_blankSectionAndContainerParas(self):
        testsections = {'Section1':'Section 1', 'Section2':'Section 2'}
        testcontainers = ['Excerpt1','Excerpt2']
        test_ends = ['END','END2']
        test_breaks = ['break1','break2']
        # create root and init (Section) para
        root, para = createXML_paraWithRun(list(testsections.keys())[0], '', '', None)
        # append subsequent paras
        # dupe root so we can mock it up as ideal outcome, then run function on root_dupe
        root_dupe = copy.deepcopy(root)
        root_dupe, badcontainer_p = createXML_paraWithRun(testcontainers[0], '', '', root_dupe, 'badcontainer_p')
        root_dupe, goodbreak_p = createXML_paraWithRun(test_breaks[1], '', '-break text-', root_dupe, 'goodbreak_p')
        root_dupe, goodtxt_p = createXML_paraWithRun('BodyTextTxt', '', 'I have text', root_dupe, 'goodtxt_p')
        root_dupe, goodend_p = createXML_paraWithRun(test_ends[0], '', 'C. End', root_dupe, 'goodend_p')
        # mock up root as expected outcome
        para.getparent().remove(para)
        root, goodbreak_p = createXML_paraWithRun(test_breaks[1], '', '-break text-', root, 'goodbreak_p')
        root, goodtxt_p = createXML_paraWithRun('BodyTextTxt', '', 'I have text', root, 'goodtxt_p')
        root, goodend_p = createXML_paraWithRun(test_ends[0], '', 'C. End', root, 'goodend_p')

        # run our transform
        report_dict = rsuite_validations.removeBlankParas({}, root_dupe, testsections, testcontainers, test_ends, test_breaks)

        #assertions
        expected_rd = {'removed_blank_para': [{'description': 'removed Section2-styled para',
                        'para_id': 'test'},
                        {'description': 'removed Excerpt1-styled para',
                        'para_id': 'badcontainer_p'}],
                        'removed_container_blank_para': [{'description': 'Excerpt1_\'n-a: ""\'',
                        'para_id': 'badcontainer_p'}],
                        'removed_section_blank_para': [{'description': 'Section2',
                        'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(etree.tostring(root), etree.tostring(root_dupe))

    # adding two blank spacebreak paras, including one in a table
    # the one in the table should be reported, not removed; the other should be reported in 2 separate categories
    def test_removeBlankParas_blankBreakAndTableParas(self):
        testsections = {'Section1':'Section 1', 'Section2':'Section 2'}
        testcontainers = ['Excerpt1','Excerpt2']
        test_ends = ['END','END2']
        test_breaks = ['break1','break2']
        # create root and init (Section) para
        root, para = createXML_paraWithRun(list(testsections.keys())[0], '', 'Section1!', None)
        # append subsequent paras
        root, goodcontainer_p = createXML_paraWithRun(testcontainers[0], '', 'Container1Starter', root, 'goodcontainer_p')
        root, goodtxt_p = createXML_paraWithRun('BodyTextTxt', '', 'Im a para with text', root, 'goodtxt_p')
        # dupe root so we can mock up ideal outcome, by not adding badtxt to dupe, but everything else to both
        root_dupe = copy.deepcopy(root)
        root, badbreak1_p = createXML_paraWithRun(test_breaks[1], '', '', root, 'badbreak1_p')
        root, goodend_p = createXML_paraWithRun(test_ends[0], '', 'C. End', root, 'goodend_p')
        tbl, badbreak2_p = createTableWithPara('', test_breaks[0], 'badbreak2_p')
        goodend_p.addnext(tbl)
        # now add only items that should not be removed, to the root_dupe
        root_dupe, goodend_p_dupe = createXML_paraWithRun(test_ends[0], '', 'C. End', root_dupe, 'goodend_p')
        tbl2, badbreak2_p_dupe = createTableWithPara('', test_breaks[0], 'badbreak2_p')
        goodend_p_dupe.addnext(tbl2)

        # run our transform
        report_dict = rsuite_validations.removeBlankParas({}, root, testsections, testcontainers, test_ends, test_breaks)

        # assertions
        expected_rd = {'removed_blank_para': [{'description': 'removed break2-styled para',
                        'para_id': 'badbreak1_p'}],
                        'removed_spacebreak_blank_para': [{'description': 'break2_\'Section2: "Section1!"\'',
                        'para_id': 'badbreak1_p'}],
                        'table_blank_para': [{'description': 'blank para found in table cell',
                        'para_id': 'badbreak2_p'}]}
        self.assertEqual(report_dict, expected_rd)
        self.assertEqual(etree.tostring(root), etree.tostring(root_dupe))

if __name__ == '__main__':
    unittest.main()
