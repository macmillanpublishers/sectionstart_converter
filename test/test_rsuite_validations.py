# -*- coding: utf-8 -*-
import unittest
# from mock import patch
import sys, os, copy, re, shutil
from lxml import etree, objectify
import logging

# key local paths
mainproject_path = os.path.join(sys.path[0],'xml_docx_stylechecks')
testfiles_basepath = os.path.join(sys.path[0], 'test', 'files_for_test')
tmpdir_basepath = os.path.join(sys.path[0], 'test', 'files_for_test', 'tmp')
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

def createXMLroot(ns_map=cfg.wordnamespaces, root_tag='document'):
    root = etree.Element("{%s}%s" % (cfg.wnamespace, root_tag), nsmap = ns_map)
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

def setupTestFilesinTmp(test_foldername, badxml_srcdir):
    test_tmpdir = os.path.join(tmpdir_basepath, test_foldername)
    # setup test files in tmp
    try:
        shutil.rmtree(test_tmpdir)
    except:
        print "SETUP NOTE: srcdir not yet present in tmp, cannot be deleted"
    shutil.copytree(badxml_srcdir, test_tmpdir)
    return test_tmpdir


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
        # testdocx_root = os.path.join(testfiles_basepath, 'test_stylecount')
        test_folder_root = setupTestFilesinTmp('test_stylecount', os.path.join(testfiles_basepath, 'test_stylecount'))
        unzipDOCX.unzipDOCX('{}{}'.format(test_folder_root,'.docx'), test_folder_root)
        # set paths & run function
        template_styles_xml = os.path.join(self.template_ziproot, 'word', 'styles.xml')
        doc_xml = os.path.join(test_folder_root, 'word', 'document.xml')
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

    def test_checkContainers_goodcases(self):
        testsections = {'Section1':'Section 1', 'Section2':'Section 2'}
        testcontainers = ['Excerpt1','Excerpt2']
        test_ends = ['END','END2']

        # create root and init (Section) para
        root, para = createXML_paraWithRun(list(testsections.keys())[0], '', 'Section1 Head', None)
        # dupe root so we can mock up mutliple outcomes,
        root_dupe = copy.deepcopy(root)
        # append subsequent paras
        root, container_p = createXML_paraWithRun(testcontainers[0], '', 'Excerpt Cntnr', root, 'container_p')
        root, txt_p = createXML_paraWithRun('BodyTextTxt', '', 'I have text', root, 'txt_p')
        root, end_p = createXML_paraWithRun(test_ends[0], '', 'C. End', root, 'end_p')
        # append subsequent paras to dupe
        root_dupe, txt_p2 = createXML_paraWithRun('BodyTextTxt', '', 'I have text', root_dupe, 'txt_p')

        # run our check(s)
        report_dict = rsuite_validations.checkContainers({}, root, testsections, testcontainers, test_ends)
        report_dict2 = rsuite_validations.checkContainers({}, root_dupe, testsections, testcontainers, test_ends)

        #assertions
        self.assertEqual(report_dict, {})   # no containers
        self.assertEqual(report_dict2, {})  # good container

    def test_checkContainers_badcases(self):
        testsections = {'Section1':'Section 1', 'Section2':'Section 2'}
        testcontainers = ['Excerpt1','Excerpt2']
        test_ends = ['END','END2']

        # create root and init (Section) para
        root_noend, para = createXML_paraWithRun(list(testsections.keys())[0], '', 'Section1!', None)
        # append subsequent paras
        root_noend, container_p = createXML_paraWithRun(testcontainers[0], '', 'Excerpt Cntnr', root_noend, 'container_p')
        root_noend, txt_p = createXML_paraWithRun('BodyTextTxt', '', 'I have text', root_noend, 'txt_p')
        # dupe root so we can mock up ideal outcome, by not adding badtxt to dupe, but everything else to both
        root_noend_container = copy.deepcopy(root_noend)
        root_noend_section = copy.deepcopy(root_noend)
        # append subsequent paras to dupes
        root_noend_container, container_p2 = createXML_paraWithRun(testcontainers[1], '', 'Excerpt Cntnr again', root_noend_container, 'container_p2')
        root_noend_container, end_p = createXML_paraWithRun(test_ends[1], '', 'C. End', root_noend_container, 'end_p')
        root_noend_section, section2 = createXML_paraWithRun(list(testsections.keys())[1], '', 'Section2!', root_noend_section, 'section2')

        # run our check(s)
        report_dict_noend = rsuite_validations.checkContainers({}, root_noend, testsections, testcontainers, test_ends)
        report_dict_noend_c = rsuite_validations.checkContainers({}, root_noend_container, testsections, testcontainers, test_ends)
        report_dict_noend_s = rsuite_validations.checkContainers({}, root_noend_section, testsections, testcontainers, test_ends)

        #assertions
        expected_rd = {'container_error': [{'para_id': 'container_p',
                        'description': 'Excerpt1'}]}
        self.assertEqual(report_dict_noend, expected_rd)
        self.assertEqual(report_dict_noend_c, expected_rd)
        self.assertEqual(report_dict_noend_s, expected_rd)

    def test_checkContainers_tableparas(self):
        testsections = {'Section1':'Section 1', 'Section2':'Section 2'}
        testcontainers = ['Excerpt1','Excerpt2']
        test_ends = ['END','END2']

        # create root and init (Section) para
        root, para = createXML_paraWithRun(list(testsections.keys())[0], '', 'Section1!', None)
        tbl1, cntnrpara = createTableWithPara('Start Extract1', testcontainers[0], 'cntnrpara')
        tbl2, endpara = createTableWithPara('Container End', test_ends[0], 'endpara')
        para.addnext(tbl1)
        tbl1.addnext(tbl2)

        # run our check(s)
        report_dict = rsuite_validations.checkContainers({}, root, testsections, testcontainers, test_ends)

        #assertions
        expected_rd = {'illegal_style_in_table': [{'description': 'Excerpt1',
                        'para_id': 'cntnrpara'},
                        {'description': 'END', 'para_id': 'endpara'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_checkContainers_extraends(self):
        testsections = {'Section1':'Section 1', 'Section2':'Section 2'}
        testcontainers = ['Excerpt1','Excerpt2']
        test_ends = ['END','END2']

        # create root and init para
        root_doubleend, end_p1 = createXML_paraWithRun(test_ends[1], '', 'C. End', None, 'end_p1')
        # append subsequent paras
        root_doubleend, sectionp = createXML_paraWithRun(list(testsections.keys())[1], '', 'Section 1!', root_doubleend, 'section1')
        root_doubleend, end_p2 = createXML_paraWithRun(test_ends[0], '', 'C. End again', root_doubleend, 'end_p2')

        # create different root for variation
        root_end_and_sectionend, para = createXML_paraWithRun(list(testsections.keys())[0], '', 'Section 2!', None)
        # append subsequent paras
        root_end_and_sectionend, end_p1b = createXML_paraWithRun(test_ends[0], '', 'C. End', root_end_and_sectionend, 'end_p1b')
        root_end_and_sectionend, container_p = createXML_paraWithRun(testcontainers[1], '', 'Excerpt Cntnr', root_end_and_sectionend, 'container_p1')
        root_end_and_sectionend, end_p2b = createXML_paraWithRun(test_ends[1], '', 'C. End again', root_end_and_sectionend, 'end_p2b')
        root_end_and_sectionend, end_p3 = createXML_paraWithRun(test_ends[0], '', 'C. End again again', root_end_and_sectionend, 'end_p3')

        # run our check(s)
        report_dict_doubleend = rsuite_validations.checkContainers({}, root_doubleend, testsections, testcontainers, test_ends)
        report_dict_end_and_sectionend = rsuite_validations.checkContainers({}, root_end_and_sectionend, testsections, testcontainers, test_ends)

        #assertions
        expected_rd_doubleend = {'container_end_error': [{'description': 'END', 'para_id': 'end_p2'},
                        {'description': 'END2', 'para_id': 'end_p1'}]}
        expected_rd_end_and_sectionend = {'container_end_error': [{'description': 'END', 'para_id': 'end_p1b'},
                        {'description': 'END', 'para_id': 'end_p3'}]}
        self.assertEqual(report_dict_doubleend, expected_rd_doubleend)
        self.assertEqual(report_dict_end_and_sectionend, expected_rd_end_and_sectionend)

    def test_getXMLroot(self):
        # quick test, borrowing another function's xml
        test_tmpdir = setupTestFilesinTmp('test_getXMLroot', os.path.join(testfiles_basepath, "test_updateStyleidInAllXML", 'badxml'))
        doc_xml = os.path.join(test_tmpdir, 'document.xml')
        # run our changes
        good_root = check_docx.getXMLroot(doc_xml)
        bad_root = check_docx.getXMLroot('/nonexistent/path')
        #assertions
        self.assertIsNotNone(good_root)
        self.assertIsNone(bad_root)

    def test_updateStyleidInAllXML(self):
        # test files include badstyles in footnotes and main doc, also badstyles present in tables in both stories
        # NOTE: had to manually change header (1st line) to match quote/caps pattern of re-written xml
        badcharstyle = "boldb0"
        goodcharstyle = "boldb"
        badparastyle = "Body-TextTx0"
        goodparastyle = "Body-TextTx"
        xmls_updated = {'docxml': False, 'footnotes': False, 'endnotes': False}

        # \/ don't need to setup tmpdir here since we are not writing files out \/
        # test_tmpdir = setupTestFilesinTmp('test_updateStyleidInAllXML', os.path.join(testfiles_basepath, "test_updateStyleidInAllXML", 'badxml'))
        badxml_dir = os.path.join(testfiles_basepath, "test_updateStyleidInAllXML", 'badxml')
        doc_root = getRoot(os.path.join(badxml_dir, 'document.xml'))
        fnotes_root = getRoot(os.path.join(badxml_dir, 'footnotes.xml'))
        enotes_root = getRoot(os.path.join(badxml_dir, 'endnotes.xml'))
        styles_root = getRoot(os.path.join(badxml_dir, 'styles.xml'))

        # run our changes
        xmls_updatedcs = check_docx.updateStyleidInAllXML(badcharstyle, goodcharstyle, styles_root, doc_root, None, fnotes_root, xmls_updated)
        xmls_updatedps = check_docx.updateStyleidInAllXML(badparastyle, goodparastyle, styles_root, doc_root, enotes_root, fnotes_root, xmls_updated)

        # expected output
        expected_dir = os.path.join(testfiles_basepath, "test_updateStyleidInAllXML", 'expectedxml')
        expected_docroot = etree.tostring(getRoot(os.path.join(expected_dir, 'document.xml')))
        expected_fnotesroot = etree.tostring(getRoot(os.path.join(expected_dir, 'footnotes.xml')))
        expected_enotesroot = etree.tostring(getRoot(os.path.join(expected_dir, 'endnotes.xml')))
        expected_stylesroot = etree.tostring(getRoot(os.path.join(expected_dir, 'styles.xml')))

        #assertions
        self.assertEqual(xmls_updatedcs, {'docxml': True, 'footnotes': True, 'endnotes': False})
        self.assertEqual(xmls_updatedps, {'docxml': True, 'footnotes': True, 'endnotes': False})
        self.assertEqual(expected_docroot, etree.tostring(doc_root))
        self.assertEqual(expected_fnotesroot, etree.tostring(fnotes_root))
        self.assertEqual(expected_enotesroot, etree.tostring(enotes_root))
        self.assertEqual(expected_stylesroot, etree.tostring(styles_root))


    # test 1 of 5: original bug-case: style-id is fine, but duplicated by rogue random style(s)
    # no change should be made
    def test_verifyStyleIDs_nochange(self):
        macmillanstyle_dict = {"Para Style (PS)": "ParaStylePS", "charstyle (cs)": "charstylecs"}
        root = createXMLroot(cfg.wordnamespaces, 'styles')

        cstyle= createMiscElement('style', cfg.wnamespace, 'styleId', 'charstylecs', cfg.wnamespace)
        cs_name = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle (cs)', cfg.wnamespace)

        pstyle = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStylePS', cfg.wnamespace)
        ps_name = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style (PS)', cfg.wnamespace)

        cstyle2 = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstylecs0', cfg.wnamespace)
        cs2_name = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle cs', cfg.wnamespace)

        pstyle2 = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStylePS0', cfg.wnamespace)
        ps2_name = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style PS', cfg.wnamespace)

        pstyle.append(ps_name)
        cstyle.append(cs_name)
        pstyle2.append(ps2_name)
        cstyle2.append(cs2_name)
        root.append(pstyle)
        root.append(pstyle2)
        root.append(cstyle)
        root.append(cstyle2)
        root_before = copy.deepcopy(root)

        # run our function(s)
        stylenames_updated, xmls_updated = check_docx.verifyStyleIDs(macmillanstyle_dict, {}, root, None, None, None)

        #assertions
        self.assertEqual(xmls_updated, {'docxml': False, 'footnotes': False, 'endnotes': False})
        self.assertEqual(etree.tostring(root), etree.tostring(root_before))
        self.assertFalse(stylenames_updated)

    # test 2 of 5: style is not present (no change)
    def test_verifyStyleIDs_nochange2(self):
        macmillanstyle_dict = {
            "Para Style 2 (PS)": "ParaStyle2PS", "charstyle 2 (cs)": "charstyle2cs",
            "Para Style (PS)": "ParaStylePS", "charstyle (cs)": "charstylecs"}
        root = createXMLroot(cfg.wordnamespaces, 'styles')

        cstyle= createMiscElement('style', cfg.wnamespace, 'styleId', 'charstylecs', cfg.wnamespace)
        cs_name = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle (cs)', cfg.wnamespace)

        pstyle = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStylePS', cfg.wnamespace)
        ps_name = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style (PS)', cfg.wnamespace)

        pstyle2 = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStylePS1', cfg.wnamespace)
        ps2_name = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style PS', cfg.wnamespace)

        pstyle.append(ps_name)
        cstyle.append(cs_name)
        pstyle2.append(ps2_name)
        root.append(pstyle)
        root.append(pstyle2)
        root.append(cstyle)
        root_before = copy.deepcopy(root)

        # run our function(s)
        stylenames_updated, xmls_updated = check_docx.verifyStyleIDs(macmillanstyle_dict, {}, root, None, None, None)

        #assertions
        self.assertEqual(xmls_updated, {'docxml': False, 'footnotes': False, 'endnotes': False})
        self.assertEqual(etree.tostring(root), etree.tostring(root_before))
        self.assertFalse(stylenames_updated)

    # # test 3 of 5: styleid is wrong, but right one is absent (styleid updates to right one)
    def test_verifyStyleIDs_styleidupdate(self):
        test_tmpdir = setupTestFilesinTmp('test_verifyStyleIDs_styleidupdate', os.path.join(testfiles_basepath, "test_verifyStyleIDs_styleidupdate", 'badxml'))
        doc_xml = os.path.join(test_tmpdir, 'document.xml')
        doc_root = getRoot(doc_xml)
        expected_doc_xml = os.path.join(testfiles_basepath, 'test_verifyStyleIDs_styleidupdate', 'expectedxml','document.xml')
        expected_doc_root = getRoot(expected_doc_xml)
        macmillanstyle_dict = {"Para-Style-New (PSN)": "ParaStyleNewPSN", "charstyle-new (cs)": "charstyle-newcs"}
        styles_root = createXMLroot(cfg.wordnamespaces, 'styles')

        # misc. elements
        cstyle= createMiscElement('style', cfg.wnamespace, 'styleId', 'charstylecs', cfg.wnamespace)
        cs_name = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle (cs)', cfg.wnamespace)
        cstyle.append(cs_name)
        styles_root.append(cstyle)
        pstyle2 = createMiscElement('style', cfg.wnamespace, 'styleId', 'Para-Style-NewPSN1', cfg.wnamespace)
        ps2_name = createMiscElement('name', cfg.wnamespace, 'val', 'Para-Style-New PSN', cfg.wnamespace)
        pstyle2.append(ps2_name)
        styles_root.append(pstyle2)

        styles_root_expected = copy.deepcopy(styles_root)

        # elements for bad root
        pstyle_bad = createMiscElement('style', cfg.wnamespace, 'styleId', 'Para-Style-NewPSN0', cfg.wnamespace)
        ps_name_bad = createMiscElement('name', cfg.wnamespace, 'val', 'Para-Style-New (PSN)', cfg.wnamespace)
        pstyle_bad.append(ps_name_bad)
        styles_root.append(pstyle_bad)
        cstyle_bad = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstyle-newcs34', cfg.wnamespace)
        cs_name_bad = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle-new (cs)', cfg.wnamespace)
        cstyle_bad.append(cs_name_bad)
        styles_root.append(cstyle_bad)

        # elements for good root
        pstyle_good = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStyleNewPSN', cfg.wnamespace)
        ps_name_good = createMiscElement('name', cfg.wnamespace, 'val', 'Para-Style-New (PSN)', cfg.wnamespace)
        pstyle_good.append(ps_name_good)
        styles_root_expected.append(pstyle_good)
        cstyle_good = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstyle-newcs', cfg.wnamespace)
        cs_name_good = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle-new (cs)', cfg.wnamespace)
        cstyle_good.append(cs_name_good)
        styles_root_expected.append(cstyle_good)

        # run our function(s)
        stylenames_updated, xmls_updated = check_docx.verifyStyleIDs(macmillanstyle_dict, {}, styles_root, doc_root, None, None)

        #assertions
        self.assertEqual(xmls_updated, {'docxml': True, 'footnotes': False, 'endnotes': False})
        self.assertEqual(etree.tostring(styles_root_expected), etree.tostring(styles_root))
        self.assertTrue(stylenames_updated)
        self.assertEqual(etree.tostring(expected_doc_root), etree.tostring(doc_root))

    # test 4 of 5: is duplicate id of a legacy style - styleid merge
    def test_verifyStyleIDs_legacyduplicate(self):
        test_tmpdir = setupTestFilesinTmp('test_verifyStyleIDs_legacyduplicate', os.path.join(testfiles_basepath, "test_verifyStyleIDs_legacyduplicate", 'badxml'))
        doc_xml = os.path.join(test_tmpdir, 'document.xml')
        doc_root = getRoot(doc_xml)
        expected_doc_xml = os.path.join(testfiles_basepath, 'test_verifyStyleIDs_legacyduplicate', 'expectedxml', 'document.xml')
        expected_doc_root = getRoot(expected_doc_xml)
        macmillanstyle_dict = {"Para Style Change (PSC)": "ParaStyleChangePSC", "charstyle changed (csc)": "charstylechangedcsc"}
        legacystyle_dict = {"char_style_changed (CsC)":[], "Para Style Change (psc)":[]}
        styles_root = createXMLroot(cfg.wordnamespaces, 'styles')
        styles_root_expected = createXMLroot(cfg.wordnamespaces, 'styles')

        # create bad (init) root
        pstyle = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStyleChangePSC0', cfg.wnamespace)
        ps_name = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style Change (PSC)', cfg.wnamespace)
        pstyle.append(ps_name)
        styles_root.append(pstyle)
        pstyle_legacy = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStyleChangepsc', cfg.wnamespace)
        ps_name_legacy = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style Change (psc)', cfg.wnamespace)
        pstyle_legacy.append(ps_name_legacy)
        styles_root.append(pstyle_legacy)
        cstyle = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstylechangedcsc1', cfg.wnamespace)
        cs_name = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle changed (csc)', cfg.wnamespace)
        cstyle.append(cs_name)
        styles_root.append(cstyle)
        cstyle_legacy = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstylechangedCsC', cfg.wnamespace)
        cs_name_legacy = createMiscElement('name', cfg.wnamespace, 'val', 'char_style_changed (CsC)', cfg.wnamespace)
        cstyle_legacy.append(cs_name_legacy)
        styles_root.append(cstyle_legacy)

        # create good (expected outcome) root
        pstyle_good = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStyleChangePSC', cfg.wnamespace)
        ps_name_good = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style Change (PSC)', cfg.wnamespace)
        pstyle_good.append(ps_name_good)
        styles_root_expected.append(pstyle_good)
        cstyle_good = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstylechangedcsc', cfg.wnamespace)
        cs_name_good = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle changed (csc)', cfg.wnamespace)
        cstyle_good.append(cs_name_good)
        styles_root_expected.append(cstyle_good)

        # run our function(s)
        stylenames_updated, xmls_updated = check_docx.verifyStyleIDs(macmillanstyle_dict, legacystyle_dict, styles_root, doc_root, None, None)

        #assertions
        self.assertEqual(xmls_updated, {'docxml': True, 'footnotes': False, 'endnotes': False})
        self.assertEqual(etree.tostring(styles_root_expected), etree.tostring(styles_root))
        self.assertTrue(stylenames_updated)
        self.assertEqual(etree.tostring(expected_doc_root), etree.tostring(doc_root))

    # test 5 of 5: wrong styleid found, right styleid in use, styleid swap takes place
    def test_verifyStyleIDs_nonlegacy_dupe(self):
        test_tmpdir = setupTestFilesinTmp('test_verifyStyleIDs_nonlegacy_dupe', os.path.join(testfiles_basepath, "test_verifyStyleIDs_nonlegacy_dupe", 'badxml'))
        doc_xml = os.path.join(test_tmpdir, 'document.xml')
        doc_root = getRoot(doc_xml)
        expected_doc_xml = os.path.join(testfiles_basepath, 'test_verifyStyleIDs_nonlegacy_dupe', 'expectedxml', 'document.xml')
        expected_doc_root = getRoot(expected_doc_xml)
        macmillanstyle_dict = {"Para Style Switch (PSS)": "ParaStyleSwitchPSS", "charstyle switch (css)": "charstyleswitchcss"}
        legacystyle_dict = {"char_style_switch (Css)":[], "Para Style Switch (Pss)":[]}  # < these are 1 cap away from bad styles, but should not get matched

        styles_root = createXMLroot(cfg.wordnamespaces, 'styles')
        styles_root_expected = createXMLroot(cfg.wordnamespaces, 'styles')

        # create bad (init) root
        pstyle = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStyleSwitchPSS0', cfg.wnamespace)
        ps_name = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style Switch (PSS)', cfg.wnamespace)
        pstyle.append(ps_name)
        styles_root.append(pstyle)
        pstyle_rogue = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStyleSwitchpss', cfg.wnamespace)
        ps_name_rogue = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style Switch (pss)', cfg.wnamespace)
        pstyle_rogue.append(ps_name_rogue)
        styles_root.append(pstyle_rogue)
        cstyle = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstyleswitchcss1', cfg.wnamespace)
        cs_name = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle switch (css)', cfg.wnamespace)
        cstyle.append(cs_name)
        styles_root.append(cstyle)
        cstyle_rogue = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstyleswitchCsS', cfg.wnamespace)
        cs_name_rogue = createMiscElement('name', cfg.wnamespace, 'val', 'char_style_switch (CsS)', cfg.wnamespace)
        cstyle_rogue.append(cs_name_rogue)
        styles_root.append(cstyle_rogue)

        # create good (expected outcome) root
        pstyle_good = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStyleSwitchPSS', cfg.wnamespace)
        ps_name_good = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style Switch (PSS)', cfg.wnamespace)
        pstyle_good.append(ps_name_good)
        styles_root_expected.append(pstyle_good)
        pstyle_bad = createMiscElement('style', cfg.wnamespace, 'styleId', 'ParaStyleSwitchpss0', cfg.wnamespace)
        ps_name_bad = createMiscElement('name', cfg.wnamespace, 'val', 'Para Style Switch (pss)', cfg.wnamespace)
        pstyle_bad.append(ps_name_bad)
        styles_root_expected.append(pstyle_bad)
        cstyle_good = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstyleswitchcss', cfg.wnamespace)
        cs_name_good = createMiscElement('name', cfg.wnamespace, 'val', 'charstyle switch (css)', cfg.wnamespace)
        cstyle_good.append(cs_name_good)
        styles_root_expected.append(cstyle_good)
        cstyle_bad = createMiscElement('style', cfg.wnamespace, 'styleId', 'charstyleswitchCsS1', cfg.wnamespace)
        cs_name_bad = createMiscElement('name', cfg.wnamespace, 'val', 'char_style_switch (CsS)', cfg.wnamespace)
        cstyle_bad.append(cs_name_bad)
        styles_root_expected.append(cstyle_bad)

        # run our function(s)
        stylenames_updated, xmls_updated = check_docx.verifyStyleIDs(macmillanstyle_dict, legacystyle_dict, styles_root, doc_root, None, None)

        #assertions
        self.assertEqual(xmls_updated, {'docxml': True, 'footnotes': False, 'endnotes': False})
        self.assertEqual(etree.tostring(styles_root_expected), etree.tostring(styles_root))
        self.assertTrue(stylenames_updated)
        self.assertEqual(etree.tostring(expected_doc_root), etree.tostring(doc_root))

    def test_checkForDuplicateStyleIDs(self):
        # setup tmpdir, define paths
        test_tmpdir = setupTestFilesinTmp('test_checkForDuplicateStyleIDs', os.path.join(testfiles_basepath, "test_checkForDuplicateStyleIDs", 'badxml'))
        legacystyles_json = os.path.join(test_tmpdir, 'legacy_styles.json')
        macmillanstyles_json = os.path.join(test_tmpdir, 'RSuite.json')
        doc_xml = os.path.join(test_tmpdir, 'document.xml')
        styles_xml = os.path.join(test_tmpdir, 'styles.xml')
        endnotes_xml = os.path.join(test_tmpdir, 'endnotes.xml')
        footnotes_xml = os.path.join(test_tmpdir, 'footnotes.xml')
        expected_output_dir = os.path.join(testfiles_basepath, 'test_checkForDuplicateStyleIDs', 'expectedxml')
        expected_doc_xml = os.path.join(expected_output_dir, 'document.xml')
        expected_styles_xml = os.path.join(expected_output_dir, 'styles.xml')
        expected_endnotes_xml = os.path.join(expected_output_dir, 'endnotes.xml')
        expected_footnotes_xml = os.path.join(expected_output_dir, 'footnotes.xml')

        # run our function(s)
        check_docx.checkForDuplicateStyleIDs(macmillanstyles_json, legacystyles_json, styles_xml, doc_xml, endnotes_xml, footnotes_xml)

        #assertions
        self.assertEqual(etree.tostring(getRoot(styles_xml)), etree.tostring(getRoot(expected_styles_xml)))
        self.assertEqual(etree.tostring(getRoot(doc_xml)), etree.tostring(getRoot(expected_doc_xml)))
        self.assertEqual(etree.tostring(getRoot(endnotes_xml)), etree.tostring(getRoot(expected_endnotes_xml)))
        self.assertEqual(etree.tostring(getRoot(footnotes_xml)), etree.tostring(getRoot(expected_footnotes_xml)))

    def test_flagCustomNoteMarks(self):
        refstyle_dict = {"endnote":cfg.endnote_ref_style, "footnote":cfg.footnote_ref_style}
        doc_xml = os.path.join(testfiles_basepath, 'test_flagCustomNoteMarks', 'document.xml')
        root = getRoot(doc_xml)
        control_docxml = os.path.join(testfiles_basepath, 'test_flagCustomNoteMarks', 'control_doc.xml')
        cntrl_root = getRoot(control_docxml)
        # run our function(s)
        report_dict = rsuite_validations.flagCustomNoteMarks(root, {}, refstyle_dict)
        cntrl_report_dict = rsuite_validations.flagCustomNoteMarks(cntrl_root, {}, refstyle_dict)

        expected_rd = {'custom_endnote_mark':
                [{'description': "custom note marker: '!', endnote id: 2", 'para_id': '4248380B'},
                {'description': "custom note marker: '!!', endnote id: 3", 'para_id': '1CAB8160'},
                {'description': "custom note marker: '111', endnote id: 5", 'para_id': '7BDB93ED'}],
            'custom_footnote_mark':
                [{'description': "custom note marker: '*', footnote id: 2", 'para_id': '7838E8E7'},
                {'description': "custom note marker: '#', footnote id: 4", 'para_id': '739DEFD3'},
                {'description': "custom note marker: '(custom ref-mark is symbol not text)', footnote id: 6", 'para_id': '72A0B3D3'}]}

        #assertions
        self.assertEqual(cntrl_report_dict, {})
        self.assertEqual(report_dict, expected_rd)


if __name__ == '__main__':
    unittest.main()
