import unittest
# from mock import patch
import sys, os, copy, re
from lxml import etree, objectify

# key local paths
mainproject_path = os.path.join(sys.path[0],'xml_docx_stylechecks')
testfiles_basepath = os.path.join(sys.path[0], 'test', 'files_for_test')
rsuite_template_path = os.path.join(sys.path[0], '..', 'RSuite_Word-template', 'StyleTemplate_auto-generate', 'RSuite.dotx')

# append main project path to system path for below imports to work
sys.path.append(mainproject_path)

# import functions for tests below
import xml_docx_stylechecks.lib.doc_prepare as doc_prepare
import xml_docx_stylechecks.lib.stylereports as stylereports
import xml_docx_stylechecks.shared_utils.os_utils as os_utils
import xml_docx_stylechecks.cfg as cfg
import xml_docx_stylechecks.shared_utils.lxml_utils as lxml_utils
import xml_docx_stylechecks.shared_utils.check_docx as check_docx
import xml_docx_stylechecks.shared_utils.unzipDOCX as unzipDOCX


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

# another way to spin up basic xml tree quickly/reproducably without going to file
def createXMLparaWithRun(pstylename, rstylename, runtxt):
    root = etree.Element("{%s}document" % cfg.wnamespace, nsmap = cfg.wordnamespaces)
    body = etree.Element("{%s}body" % cfg.wnamespace)
    root.append(body)
    # create para
    new_para = etree.Element("{%s}p" % cfg.wnamespace)
    new_para.attrib["{%s}paraId" % cfg.w14namespace] = "test"
    # create new para properties element
    new_para_props = etree.Element("{%s}pPr" % cfg.wnamespace)
    new_para_props_style = etree.Element("{%s}pStyle" % cfg.wnamespace)
    new_para_props_style.attrib["{%s}val" % cfg.wnamespace] = pstylename
    # append props element to para element
    new_para_props.append(new_para_props_style)
    new_para.append(new_para_props)
    # create run
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
    new_para.append(run)
    body.append(new_para)

    return root, new_para

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
        root, para = createXMLparaWithRun(fullstylename, '', badfilename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, badfilename)

        expected_rd = {'image_holder_badchar': \
            [{'description': "{}_{}".format(fullstylename, badfilename), \
                'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_validateImageHolders_wrongext(self):
        fullstylename = 'Image-Placement (Img)'
        bad_ext = '.png'
        filebasename = "filename-_3"
        badfilename = '{}{}'.format(filebasename, bad_ext)
        # setup
        root, para = createXMLparaWithRun(fullstylename, '', badfilename)
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
        root, para = createXMLparaWithRun(fullstylename, '', badfilename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, badfilename)

        expected_rd = {'image_holder_ext_error': \
            [{'description': "{}_{}".format(fullstylename, badfilename), \
                'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

    def test_validateImageHolders_inlineholder(self):
        fullstylename = 'cs-image-placement (cimg)'
        badfilename = "file,![ ]@#*<: name-_3.jpg"#"file,.!@#*<: name-_3.jpg"
        # setup
        root, para = createXMLparaWithRun('test', fullstylename, badfilename)
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
        root, para = createXMLparaWithRun(fullstylename, '', badfilename)
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
        root, para = createXMLparaWithRun(fullstylename, '', filename)
        # run function
        report_dict = stylereports.validateImageHolders({}, root, fullstylename, para, filename)

        expected_rd = {}
        self.assertEqual(report_dict, expected_rd)

    def test_logTextOfRunsWithStyle(self):
        runstylename = 'test-style'
        run1txt, run2txt, run3txt = "How are ", "you today ", " the end "
        # setup
        root, para = createXMLparaWithRun("Pteststyle", '', 'leading non sequitur: ')
        para = appendRuntoXMLpara(para, runstylename, run1txt)
        para = appendRuntoXMLpara(para, runstylename, run2txt)
        para = appendRuntoXMLpara(para, '', ' , trailing non sequitur.')
        para = appendRuntoXMLpara(para, runstylename, run3txt)
        # run function
        report_dict = stylereports.logTextOfRunsWithStyle({}, root, runstylename, 'demo_report_category')
        expected_rd = {'demo_report_category': [{'description': 'How are you today '.format(run1txt + run2txt), 'para_id': 'test'}, \
        {'description': run3txt, 'para_id': 'test'}]}
        self.assertEqual(report_dict, expected_rd)

if __name__ == '__main__':
    unittest.main()
