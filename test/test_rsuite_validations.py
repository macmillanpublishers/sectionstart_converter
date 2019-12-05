import unittest
# from mock import patch
import sys, os, copy, re
from lxml import etree, objectify

from xml_docx_stylechecks.basicfunction import BasicFunction

# key local paths
mainproject_path = os.path.join(sys.path[0],'xml_docx_stylechecks')
testfiles_basepath = os.path.join(sys.path[0], 'test', 'files_for_test')

# append main project path to system path for below imports to work
sys.path.append(mainproject_path)

# import functions for tests below
import xml_docx_stylechecks.lib.doc_prepare as doc_prepare
import xml_docx_stylechecks.shared_utils.os_utils as os_utils
import xml_docx_stylechecks.cfg as cfg
import xml_docx_stylechecks.shared_utils.lxml_utils as lxml_utils

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

# this function helps with comparing xmldata that was prettified or manually prepared:
# reads xml from file and passes to function above
def normalizeXMLfile(xmlfile):
    with open(xmlfile,'r') as f:
        filecontents = f.read()
    xml_string = normalizeXML(filecontents)
    return xml_string

class TestBasicFunction(unittest.TestCase):
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

    # can `pip mock` to use mock lib, then use "patch" to replace a globally scoped a value for a given test/module, as a decorator
    # @patch('xml_docx_stylechecks.lib.doc_prepare.docroot' = {}), xml_root = {})
    def test_deleteObjects(self):#, xml_root):
        # get the bad xml, save a copy for compare
        bad_xml = os.path.join(testfiles_basepath, sys._getframe().f_code.co_name, 'badxml', 'document.xml')
        badxml_tree = etree.parse(bad_xml)
        badxml_root = badxml_tree.getroot()

        # run the function
        report_dict, xml_root = doc_prepare.deleteObjects({}, badxml_root, ['mc:AlternateContent', 'w:drawing'], "shapes")

        ### \/ useful for troubleshooting, when diff-ing xml outputs
        # test_xml = os.path.join(testfiles_basepath, sys._getframe().f_code.co_name, 'expectedxml', 'testing.xml')
        # os_utils.writeXMLtoFile(xml_root, test_xml)

        # run the function again, on output from last run
        xml_root_copy = copy.deepcopy(xml_root)
        report_dict__secondrun, xml_root__secondrun = doc_prepare.deleteObjects({}, xml_root_copy, ['mc:AlternateContent', 'w:drawing'], "shapes")

        # run the function again, on small, constructed xml object
        report_dict__basic, xml_root__basic = doc_prepare.deleteObjects({}, self.testroot, ['unwanted_object'], "bad_object")

        # get string of expected xml from known-good file
        expected_xml = os.path.join(testfiles_basepath, sys._getframe().f_code.co_name, 'expectedxml', 'document.xml')

        # # # ASSERTIONS
        # compare report_dict output, xml_strings, with expected
        self.assertEqual(report_dict, {'deleted_objects-shapes': [{'description': 'deleted shapes of type mc:AlternateContent','para_id': '19599D3D'},{'description': 'deleted shapes of type w:drawing','para_id': '3BED26FA'}]})
        self.assertEqual(normalizeXML(xml_root), normalizeXMLfile(expected_xml))
        # assert for second run on same/edited file (should be no change)
        self.assertEqual(report_dict__secondrun, {})
        self.assertEqual(normalizeXML(xml_root__secondrun), normalizeXML(xml_root))
        # assert for basic run on construvted basic xml object
        self.assertEqual(report_dict__basic, {'deleted_objects-bad_object': [{'description': 'deleted bad_object of type unwanted_object','para_id': 'test_id'}]})
        self.assertEqual(normalizeXML(self.testroot), normalizeXML(self.expected_root))

if __name__ == '__main__':
    unittest.main()

# note for future development: have test.docx with static name ('test.docx'), and unzip it on the fly and import document.xml, to save a step on re-building testdoc    
