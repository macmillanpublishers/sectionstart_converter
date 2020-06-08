# pip install xmldiff #!<======
import sys, os, copy, re
import logging
import shutil
import subprocess
import difflib
from lxml import etree
from xmldiff import main, formatting

if len(sys.argv) == 2:
    if sys.argv[1] == 'update_valid_outputs':
        update_valid_outputs = True
        single_infile = sys.argv[1]
    elif sys.argv[1]:
        update_valid_outputs = False
        single_infile = ''
else:
    update_valid_outputs = False
    single_infile = ''

# key local paths
rsv_main_path = os.path.join(sys.path[0], '..', 'xml_docx_stylechecks', 'rsuitevalidate_main.py')
transform_testfiles_dir = os.path.join(sys.path[0], 'files_for_test', 'full_transform', 'test_docx_files')
validfiles_basedir = os.path.join(sys.path[0], 'files_for_test', 'full_transform', 'validated_output')
tmpdir_base = os.path.join(sys.path[0], 'files_for_test', 'tmp')
debug_diff_only = True  # < - for testing diff parameters
# rsuite_template_path = os.path.join(sys.path[0], '..', 'RSuite_Word-template', 'StyleTemplate_auto-generate', 'RSuite.dotx')

# append main project path to system path for below imports to work
# sys.path.append(mainproject_path)
#
# print 'mainproject_path', mainproject_path
#
# # import functions for tests below
# import xml_docx_stylechecks.lib.doc_prepare as doc_prepare
# import xml_docx_stylechecks.lib.rsuite_validations as rsuite_validations
# import xml_docx_stylechecks.lib.stylereports as stylereports
# import xml_docx_stylechecks.shared_utils.os_utils as os_utils
# import xml_docx_stylechecks.cfg as cfg
# import xml_docx_stylechecks.shared_utils.lxml_utils as lxml_utils
# import xml_docx_stylechecks.shared_utils.check_docx as check_docx
# import xml_docx_stylechecks.shared_utils.unzipDOCX as unzipDOCX

# # # # # # Set testing env variable:
os.environ["TRANSFORM_TEST_FLAG"] = 'true'
diff_file_list = [
    "{fname_noext}_StyleReport.txt",
    "stylereport.json",
    "{fname_noext}_unzipped/word/document.xml",
    "{fname_noext}_unzipped/word/footnotes.xml",
    "{fname_noext}_unzipped/word/endnotes.xml",
    "{fname_noext}_unzipped/word/styles.xml",
    "{fname_noext}_unzipped/word/comments.xml",
    "{fname_noext}_unzipped/word/commentsExtended.xml",
    "{fname_noext}_unzipped/word/commentsIds.xml",
    ]

# # # # # # LOCAL FUNCTIONS
# assumes that file lives in transform_testfiles_dir
def setupFileTest(file, tmpdir_base, transform_testfiles_dir):
    tmpfile = ''
    file = os.path.basename(file)
    fname_noext, ext = os.path.splitext(file)
    if ext == '.docx':
        tmpdir = os.path.join(tmpdir_base, fname_noext)
        if debug_diff_only != True:
            if os.path.exists(tmpdir):
                shutil.rmtree(tmpdir)
            os.mkdir(tmpdir)
            shutil.copy(os.path.join(transform_testfiles_dir, file), tmpdir)
        tmpfile = os.path.join(tmpdir, file)
    return tmpfile

def runTest(testfile):
    popen_params = ['python', rsv_main_path, testfile, 'direct', 'test.user@test.com', 'Test User']
    print popen_params
    if debug_diff_only != True:
        p = subprocess.Popen(popen_params)
        exitcode = p.wait()
    else:
        exitcode = 0
    # can wait on multitple processes if we want to 'multithread', move out of this function
    # exit_codes = [p.wait() for p in p1, p2]
    return exitcode

def udpateValidFiles(testfile, validfiles_basedir, update_valid_outputs, diff_file_list):
    tmpdir = os.path.dirname(testfile)
    fname_noext, ext = os.path.splitext(os.path.basename(testfile))
    validfiles_dir = os.path.join(validfiles_basedir, fname_noext)
    # # we are updating all via input param, delete existing valid_files to recreate
    if update_valid_outputs == True and os.path.exists(validfiles_dir):
        shutil.rmtree(validfiles_dir)
    # vaildated output does not exist, create via current output
    if not os.path.exists(validfiles_dir):
        os.makedirs(os.path.join(validfiles_dir, fname_noext, '{}_unzipped'.format(fname_noext), 'word'))
        for valid_file in diff_file_list:
            v_file = os.path.join(tmpdir, valid_file.format(fname_noext=fname_noext))
            v_file_dest = os.path.join(validfiles_dir, valid_file.format(fname_noext=fname_noext))
            shutil.copy(v_file, v_file_dest)

def getPrettyXml(xml_fname):
    dom = xml.dom.minidom.parse(xml_fname) #xml.dom.minidom.parseString(xml_string)
    prettyxml = dom.toprettyxml()
    return prettyxml

def diffFiles(diff_file, testfile, validfiles_basedir):
    # figure out dir paths
    tmpdir = os.path.dirname(testfile)
    fname_noext, ext = os.path.splitext(os.path.basename(testfile))
    validfiles_dir = os.path.join(validfiles_basedir, fname_noext)
    # define new and validated files
    valid_file = os.path.join(validfiles_dir, diff_file.format(fname_noext=fname_noext))
    new_file = os.path.join(tmpdir, diff_file.format(fname_noext=fname_noext))
    # run the diff!
    if os.path.splitext(diff_file)[1] == '.xml':
        # valid_file_txt = getPrettyXml(valid_file)#.decode('utf-8', 'replace')
        # new_file_txt = getPrettyXml(new_file)#.decode('utf-8', 'replace')
        # diff = difflib.unified_diff(
        #     valid_file_txt,
        #     new_file_txt,
        #     fromfile='validated: {}'.format(diff_file.format(fname_noext=fname_noext)),
        #     tofile='new: {}'.format(diff_file.format(fname_noext=fname_noext)),
        #     n=0,
        # )
        # for line in diff:
        #     sys.stdout.write(line)
        formatter = formatting.XMLFormatter(normalize=WS_BOTH, pretty_print=True, text_tags=(), formatting_tags=())
        # formatter = formatting.XMLFormatter(normalize=formatting.WS_BOTH, pretty_print=True)
        diff = main.diff_files(valid_file, new_file)
        print "xml diff", diff

    else:
        valid_file_txt = open(valid_file, 'r')
        new_file_txt = open(new_file, 'r')
        # no_id_vft = [line for line in valid_file_txt if '"para_id":' not in line]
        # no_id_nft = [line for line in new_file_txt if '"para_id":' not in line]
        diff = difflib.unified_diff(
            valid_file_txt.readlines(),
            new_file_txt.readlines(),
            fromfile='validated: {}'.format(diff_file.format(fname_noext=fname_noext)),
            tofile='new: {}'.format(diff_file.format(fname_noext=fname_noext)),
            n=0,
        )
        for line in diff:
            # for prefix in ('\t"para_id":'):
            #     if prefix in line:
            #         break
            # else:
            sys.stdout.write(line[1:])


            # sys.stdout.write(line)
    # diff = difflib.unified_diff(
    # text1_lines,
    # text2_lines,
    # lineterm='',
    # )

    # with open(valid_file, 'r') as valid_file_txt:
    #     with open(new_file, 'r') as new_file_txt:
    #         if os.path.splitext(diff_file)[1] == '.xml':
    #             valid_file_txt = getPrettyXml(valid_file_txt)
    #             new_file_txt = getPrettyXml(new_file_txt)
    # diff = difflib.unified_diff(
    #     valid_file_txt.readlines(),
    #     new_file_txt.readlines(),
    #     fromfile='validated: {}'.format(diff_file),
    #     tofile='new: {}'.format(diff_file),
    #     n=0,
    # )
    # for line in diff:
    #     sys.stdout.write(line)
        # for prefix in ('---', '+++', '@@'):
        #     if line.startswith(prefix):
        #         break
        # else:
        #     sys.stdout.write(line[1:])


if __name__ == '__main__':
    testfiles = []
    # setup tests
    if single_infile:
        testfile = setupFileTest(single_infile, tmpdir_base, transform_testfiles_dir)
        testfiles.append(testfile)
    for file in os.listdir(transform_testfiles_dir):
        testfile = setupFileTest(file, tmpdir_base, transform_testfiles_dir)
        if testfile:
            testfiles.append(testfile)
    # run tests
    for testfile in testfiles:
        exitcode = runTest(testfile)
        print "exitcode: ", exitcode
        if exitcode != 0:
            print "ALLEEERRT"
            break
        else:
            # update validated output
            udpateValidFiles(testfile, validfiles_basedir, update_valid_outputs, diff_file_list)
            # run diffs
            for diff_file in diff_file_list:
                diffFiles(diff_file, testfile, validfiles_basedir)

                # output = diffFiles(diff_file, testfile, validfiles_basedir)
                # print "{}\n".format(output)


    # for file in dir, create dir in tmp dir, copy file in, run thingy.
    # if no validation dir, create & copy em in; (make note in logs)
    # if validation files, diff / report on key items
    # could run a test singly by including it as input param!
    # NOW diff and look at output. two ways

    # unittest.main()
