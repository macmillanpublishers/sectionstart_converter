# pip install xmldiff #!<======
import sys, os, copy, re
import logging
import shutil
import subprocess
from xml.dom import minidom
import time

# # # append main project path to system path for imports to work
# mainproject_path = os.path.join(sys.path[0],'..')
# sys.path.append(mainproject_path)
# from xml_docx_stylechecks.shared_utils.decorators import benchmark as benchmark

# accept arguments
if len(sys.argv) == 2:
    if sys.argv[1] == 'update_valid_outputs':
        update_valid_outputs = True
        single_infile = ''
    elif sys.argv[1]:
        update_valid_outputs = False
        single_infile = sys.argv[1]
else:
    update_valid_outputs = False
    single_infile = ''

# key local paths
rsv_main_path = os.path.join(sys.path[0], '..', 'xml_docx_stylechecks', 'rsuitevalidate_main.py')
transform_testfiles_dir = os.path.join(sys.path[0], 'files_for_test', 'full_transform', 'test_docx_files')
validfiles_basedir = os.path.join(sys.path[0], 'files_for_test', 'full_transform', 'validated_output')
tmpdir_base = os.path.join(sys.path[0], 'files_for_test', 'tmp')
diff_outputdir = os.path.join(tmpdir_base, 'diff_outputs_{}'.format(time.strftime("%y%m%d%H%M%S")))

# # # # # # Set testing env variables:
debug_diff_only = False  # < - for testing diff parameters
os.environ["TRANSFORM_TEST_FLAG"] = 'true'
required_diff_files = [
    "{fname_noext}_StyleReport.txt",
    "stylereport.json",
    "{fname_noext}_unzipped/word/document.xml",
    ]
# styles & numbering are only here for duplicate style fix
nonrequired_diff_files = [
    "{fname_noext}_unzipped/word/footnotes.xml",
    "{fname_noext}_unzipped/word/endnotes.xml",
    "{fname_noext}_unzipped/word/comments.xml",
    "{fname_noext}_unzipped/word/commentsExtended.xml",
    "{fname_noext}_unzipped/word/commentsIds.xml",
    "{fname_noext}_unzipped/word/styles.xml",
    "{fname_noext}_unzipped/word/numbering.xml",
    "alerts.json",
    "WARNING.txt",
    "ERROR.txt",
    "NOTICE.txt"
    ]
diff_file_list = required_diff_files + nonrequired_diff_files

# # # # # # LOCAL FUNCTIONS
# assumes that file lives in transform_testfiles_dir
def setupFileTest(file, tmpdir_base, transform_testfiles_dir):
    tmpfile = ''
    file = os.path.basename(file)
    fname_noext, ext = os.path.splitext(file)
    ## setting exception with sanitizing fnames b/c it is prohibiting testing bad filenames!
    if not file.startswith('err_'):
        new_fname_noext = re.sub('[^\w-]','',fname_noext)
        new_fname = "{}{}".format(new_fname_noext,ext)
    else:
        new_fname_noext = fname_noext
        new_fname = file
    if ext == '.docx':
        tmpdir = os.path.join(tmpdir_base, new_fname_noext)
        tmpfile = os.path.join(tmpdir, new_fname)
        if debug_diff_only != True:
            if os.path.exists(tmpdir):
                shutil.rmtree(tmpdir)
            if not os.path.exists(tmpdir):
                os.makedirs(tmpdir)
            shutil.copy(os.path.join(transform_testfiles_dir, file), tmpdir)
            if new_fname_noext != fname_noext:
                os.rename(os.path.join(tmpdir, file), tmpfile)
    return tmpfile

def runTest(testfile):
    popen_params = ['python', rsv_main_path, testfile, 'direct', 'local']
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
    udpated_testfile_name = ''
    tmpdir = os.path.dirname(testfile)
    fname_noext, ext = os.path.splitext(os.path.basename(testfile))
    validfiles_dir = os.path.join(validfiles_basedir, fname_noext)
    # # we are updating all via input param, delete existing valid_files to recreate below
    if update_valid_outputs == True and os.path.exists(validfiles_dir):
        shutil.rmtree(validfiles_dir)
    # vaildated output does not exist, create via current output
    if not os.path.exists(validfiles_dir):
        os.makedirs(os.path.join(validfiles_dir, '{}_unzipped'.format(fname_noext), 'word'))
        for valid_file in diff_file_list:
            v_file = os.path.join(tmpdir, valid_file.format(fname_noext=fname_noext))
            v_file_dest = os.path.join(validfiles_dir, valid_file.format(fname_noext=fname_noext))
            # required files should be present, and copied. nonrequired we check if they exist first
            if valid_file in required_diff_files:
                shutil.copy(v_file, v_file_dest)
            elif valid_file in nonrequired_diff_files and os.path.exists(v_file):
                shutil.copy(v_file, v_file_dest)
        udpated_testfile_name = os.path.basename(testfile)
    return udpated_testfile_name

def getPrettyXml(xml_fname):
    filename_noext = os.path.splitext(xml_fname)[0]
    pretty_filename = '{}_pretty.xml'.format(filename_noext)
    if not os.path.isfile(pretty_filename):
        xmlstr = minidom.parse(xml_fname).toprettyxml(indent="   ")
        with open(pretty_filename, "w") as f:
            f.write(xmlstr.encode('utf-8'))
    return pretty_filename

def diffFiles(diff_file_list, testfile, validfiles_basedir, diff_outputdir):
    # figure out dir paths, init var
    files_with_diffs = []
    tmpdir = os.path.dirname(testfile)
    fname_noext, ext = os.path.splitext(os.path.basename(testfile))
    validfiles_dir = os.path.join(validfiles_basedir, fname_noext)
    for diff_file in diff_file_list:
        # define new and validated files
        valid_file = os.path.join(validfiles_dir, diff_file.format(fname_noext=fname_noext))
        new_file = os.path.join(tmpdir, diff_file.format(fname_noext=fname_noext))
        # if valid file exists, new one should too.
        if os.path.exists(valid_file):
            diff_outfile = os.path.join(diff_outputdir, '{}.txt'.format(fname_noext))
            df_shortname = os.path.basename(diff_file.format(fname_noext=fname_noext))
            # create rewrite prettified xml for good diff
            if os.path.splitext(diff_file)[1] == '.xml':
                valid_file = getPrettyXml(valid_file)
                new_file = getPrettyXml(new_file)

            # run the diff!
            f = open(diff_outfile,'a')
            f.write('\n--------- < = VALIDATED file | {} | NEW file = > ---------\n\n'.format(df_shortname))
            f.close()
            f = open(diff_outfile,'a')
            # run a diff and print results to file
            diff_val = subprocess.call(['diff', '-I', '"para_id":', '-I', '"para_index":', '-I', 'w14:paraId=', valid_file, new_file], stdout=f)
            # \/ optional different diff: unified, with 2 lines of context
            # diff_val = subprocess.call(['diff', '-U', '2', '-I', '"para_id":', '-I', '"para_index":', '-I', 'w14:paraId=', valid_file, new_file], stdout=f)
            f.close()

            # wrap up
            print "xml diff {} exit code: {}".format(df_shortname, diff_val)
            if diff_val != 0:
                files_with_diffs.append(df_shortname)
    return files_with_diffs


if __name__ == '__main__':
    # init variables to collect data
    testfiles = []
    all_files_with_diffs = {}
    err_testfiles = []
    validated_files_updated = []
    # setup tests
    if single_infile:
        testfile = setupFileTest(single_infile, tmpdir_base, transform_testfiles_dir)
        testfiles.append(testfile)
    else:
        for file in os.listdir(transform_testfiles_dir):
            testfile = setupFileTest(file, tmpdir_base, transform_testfiles_dir)
            if testfile:
                testfiles.append(testfile)
    # run tests
    for testfile in testfiles:
        exitcode = runTest(testfile)
        print "exitcode: ", exitcode
        if exitcode != 0:
            print "Fatal Error occurred during transform for {}".format(testfile)
            err_testfiles.append(testfile)
            break
        else:
            # update validated output
            updated_tf_name = udpateValidFiles(testfile, validfiles_basedir, update_valid_outputs, diff_file_list)
            if updated_tf_name:
                validated_files_updated.append(updated_tf_name)
            # run diffs
            if not os.path.exists(diff_outputdir):
                os.mkdir(diff_outputdir)
            files_with_diffs = diffFiles(diff_file_list, testfile, validfiles_basedir, diff_outputdir)
            if files_with_diffs:
                all_files_with_diffs[os.path.basename(testfile)] = files_with_diffs
    # TEST OUTPUT
    if validated_files_updated:
        print "\n\n * * Updated Validated files for testfile(s):"
        for file in validated_files_updated:
            print "\t- {}".format(file)
    if not all_files_with_diffs and not err_testfiles:
        print "\n\n * * * TESTS PASSED SUCCESSFULLY * * *\n"
        print ".docx files tested:"
        for file in testfiles:
            print "\t- {}".format(os.path.basename(file))
        # remove our diff folder, just for general cleanup
        shutil.rmtree(diff_outputdir)
    else:
        print "\n\n * * *  TESTS FAILED  * * * \n"
        if err_testfiles:
            print "Errors occurred running tests for the following testfiles:\n\t{}\n".format(err_testfiles)
        if all_files_with_diffs:
            print "Diffs in output were found for testfiles listed below; detailed diff output here: \n\t{}).\n".format(diff_outputdir)
            for testfile in all_files_with_diffs:
                print "- testfile: {}\n\tfiles with differences: {}".format(testfile, all_files_with_diffs[testfile])
    print "\n"
