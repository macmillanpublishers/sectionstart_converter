* * * * *  TESTING DEPENDENCIES  * * * * *
These tests require that the RSuite-Word-template git-repo be cloned locally, into the same parent dir as this repo.
1) cd to parent dir of 'sectionstart_converter'
2) `git clone https://github.com/macmillanpublishers/RSuite_Word-template.git`


* * * * *  UNIT TESTING  * * * * *
You can run all unit tests from a Mac by double-clicking the file:
  "sh_run_unit_tests.command"   (found in the same dir as this README)

Otherwise you can run them manually via cmd-line as follows:
- Run all unit tests: cd to parent dir of /test and run `python -m unittest discover -v`
- Run all unit tests from one file: `python -m unittest test.test_rsuite_validations`
- Run all unit tests in class: `python -m unittest test.test_rsuite_validations.Tests`
- Run one unit test: `python -m unittest test.test_rsuite_validations.Tests.test_deleteObjects_fromNode`


* * * * *  INTEGRATION TESTING  * * * * *
There are two sets of transform tests: one for rsuite_validate and one for validator_isbncheck (used by egalleymaker)

RSUITE_VALIDATE transform tests:
You can run an rsuite_validate Integration test for all test docs in 'test/files_for_test/full_transform/test_docx_files'
  from a Mac by double-clicking file:
  "sh_run_transform_tests.command"    (in the same dir as this README)

Otherwise you can run them manually via cmd-line:
- Run transform test for all files: `python rsvalidate_transform_tests.py`
- Run transform test for one test-file: `python rsvalidate_transform_tests.py testfile.docx (filename of file in test_docxfiles)`

This runs a diff for each docx against validated output files from a previous known-good run.
  (It does not test I/O such as rest-api / Drive api / sending mails)
Results are summarized at the end of the run (takes about 2 min per testfile).

VALIDATOR_ISBNCHECK transform tests:
You can run a validator_isbncheck Integration test for all test docs in 'test/files_for_test/full_transform/test_docx_files'
  from a Mac by double-clicking file:
  "sh_run_isbncheck_transform_tests.command"    (in the same dir as this README)

Otherwise you can run them manually via cmd-line:
- Run transform test for all files: `python isbncheck_transform_tests.py`
- Run transform test for one test-file: `python isbncheck_transform_tests.py testfile.docx (filename of file in test_docxfiles)`


* * * * *  Updating Validated Files for Integration Tests * * * * *
Validated files are stored in 'test/files_for_test/full_transform/test_docx_files/validated_output'
  (* for validator_isbncheck the finaldir is 'vi_validated_output' instead)
Once you've verified the output changes are as expected,
  there are three ways to update validated files with tmp files from a new run:

--- Update validated files for ALL testfiles:
- On a Mac, double-click the file:  "sh_update_valid_files.command"
- Run via cmd line with parameter: `python rsvalidate_transform_tests.py update_valid_outputs`
--- Update validated files for one specific testfile:
- Run via cmd line with parameter: `python rsvalidate_transform_tests.py update_valid_outputs (filename of file in test_docxfiles)` 
- Drag 'n' drop files from 'test/files_for_test/tmp' to relative 'validated_output' dir.
