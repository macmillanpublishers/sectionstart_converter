# sectionstart_converter
This repo was originally intended to host several interrelated tools, all for parsing, reporting on, and editing MS Word files via python/lxml.
Currently only one standalone product is in use: 'rsuite_validate'.
The 'validator_isbncheck' tool is also used, as part of the [egalleymaker](https://confluence.macmillan.com/display/EB/Egalleys) toolchain.
See more about legacy products originally served via this repo at the bottom of the README ('Legacy products')

### Docker implementation
For information on using the 'containerized' version of this tool, refer instead to this readme: ./docker_rsv/[README_docker.txt](https://github.com/macmillanpublishers/sectionstart_converter/blob/master/docker_rsv/README_docker.txt)
___
# Dependencies
* Python is required, versions 2.7.x and 3.9.x are supported.
* Requires the lxml library for python, install like so: `pip install lxml`
### Dependency: git submodule
External (Macmillan) git repo [RSuite_Word-template](https://github.com/macmillanpublishers/RSuite_Word-template) is added here as a submodule. It's checked out at a release tag, currently: _*v6.5.0*_
#### Initialize submodule
To initialize and update the submodule the first time after cloning or pulling the _sectionstart_converter_ repo, run: `git submodule update --init --recursive`
#### Update submodule
To update the submodule when pulling or switching branches (as needed), run: `git submodule update RSuite_Word-template`
#### Edit submodule checked-out commit
To peg the submodule HEAD to a new tag, first update it with the above command. Then cd into the submodule dir, checkout the new tag, and commit your changes.
___
# Products
## rsuite_validate
This tool accepts Word manuscripts and validates against a number of criteria, makes small edits not related to content or large errors, and returns a report and the edited document to the user, both in an outfolder and via email.
Internal documentation available [here](https://confluence.macmillan.com/display/RSUITE/RSuite+Validation).
#### Product-specific Dependencies
Dependencies for tests, local runs:
* Supplemental python libraries are required, install via pip like so: `pip install requests six`
Additional dependencies for production or staging environment (unless running via Docker):
* git-repo: '[bookmaker_connectors](https://github.com/macmillanpublishers/bookmaker_connectors)' must be cloned locally as a sibling directory to this repo ('sectionstart_converter').
* git-repo: '[bookmaker_authkeys](https://github.com/macmillanpublishers/bookmaker_authkeys)' must be cloned locally as a sibling directory to this repo ('sectionstart_converter'). This repo is private and will also require [decryption](https://confluence.macmillan.com/display/PWG/Using+git-crypt+to+encrypt+files+on+github).

#### Running rsuite_validate
To run this tool directly in the cmd line:

`python /path/to/rsuitevalidate_main.py '/path/to/file.docx' 'direct' 'local'`
* Running with the 'local' parameter above skips sending notification emails, skips posting final files to the OUTfolder via api, and preserves the tmpfolder contents for troubleshooting (working tmpfiles and dirs will be created in the same directory as testfile.docx)
* You can change loglevel from INFO to DEBUG etc. in _xml_docx_stylechecks/cfg.py_
* To run rsuite_validate with emails and api, the call looks like this instead:
`python /path/to/rsuitevalidate_main.py /path/to/file.docx 'direct' 'user.email@domain.com' 'User Name'`

#### Tests
Unit and integration tests for rsuite_validate are documented in ./test/[README_tests.txt](https://github.com/macmillanpublishers/sectionstart_converter/blob/master/test/README_tests.txt)
___
## validator_isbncheck
This tool is run as part of the egalleymaker process, to capture styled ISBNs, and style & capture unstyled ISBN's where needed. It logs them to a JSON where the rest of the egalleymaker process can use them.

#### Running validator_isbncheck.py
This command takes two args: the manuscript to be edited and the existing logfile in use by bookmaker_validator, so we can append to it instead of writing our own.
`python /path/to/validator_isbncheck.py /path/to/file.docx /path/to/existing/logfile.txt`
___
## Other standalone tools
#### Unzip a Word document with the unzipDOCX.py
The command takes two args: the .docx file and the output dir for the root of the unzipped docx:
`python /path/to/unzipDOCX.py /path/to/file.docx /path/to/target/dir`
#### Re-Zip & deflate an unzipped .docx
This command takes two args: the root (parent folder) of the unzipped files, and path and name of the output .docx:
`python /path/to/zipDOCX.py /path/to/unzip_root /path/to/new/file.docx`
___

# Legacy products
These were items that this repo was initially intended to serve as well, all are retired for now, not refactored out of the code as of yet:

* xml_docx_stylechecks/converter_main.py -
This tool is to update Microsoft Word documents that were styled with Macmillan styles prior to the release of our new Section Start styling.
The document.xml file will be directly edited using python/lxml, in order to add and update Section Start styles to conform with updated bookmaker and egalleymaker toolchains.

* xml_docx_stylechecks/reporter_main.py -
This tool is to run functions formerly handled in our VBA Stylecheck macro(s). It will output a 'Style Report', both as a txt file, and send an email to the submitter. The original manuscript is not edited.

* xml_docx_stylechecks/validator_main.py -
This tool is to prepare a manuscript for egalley creation, as part of the bookmaker_validator toolchain; it fixes errors found in the 'Stylecheck' plus some other unique ones.
