# sectionstart_converter
This repo was originally intended to host several interrelated tools, all for parsing, reporting on, and editing MS Word files via python/lxml.
Currently only one is in use: 'rsuite_validate', invoked via main script: rsuitevalidate_main.py.

Also originally intended to use dropbox api's to gather submitter info: now it's being served via GoogleDrive and passed necessary parameters at runtime.

#### rsuite_validate
This tool accepts Word manuscripts and validates against a number of criteria, makes small edits not related to content or large errors, and returns a report and the edited document to the user, both in an outfolder and via email.
Internal documentation available [here](https://confluence.macmillan.com/display/RSUITE/RSuite+Validation).


#### Legacy products
These were items that this repo was initially intended to serve as well, all are retired for now, not refactored out of the code as of yet:

* xml_docx_stylechecks/converter_main.py -
This tool is to update Microsoft Word documents that were styled with Macmillan styles prior to the release of our new Section Start styling.
The document.xml file will be directly edited using python/lxml, in order to add and update Section Start styles to conform with updated bookmaker and egalleymaker toolchains.

* xml_docx_stylechecks/reporter_main.py -
This tool is to run functions formerly handled in our VBA Stylecheck macro(s). It will output a 'Style Report', both as a txt file, and send an email to the submitter. The original manuscript is not edited.

* xml_docx_stylechecks/validator_main.py -
This tool is to prepare a manuscript for egalley creation, as part of the bookmaker_validator toolchain; it fixes errors found in the 'Stylecheck' plus some other unique ones.


## Setup/Config

##### Dependencies
The xml processing requires the lxml library for python, install like so:
`pip install lxml`

##### cfg.py
Key setup items are found in cfg.py.
* tmpdir - this is a static path to a location on the host environment, and needs to be set.
* Global loglevel etc can also be set here.

## Testing
Unit and integration tests for rsuite_validate are documented in ./test/README.txt

## Running the scripts

#### Unzip a Word document with the unzipDOCX.py
The command takes two args: the .docx file and the output dir for the root of the unzipped docx:
`python /path/to/unzipDOCX.py /path/to/file.docx /path/to/target/dir`

#### Re-Zip & deflate an unzipped .docx
This command takes two args: the root (parent folder) of the unzipped files, and path and name of the output .docx:
`python /path/to/zipDOCX.py /path/to/unzip_root /path/to/new/file.docx`

#### reporter_main.py, converter_main.py or rsuitevalidate_main.py
Either of these commands takes one arg: the manuscript being evaluated / edited:
`python /path/to/reporter_main.py /path/to/file.docx`

#### validator_main.py
This command takes two args: the manuscript to be edited and the existing logfile in use by bookmaker_validator, so we can append to it instead of writing our own.
`python /path/to/validator_main.py /path/to/file.docx /path/to/existing/logfile.txt`
