# sectionstart_converter (IN PROGRESS)
This tool is being expanded, and now has three main, distinct functions, with a 'main' script for each. They share many of the same functions, logging and dependencies, and key scripts (zipDOCX.py & unzipDOCX.py) can be run independently as standalone processes.

##### xml_docx_stylechecks/converter_main.py -
This tool is to update Microsoft Word documents that were styled with Macmillan styles prior to the release of our new Section Start styling.
The document.xml file will be directly edited using python/lxml, in order to add and update Section Start styles to conform with updated bookmaker and egalleymaker toolchains.

##### xml_docx_stylechecks/reporter_main.py -
This tool is to run functions formerly handled in our VBA Stylecheck macro(s). It will output a 'Style Report', both as a txt file, and send an email to the submitter. The original manuscript is not edited.

##### xml_docx_stylechecks/validator_main.py -
This tool is to prepare a manuscript for egalley creation, as part of the bookmaker_validator toolchain; it fixes errors found in the 'Stylecheck' plus some other unique ones.


## Dependencies
The xml processing requires the lxml library for python, install like so:

`pip install lxml`

## Setup
Setup items are found in cfg.py. Key items:
* tmpdir - this is a static path to a location on the host environment, and needs to be set.
* dropboxfolder - This is a path set for the default Macmillan Dropbox folder on Windows or Mac, and would need to be edited for use in any other environment.

Most other items are setup in relation to these two.

## Running the scripts

#### Unzip a Word document with the unzipDOCX.py
The command takes two args: the .docx file and the output dir for the root of the unzipped docx:
`python /path/to/unzipDOCX.py /path/to/file.docx /path/to/target/dir`

#### Re-Zip & deflate an unzipped .docx
This command takes two args: the root (parent folder) of the unzipped files, and path and name of the output .docx:
`python /path/to/zipDOCX.py /path/to/unzip_root /path/to/new/file.docx`

#### reporter_main.py or converter_main.py
Either of these commands takes one arg: the manuscript being evaluated / edited:
`python /path/to/reporter_main.py /path/to/file.docx`

#### validator_main.py
This command takes two args: the manuscript to be edited and the existing logfile in use by bookmaker_validator, so we can append to it instead of writing our own.
`python /path/to/validator_main.py /path/to/file.docx /path/to/existing/logfile.txt`
