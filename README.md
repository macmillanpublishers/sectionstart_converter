# sectionstart_converter (IN PROGRESS)
This tool is to update Microsoft Word documents that were styled with Macmillan styles prior to the release of our new Section Start styling.
The document.xml file will be directly edited using python/lxml, in order to add and update Section Start styles to conform with updated bookmaker and egalleymaker toolchains.

## Dependencies
The xml processing requires the lxml library for python, install like so:

`pip install lxml`

## Running the scripts

#### Unzip a Word document with the unzipDOCX.py
The command takes two args: the .docx file and the output dir for the root of the unzipped docx:
`python /path/to/unzipDOCX.py /path/to/file.docx /path/to/target/dir`


#### Edit unzipped (zip_root)/word/document.xml
This command takes one args: the root (parent folder) of the unzipped files:
`python /path/to/editDOCX.py /path/to/unzip_root`

It parses the document.xml so edits can be made, and writes the updated xml back to the same file upon completion.
This is an early draft, at this point it's just building block functions that we can use as we build out our full implementation.

Beyond actually building out our specific transformations, we will need to decide how to update the attached style template.
This may be able to be done with python/lxml; the app.xml file contains the template name, and styles.xml contains styles (including those from the attached template).
Or it could be done via Word/powershell prior to the conversion.


#### Re-Zip & deflate an unzipped .docx
This command takes two args: the root (parent folder) of the unzipped files, and path and name of the output .docx:
`python /path/to/zipDOCX.py /path/to/unzip_root /path/to/new/file.docx`


## Tracking Changes for reporting
There is a method in called "trackEdit" in _editDOCX.py_ that is logging every change made to the xml, for use in reporting later.
The structure, naming etc is up-for-grabs, but I think something like this can be translated into user-facing reports.

Right now it takes 3 parameters for every change: a paragraph element, a description, and an 'action' (a member of this set: ['insert', 'edit', 'remove']).  The function finds the paragraph-id value from the xml (if the action is 'remove', it gets the para-id of the previous paragraph since the current one will be gone), and writes the para-id description and action as a new dict in the changelog list.

At the end of the script, once all transformations are finished, the paragraph index for each change in the changelog list is found and and added to the dict. The paragraph index should allow us to find the page number in Word with VBA for user-facing reporting.
