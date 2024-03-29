######### IMPORT PY LIBRARIES
import logging
import textwrap

# # initialize logger
logger = logging.getLogger(__name__)


# #---------------------  METHODS
def getBanners():
    banners = {
        "validator_noerr": textwrap.dedent("""\
            EGALLEY VALIDATION REPORT

            The metadata, sections, and illustrations that Egalleymaker identified in your manuscript are listed below. If you discover incorrect information, correct the manuscript and run it through Egalleymaker again.
        """),
        "validator_err": textwrap.dedent("""\
            EGALLEY VALIDATION REPORT

            Please peruse items below to verify document info.

            ATTN: One or more items of note turned up during document validation:
            {v_warning_banner}

            See below for details.

            (Need help reviewing this report? Visit {helpurl})
        """),
        "converter": textwrap.dedent("""\
            The Style Converter has processed your manuscript. The revised file is attached here.

            YOU ARE RESPONSIBLE FOR VERIFYING THAT THE SECTION-START PARAGRAPHS HAVE BEEN INSERTED CORRECTLY.

            The table below lists:
            (1) every section that the Style Converter identified based on correct rules for using styles, and
            (2) the text that it added to the section-start paragraph (for the ebook TOC and NCX links).

            IF ANY SECTION-START PARAGRAPHS ARE INCORRECT OR MISSING, IT IS YOUR RESPONSIBILITY TO MAKE THE CORRECTIONS.

            (For assistance visit {helpurl})
        """),
        "reporter_noerr": textwrap.dedent("""\
            CONGRATULATIONS! YOU PASSED!

            But you're not done yet. Please check the info listed below.
        """),
        "reporter_err": textwrap.dedent("""\
            OOPS!

            Problems were found with the styles in your document.
        """)
    }
    return banners

# This method defines what goes in the StyleReport txt and mail outputs, in what order, + formatting.
# See the commented "SAMPLE RECIPE ENTRY" below for details on each field.  All fields should be optional,
#   though text, title or dict_category_name must be present for something to print
def getReportRecipe(titlestyle, authorstyle, isbnstyle, logostyle, booksection_stylename, notessection_stylename, support_email_address):
    report_recipe = {
        # # # # # # # # # # # # #  SAMPLE RECIPE ENTRY:
        # # # # # # # # # # # # # # # # # # # # # # # # # #
    	# "00_example": {                            # < The leading digits are to set the order of appearance on the report
        #   "include_for": ["reporter"],       # < Add the parent script basename to this list if you want to include this item for said script's report
        #                                            #   it when generateReport() is invoked from that parent script.
    	# 	"title": "TEST",                         # < This is title of a section on the report: centered and surrounded by hyphens
    	# 	"text": "Example {}".format("test"),     # < This is a line of text that will under title if present.
    	# 	"dict_category_name": "title_paras",     # < The corresponding category name in report_dict
    	# 	"line_template": "Contents: {para_string:.>33}",    # < This is a string template for how you want each entry from
        #                                                       #   the report_dict category data presented. To insert a value
        #                                                       #   from a report_dict entry, put the key name in {brackets}
        #                                                       #   (the trailing colons etc are are extra formatting)
        #   "suppress_table_note": True,    # Even if this is a table para don't include "  (< this item is from a table)" from gen_report.py
    	# 	"required": True,       # < If it's an error if a report_dict category is empty or not present, mark this True
    	# 	                        #   (If this is true you will need an "errstring" entry too)
        #   "v_warning_banner": "Alert string",   # this is for validator only scripts - if any edits were made or unsupported styles were found, we want to surface
        #                                       a different banner on the report output. Including this key=True signals that we want that warning.
        #   "badnews": 'any',        # < If you want any entry from this report_dict category in the Error List,mark this True.. if one entry is ok but more are errors, use value 'one_allowed'
        #   "badnews_type": 'warning' # Specify type of badnews, whether a warning or error. If neither specified, error is presumed
        #   "errstring": "No paragraphs."   # < The base string you want used to appear in the report's Error list
        #   "summary": 'true' # Warnings with badnews: 'any' are typically listed singly. This value overrides that. Vice versa for badnews_type: note
        #   "alternate_content": {          # < If you want an alternate title or text element to appear when
        #       "title": "TEST FAIL"        #   report_dict category is empty or not present, set them here. If you
        #       "text": "No title paras."   #   set text but not title, the original title value will be used, and vice versa
        #    }
    	# },
    	"01_metadata_heading": {
            "include_for": ["reporter", "validator", "rsuitevalidate"],
    		"title": "METADATA",
    		"text": "If any of the information below is wrong, please fix the associated styles in the manuscript."
    	},
    	"02_metadata_title": {
            "include_for": ["reporter", "validator"],
    		"title": "",
    		"text": "** {} **".format(titlestyle),
    		"dict_category_name": "title_paras",
    		"line_template": "{para_string}",
    		"required": True,
            "errstring": "No paragraph styled with '{}' found".format(titlestyle),
            "alternate_content": {
                "text": "** {} **\nNo title paras detected.".format(titlestyle)
            }
    	},
    	"02_metadata_title(rsuite_validate)": {
            "include_for": ["rsuitevalidate"],
    		"title": "",
    		"text": "** {} **".format(titlestyle),
    		"dict_category_name": "title_paras",
    		"line_template": "{para_string}",
            "alternate_content": {
                "text": "** {} **\nNo title paras detected.".format(titlestyle)
            }
    	},
    	"03_metadata_author": {
            "include_for": ["reporter", "validator"],
    		"title": "",
    		"text": "** {} **".format(authorstyle),
    		"dict_category_name": "author_paras",
    		"line_template": "{para_string}",
    		"required": True,
            "errstring": "No paragraph styled with '{}' found".format(authorstyle),
            "alternate_content": {
                "text": "** {} **\nNo author paras detected.".format(authorstyle)
            }
    	},
    	"03_metadata_author(rsuite_validate)": {
            "include_for": ["rsuitevalidate"],
    		"title": "",
    		"text": "** {} **".format(authorstyle),
    		"dict_category_name": "author_paras",
    		"line_template": "{para_string}",
            "alternate_content": {
                "text": "** {} **\nNo author paras detected.".format(authorstyle)
            }
    	},
    	"04_metadata_isbn": {
            "include_for": ["reporter"],
    		"title": "",
    		"text": "** {} **".format(isbnstyle),
    		"dict_category_name": "isbn_spans",
    		"line_template": "{para_string}",
    		"required": True,
            "errstring": "No ISBN styled with '{}' detected".format(isbnstyle),
            "alternate_content": {
                "text": "** {} **\nNo styled isbns detected.".format(isbnstyle)
            }
    	},
    	"04_metadata_isbn(rsuite_validate)": {
            "include_for": ["rsuitevalidate"],
    		"title": "",
    		"text": "** {} **".format(isbnstyle),
    		"dict_category_name": "isbn_spans",
    		"line_template": "{para_string}",
            "alternate_content": {
                "text": "** {} **\nNo styled isbns detected.".format(isbnstyle)
            }
    	},
    	"04_metadata_isbn(validator)": {
            "include_for": ["validator"],
    		"title": "",
    		"text": "** {} **".format(isbnstyle),
    		"dict_category_name": "isbn_spans",
    		"line_template": "{para_string}  <---- (This is the ebook ISBN)",
    		"required": True,
            "errstring": "No ISBN styled with '{}' detected".format(isbnstyle),
            "alternate_content": {
                "text": "** {} **\nNo styled isbns detected.".format(isbnstyle)
            }
    	},
    	"05_image_holders": {
            "include_for": ["reporter", "validator", "rsuitevalidate"],
    		"title": "ILLUSTRATION LIST",
    		"text": "Verify that this list of illustrations includes only the filenames of your illustrations.\n",
    		"dict_category_name": "image_holders__sort_by_index",
    		"line_template": "{description}",#\n    -located in {parent_section_start_type}: {parent_section_start_content}.",# (Paragraph {para_index})",
            "alternate_content": {
                "text": "no illustrations detected."
            }
    	},
    	"06_section_start_list": {
            "include_for": ["reporter", "validator", "rsuitevalidate"],
    		"title": "SECTIONS FOUND",
    		"text": "{:90}\n".format("The Style Report identified the following sections; note that the content of each\nSection-Start paragraph will be used for the ebook TOC/NCX. If any of these are incorrect,\nedit the manuscript file to add or remove incorrect Section-Start styles or content."),
    		# "text": "",""#"Here is a list of all sections detected in your manuscript",
    		"dict_category_name": "section_start_found",
    		"line_template": "{parent_section_start_type:.<33} {parent_section_start_content:57}",
    		"required": True,
            "errstring": "No sections found. Sections must be styled with Section-Start styles.",
            "alternate_content": {
                "text": "No sections detected."
            }
    	},
    	"07_macmillan_style_1st_use": {
            "include_for": ["reporter", "rsuitevalidate"],
    		"title": "MACMILLAN STYLES IN USE (BY SECTION)",
    		"text": "\n{:_^45} {:_^55}".format("PARAGRAPH STYLES IN ORDER OF FIRST USE","EXCERPT FROM FIRST USE"),
    		# "text": "\n{:_^40} {:_^50}".format("PARAGRAPH STYLES IN ORDER OF FIRST USE","EXCERPT FROM FIRST USE"),
    		"dict_category_name": "Macmillan_style_first_use",
    		"new_section_text": "\n* {parent_section_start_type}: {parent_section_start_content}",
    		"line_template": "{description:.<45} {para_string:60}",
    		"required": True,
            "errstring": "No Macmillan styled paragraphs were found in the manuscript.",
            "alternate_content": {
                "text": "No Macmillan styles detected."
            }
        },       #
    	"08_macmillan_character_style_1st_use": {
            "include_for": ["reporter", "rsuitevalidate"],
    		"text": "{:_^45}".format("CHARACTER STYLES IN USE"),
    		# "text": "{:_^40}".format("CHARACTER STYLES IN USE"),
    		"dict_category_name": "Macmillan_charstyle_first_use",
    		"line_template": "{description}",
    		"required": "n-a",
            "alternate_content": {
                "text": "{:_^45}\nNo character styles detected.".format("character_styles_in_use")
            }
    	},
        # re-setting count at 20 for coverter-specific items to make renumbering simpler
    	"20_section_start_added(converter)": {
            "include_for": ["converter"],
    		"title": "SECTION START PARAGRAPHS INSERTED",
    		# "text": "{:^48} {:_^50}".format("\033[4mstyles_in_order_of_appearance\033[0m","styled_content_excerpt_from_first_use"),
    		"text": "{:_^40} {:_^40}".format("Section-Start style","paragraph content"),
    		"dict_category_name": "section_start_found",
    		"line_template": "{parent_section_start_type:.<40} {para_string:50}",#description
    		"required": True,
            "errstring": "No Section-Start insertion points detected.",
            "alternate_content": {
                "text": "No Section-Start insertion points detected; so no Section-Start paras have been added."
            }
    	},
        # re-setting count at 25 for validator-specific items to make renumbering simpler
    	"25_corrections_heading(validator)": {
            "include_for": ["validator"],
    		"title": " ************************** UNSUPPORTED STYLES FOUND ************************* ",
    		"text": "If any non-Macmillan or non-Bookmaker styles were detected in the manuscript, they are displayed here:"
    	},
    	"26_non-Macmillan_styles(validator)": {
            "include_for": ["validator"],
    		"title": "NON-MACMILLAN PARAGRAPH STYLES",
    		"text": "Non-Macmillan styles detected.\nContent styled with non-Macmillan styles may not appear properly-styled in your egalley.\n",
    		"dict_category_name": "non-Macmillan_style_used",
    		"line_template": "- style '{description}': found in section '{parent_section_start_type}': {parent_section_start_content}",
    		"required": "n-a",
            "v_warning_banner": "Unsupported (non-Macmillan) style(s) found."
    	},
    	"27_non-Macmillan_charstyles(validator)": {
            "include_for": ["validator"],
    		"title": "NON-MACMILLAN CHARACTER STYLES",
    		"text": "Non-Macmillan styles detected.\nContent styled with non-Macmillan styles may not appear properly-styled in your egalley.\n",
    		"dict_category_name": "non-Macmillan_charstyle_used",
    		"line_template": "- character style '{description}': found in use.",
    		"required": "n-a",
            "v_warning_banner": "Unsupported (non-Macmillan) style(s) found."
    	},
    	"28_non-Bookmaker_styles(validator)": {
            "include_for": ["validator"],
    		"title": "NON-BOOKMAKER STYLES",
    		"text": "Non-Bookmaker styles detected.\nContent styled with non-Bookmaker styles may not appear properly-styled in your egalley.\n",
    		"dict_category_name": "non_bookmaker_macmillan_style",
    		"line_template": "- style '{description}': found in section '{parent_section_start_type}': {parent_section_start_content}",
    		"required": "n-a",
            "v_warning_banner": "Unsupported (non-Bookmaker) style(s) found."
    	},
    	"29_corrections_heading(validator)": {
            "include_for": ["validator"],
    		"title": " ******************** EDITS MADE DURING DOCUMENT VALIDATION ******************* ",
    		"text": "Egalleymaker makes some automatic adjustments and corrections to your manuscript to help make better egalleys.\nIf any changes were made, they are noted below:"
    	},
    	"30_section_start_added(validator)": {
            "include_for": ["validator"],
    		"title": "SECTION START PARAS AUTO-INSERTED",
    		"text": "{:_^40} {:_^40}".format("Section-Start_style","paragraph_content"),
    		"dict_category_name": "section_start_needed__sort_by_index",
    		"line_template": "{parent_section_start_type:.<40} {para_string:50}", #description
    		"required": "n-a",
            "v_warning_banner": "Edit(s) made during document validation: Section Start paragraph(s) inserted."
    	},
    	"31_added_content_to_sectionstart_para(validator)": {
            "include_for": ["validator"],
    		"title": "ADDED CONTENT TO SECTION START PARAGRAPH(S)",
    		"text": "Section-Start paragraphs cannot be empty. Content was auto-added to the following Section-Start paragraph(s):\n",
    		"dict_category_name": "wrote_to_empty_section_start_para",
    		"line_template": "{parent_section_start_type:.<40} New text: {para_string}",
    		"required": "n-a",
            "v_warning_banner": "Edit(s) made during document validation: content was added to empty Section-Start paragraph(s)."
    	},
    	"32_removed_empty_firstlast_para(validator)": {
            "include_for": ["validator"],
    		"title": "REMOVED EMPTY FIRST OR LAST PARA",
    		"text": "An empty first or last paragraph causes problems with bookmaker: one (or both) were found and removed.\n",
    		"dict_category_name": "removed_empty_firstlast_para",
    		"line_template": "- {description}",
    		"required": "n-a",
            "v_warning_banner": "Edit(s) made during document validation: removed empty first/last paragraph(s)."
    	},
    	"33_rm_charstyles_in_headings(validator)": {
            "include_for": ["validator"],
    		"title": "CHARACTER STYLES REMOVED FROM HEADINGS",
    		"text": "Character styles in headings cause problems with ebook TOC creation... some were found and removed:\n",
    		"dict_category_name": "rm_charstyle_from_heading",
    		"line_template": "- {description} (from '{parent_section_start_type}': {parent_section_start_content})",
    		"required": "n-a",
            "v_warning_banner": "Edit(s) made during document validation: removed character styles from Heading(s)."
    	},
    	"34_added_reqrd_sectionstart(validator)": {
            "include_for": ["validator"],
    		"title": "ADDED TITLEPAGE AND/OR COPYRIGHT PAGE",
    		"text": "Titlepage and Copyright-page are required. One or both was missing and had to be added:\n",
    		"dict_category_name": "added_required_section_start",
    		"line_template": "- {description}",
    		"required": "n-a",
            "v_warning_banner": "Edit(s) made during document validation: added required book section(s) (Titlepage or Copyright page)."
    	},
    	"35_added_bookinfo_to_titlepage(validator)": {
            "include_for": ["validator"],
    		"title": "ADDED METADATA TO TITLEPAGE",
    		"text": "The Titlepage section must contain styled Title and Author paras. One or both was missing and had to be added:\n",
    		"dict_category_name": "added_required_book_info",
    		"line_template": "- {description}",
    		"required": "n-a",
            "v_warning_banner": "Edit(s) made during document validation: added missing title &/or author info to Titlepage."
    	},
    	"36_rm_shapesandbreaks(validator)": {
            "include_for": ["validator"],
    		"title": "SECTION BREAKS, SHAPES AND GRAPHICS REMOVED",
    		"text": "Section Break(s) and or inserted graphics were removed in the following sections:\n",
    		"dict_category_name": "deleted_objects",
    		"line_template": "- {description} from '{parent_section_start_type}': {parent_section_start_content}",
    		"required": "n-a",
            "v_warning_banner": "Edit(s) made during document validation: removed section break(s) &/or shape(s)."
    	},
    	"37_get_oneline_titlepara(validator)": {
            "include_for": ["validator"],
    		"title": "COMBINED MULTILINE TITLE",
    		"text": "The Title must not contain softbreaks or multiple paragraphs. This was corrected in your manuscript.\n",
    		"dict_category_name": "concatenated_extra_titlepara_and_removed",
    		"line_template": "New title: '{description}'",
    		"required": "n-a",
            "v_warning_banner": "Edit(s) made during document validation: combined multiline Title."
    	},
    	"70_non_section_start_styled_firstpara": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["reporter", "validator"],
    		"dict_category_name": "non_section_start_styled_firstpara",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "First paragraph of document styled with non-Section Start style ('{description}')."# (Paragraph {para_index})"
    	},
    	"71__non_section_BOOK_styled_firstpara": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "non_section_BOOK_styled_firstpara",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "First paragraph of document is styled with '{description}' instead of '%s'." % booksection_stylename
    	},
    	"72__non_section_start_styled_secondpara": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "non_section_start_styled_secondpara",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Second paragraph of document styled with non-Section Start style: '{description}'."# (Paragraph {para_index})"
    	},
    	"72.1_missing_required_notes_section": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "missing_notes_section",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "File contains embedded endnotes but there is no para styled '%s' in the main body of the document. This section is required to process notes -- please add." % notessection_stylename# (Paragraph {para_index})"
    	},
    	"73_non_macmillan_styles": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["reporter", "rsuitevalidate"],
    		"dict_category_name": "non-Macmillan_style_used",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Non-Macmillan style '{description}' in {parent_section_start_type}: {parent_section_start_content}."# (Paragraph {para_index})"
    	},
    	"73.2_non_macmillan_styles_in_table": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["reporter", "rsuitevalidate"],
    		"dict_category_name": "non-Macmillan_style_used_in_table",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'error',
            "summary": True,
            "suppress_table_note": True,
            "errstring": "{section_count} paragraphs styled with Non-Macmillan style '{description}' were found in a table(s). The parent table(s) can be found in {parent_section_start_type}: {parent_section_start_content}."# (Paragraph {para_index})"
    	},
    	"73.5_non_macmillan_charstyles": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["reporter", "rsuitevalidate"],
    		"dict_category_name": "non-Macmillan_charstyle_used",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Non-Macmillan character style found in use: '{description}'."# (Paragraph {para_index})"
    	},
    	"73.6_non_macmillan_charstyles_removed": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "non-Macmillan_charstyle_removed",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'note',
            "errstring": "Found and removed non-Macmillan character style from manuscript: '{description}'."# (Paragraph {para_index})"
    	},
    	"74_non_bookmaker_style": {
            "include_for": ["reporter"],
    		"dict_category_name": "non_bookmaker_macmillan_style",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Non-Bookmaker style: '{description}' in {parent_section_start_type}: {parent_section_start_content}."# (Paragraph {para_index})"
    	},
    	"75_empty_section_start_para": {
            "include_for": ["reporter"],
    		"dict_category_name": "empty_section_start_para",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Empty Section-Start paragraph: found a '{description}' para with no text."# (Paragraph {para_index})"
    	},
    	"75.5_table_blank_para": {
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "table_blank_para",
    		"line_template": "",
            "suppress_table_note": True,
    		"badnews": 'any',
            "badnews_type": 'warning',
            "summary": True,
            "errstring": "{section_count} table cell(s) containing only blank paragraph(s) were found in {parent_section_start_type}: {parent_section_start_content}. Please confirm these empty table cells are intended, and if not, remove as needed."
    	},
    	"75.6_table_blank_para_notes": {
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "table_blank_para_notes",
    		"line_template": "",
            "suppress_table_note": True,
    		"badnews": 'any',
            "badnews_type": 'warning',
            "summary": True,
            "errstring": "{notes_count} table cell(s) containing only blank paragraph(s) were found in '{notes_type}'. Please confirm these empty table cells are intended, and if not, remove as needed."
    	},
    	"76_section_blankpara": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "removed_section_blank_para",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'warning',
            "errstring": "A blank Section-Start paragraph was removed: '{description}'. The removal of a Section-Start paragraph may have led to other errors with your manuscript."# (Paragraph {para_index})"
    	},
    	"76.5_fm_section_in_body": {
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "fm_section_in_body",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'error',
            "errstring": "'{descriptionA}' found in the body of the MS. This style can only be used in the front matter. Please restyle with a valid section style."
    	},
    	"77_container_blankpara": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "removed_container_blank_para",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'warning',
            "errstring": "A blank Container paragraph with style '{descriptionA}' was removed from {descriptionB}. This is being brought to your attention in case it led to other errors with your manuscript."# (Paragraph {para_index})"
    	},
    	"78_spacebreak_blankpara": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "removed_spacebreak_blank_para",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'warning',
            "errstring": "A blank paragraph with style '{descriptionA}' was removed from {descriptionB}. Even a Space-break or Separator paragraph must have content to be processed by RSuite."# (Paragraph {para_index})"
    	},
    	"79_container_error": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "container_error",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "No Container END para found for: '{description}', in {parent_section_start_type}: {parent_section_start_content}."# (Paragraph {para_index})"
    	},
    	"79.5_container_error": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "container_end_error",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Found a Container END para with no corresponding Container Start; in {parent_section_start_type}: {parent_section_start_content}."# (Paragraph {para_index})"
    	},
    	"80_list_error": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "list_nesting_err",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Improper list nesting: {description} in {parent_section_start_type} {parent_section_start_content} (starts with text: {para_string})"# (Paragraph {para_index})"
    	},
    	"81_list_change_error": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "list_change_err",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "List type changed in the middle of a list: {description} in {parent_section_start_type} {parent_section_start_content} (starts with text: {para_string})"# (Paragraph {para_index})"
    	},
    	"81_illegal_style_in_table": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "illegal_style_in_table",
    		"line_template": "",
            "suppress_table_note": True,
    		"badnews": 'any',
            "errstring": "Found a Container or Section styled paragraph in a table cell: {description} (in {parent_section_start_type} {parent_section_start_content}). Sections and/or Containers cannot be inside of tables."# (Paragraph {para_index})"
    	},
    	"82_endnote_text_misstyled": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "improperly_styled_endnote",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Endnote paragraph styled as '{description}' instead of 'Endnote Text': (Note beginning {para_string})."
    	},
    	"83_footnote_text_misstyled": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "improperly_styled_footnote",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Footnote paragraph styled as '{description}' instead of 'Footnote Text' (Note beginning {para_string})."# (Paragraph {para_index})"
    	},
    	"84_bad_image_holder_ext": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "image_holder_ext_error",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Missing or unsupported file extension for '{descriptionA}'-styled text: '{descriptionB}' (located in {parent_section_start_type}: {parent_section_start_content}). Your filename must end with a supported file extension (one of {valid_file_extensions})"
    	},
    	"85_bad_image_holder_char": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "image_holder_badchar",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Unsupported character(s) in '{descriptionA}'-styled text: '{descriptionB}' (located in {parent_section_start_type}: {parent_section_start_content}). Image placement styles may contain only alphanumeric characters, underscores, or hyphens."
    	},
        "86_invalid_symfonts": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "invalid_symfonts",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Encountered use(s) of unsupported symbol-font: '{description}'. Please email %s for assistance resolving this issue." % support_email_address
    	},
    	"90_list_change_warning": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "list_change_warning",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'warning',
            "errstring": "List type changed in the middle of a list: {description} in {parent_section_start_type} {parent_section_start_content} (starts with text: {para_string})"
    	},
    	"93_note_markers_wrong_style": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "note_markers_wrong_style",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'note',
            "errstring": "Fixed {count} Endnote and/or Footnote markers with incorrect character styles applied."
    	},
    	"93.5_rogue_noteref_style_use": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "rogue_noteref_style_use",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'note',
            "errstring": "Removed {count} extraneous instance(s) of 'Endnote Ref' and/or 'Footnote Ref' character styles."
    	},
    	"94_deleted_generic_blankpara": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "removed_blank_para",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'note',
            "errstring": "{count} blank paragraph(s) deleted from the manuscript."
    	},
    	"95_deleted_shape_summary": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "deleted_objects-shapes",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'note',
            "errstring": "{count} shape object(s) deleted from the manuscript."
    	},
    	"95.5_found_empty_note": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "found_empty_note",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'warning',
            "summary": True,
            "errstring": "{count} {description}(s) without content found. Placeholder text: {para_string} has been inserted."
    	},
    	"95.6_custom_endnote_mark": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "custom_endnote_mark",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'error',
            "summary": True,
            "errstring": "{count} custom Endnote mark(s) found. These are not supported by RSuite and must be replaced with standard Endnote mark(s)."
    	},
    	"95.7_custom_footnote_mark": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "custom_footnote_mark",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'error',
            "summary": True,
            "errstring": "{count} custom Footnote mark(s) found. These are not supported by RSuite and must be replaced with standard Footnote mark(s)."
    	},
    	"95.8_noteref_in_noncontent_pstyle": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "noteref_in_noncontent_pstyle",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'error',
            "errstring": "The reference mark for {descriptionA} is located in a paragraph that is styled with a non-printing paragraph-style: '{descriptionB}'. This will cause RSuite transforms to fail."
    	},
    	"96_deleted_bookmark_summary": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "deleted_objects-bookmarks",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'note',
            "errstring": "{count} bookmark(s) deleted from the manuscript."
    	},
    	"97_deleted_comment_summary": {   # using high digits for "error only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "deleted_objects-comments-comments_xml",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'note',
            "errstring": "{count} comment(s) deleted from the manuscript."
    	},
    	"98_too_many_title_paras": {
            "include_for": ["reporter"],
    		"dict_category_name": "title_paras",
    		"line_template": "",
    		"badnews": 'one_allowed',
            "errstring": "Too many '%s' paragraphs detected ({count}), only one is allowed." % titlestyle
    	},
    	# "98_too_many_title_paras(rsuite_validate)": {
        #     "include_for": ["rsuitevalidate"],
    	# 	"dict_category_name": "title_paras",
    	# 	"line_template": "",
    	# 	"badnews": 'one_allowed',
        #     "badnews_type": 'warning',
        #     "errstring": "Too many '%s' paragraphs detected ({count}), only one is allowed." % titlestyle
    	# },
    	"98.1_too_many_certain_section_para": {
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "too_many_section_para",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'error',
            "errstring": "Too many '{descriptionA}' paragraphs detected: {descriptionB} were found, only one is allowed."
    	},
    	"98.2_too_many_mainhead_para_per_section": {
            "include_for": ["rsuitevalidate"],
    		"dict_category_name": "too_many_heading_para__sort_by_index",
    		"line_template": "",
    		"badnews": 'any',
            "badnews_type": 'error',
            "errstring": "{descriptionB} '{descriptionA}' paragraphs found in {parent_section_start_type}: {parent_section_start_content}. Only one '{descriptionA}' paragraph is allowed per section."
    	},
    	"99_no_logo_paras": {
            "include_for": ["reporter"],
    		"dict_category_name": "logo_paras",
    		"line_template": "",
    		"suggested": True,
            "errstring": "No styled '{}' line detected. If you would like a logo included on your titlepage, please add this style.".format(logostyle)
    	}
    }
    # print report_recipe
    return report_recipe

# #---------------------  MAIN
# # only run if this script is being invoked directly (for testing)
if __name__ == '___':
    # hardcoding values, just for testing
    titlestyle = "Titlepage Book Title (tit)"
    logostyle = "Titlepage Book Title (tit)"
    isbnstyle = "span ISBN (isbn)"
    authorstyle = "Titlepage Author Name (au)"

    report_recipe = getReportRecipe(titlestyle, authorstyle, isbnstyle)
    print (report_recipe["metadata_heading"])
