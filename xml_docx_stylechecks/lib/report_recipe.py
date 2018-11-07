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
def getReportRecipe(titlestyle, authorstyle, isbnstyle, logostyle):
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
    	# 	"required": True,       # < If it's an error if a report_dict category is empty or not present, mark this True
    	# 	                        #   (If this is true you will need an "errstring" entry too)
        #   "v_warning_banner": "Alert string",   # this is for validator only scripts - if any edits were made or unsupported styles were found, we want to surface
        #                                       a different banner on the report output. Including this key=True signals that we want that warning.
        #   "badnews": 'any',        # < If you want any entry from this report_dict category in the Error List,mark this True.. if one entry is ok but more are errors, use value 'one_allowed'
        #   "errstring": "No paragraphs."   # < The base string you want used to appear in the report's Error list
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
    	"05_illustration_holders": {
            "include_for": ["reporter", "validator", "rsuitevalidate"],
    		"title": "ILLUSTRATION LIST",
    		"text": "Verify that this list of illustrations includes only the filenames of your illustrations.\n",
    		"dict_category_name": "illustration_holders__sort_by_index",
    		"line_template": "{description}\n    -located in {parent_section_start_type}: {parent_section_start_content}.",# (Paragraph {para_index})",
            "alternate_content": {
                "text": "no illustrations detected."
            }
    	},
    	"06_section_start_list": {
            "include_for": ["reporter", "validator"],
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
            "include_for": ["reporter"],
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
            "include_for": ["reporter"],
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
    		"title": "NON-MACMILLAN STYLES",
    		"text": "Non-Macmillan styles detected.\nContent styled with non-Macmillan styles may not appear properly-styled in your egalley.\n",
    		"dict_category_name": "non-Macmillan_style_used",
    		"line_template": "- style '{description}': found in section '{parent_section_start_type}': {parent_section_start_content}",
    		"required": "n-a",
            "v_warning_banner": "Unsupported (non-Macmillan) style(s) found."
    	},
    	"27_non-Bookmaker_styles(validator)": {
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
    		"dict_category_name": "deleted_shapes_and_sectionbreaks",
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
    	"89_non_section_start_styled_firstpara": {   # using high digits for "errror only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["reporter", "validator"],
    		"dict_category_name": "non_section_start_styled_firstpara",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "First paragraph of document styled with non-Section Start style ('{description}')."# (Paragraph {para_index})"
    	},
    	"90_non_macmillan_styles": {   # using high digits for "errror only" items; since they're order agnostic & we may have to renumber the others
            "include_for": ["reporter"],
    		"dict_category_name": "non-Macmillan_style_used",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Non-Macmillan style '{description}' in {parent_section_start_type}: {parent_section_start_content}."# (Paragraph {para_index})"
    	},
    	"91_non_bookmaker_style": {
            "include_for": ["reporter"],
    		"dict_category_name": "non_bookmaker_macmillan_style",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Non-Bookmaker style: '{description}' in {parent_section_start_type}: {parent_section_start_content}."# (Paragraph {para_index})"
    	},
    	"92_empty_section_start_para": {
            "include_for": ["reporter"],
    		"dict_category_name": "empty_section_start_para",
    		"line_template": "",
    		"badnews": 'any',
            "errstring": "Empty Section-Start paragraph: found a '{description}' para with no text."# (Paragraph {para_index})"
    	},
    	"93_too_many_title_paras": {
            "include_for": ["reporter"],
    		"dict_category_name": "title_paras",
    		"line_template": "",
    		"badnews": 'one_allowed',
            "errstring": "Too many '{}' paragraphs detected, only one is allowed.".format(titlestyle)
    	},
    	"94_no_logo_paras": {
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
    print report_recipe["metadata_heading"]
