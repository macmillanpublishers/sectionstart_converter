######### IMPORT PY LIBRARIES
import logging

# # initialize logger
logger = logging.getLogger(__name__)


# #---------------------  METHODS
# This method defines what goes in the StyleReport txt and mail outputs, in what order, + formatting.
# See the commented "SAMPLE RECIPE ENTRY" below for details on each field.  All fields should be optional,
#   though text, title or dict_category_name must be present for something to print
def getReportRecipe(titlestyle, authorstyle, isbnstyle):
    report_recipe = {
        # # # # # # # # # # # # #  SAMPLE RECIPE ENTRY:
        # # # # # # # # # # # # # # # # # # # # # # # # # #
    	# "00_example": {                            # < The leading digits are to set the order of appearance on the report
        #   "exclude_from": ["reporter_main"],       # < Add the parent script basename to this list if you want to suppress
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
        #   "badnews": True,        # < If you want each entry from this report_dict category in the Error List,mark this True
        #   "errstring": "No paragraphs."   # < The base string you want used to appear in the report's Error list
        #   "alternate_content": {          # < If you want an alternate title or text element to appear when
        #       "title": "TEST FAIL"        #   report_dict category is empty or not present, set them here. If you
        #       "text": "No title paras."   #   set text but not title, the original title value will be used, and vice versa
        #    }
    	# },
    	"01_metadata_heading": {
            # "exclude_from": ["generate_report"],      # debug test, remove
    		"title": "METADATA",
    		"text": "If any of the information below is wrong, please fix the associated styles in the manuscript."
    	},
    	"02_metadata_title": {
            # "exclude_from": ["reporter_main"],      # debug test, remove
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
    	"05_illustration_holders": {
    		"title": "ILLUSTRATION LIST",
    		"text": "Verify that this list of illustrations includes only the filenames of your illustrations.\n",
    		"dict_category_name": "illustration_holders",
    		"line_template": "{para_string}\n    -located in {parent_section_start_type}: {parent_section_start_content}. (Paragraph {para_index})",
            "alternate_content": {
                "text": "no illustrations detected."
            }
    	},
    	"06_section_start_list": {
    		"title": "SECTIONS FOUND",
    		"text": "",#"Here is a list of all sections detected in your manuscript",
    		"dict_category_name": "section_start_found",
    		"line_template": "{parent_section_start_type:.<33} {parent_section_start_content:57}",
    		"required": True,
            "errstring": "No sections found. Sections must be styled with Section-Start styles.",
            "alternate_content": {
                "text": "No sections detected."
            }
    	},
    	"07_macmillan_style_1st_use": {
    		"title": "MACMILLAN STYLES IN USE",
    		# "text": "{:^48} {:_^50}".format("\033[4mstyles_in_order_of_appearance\033[0m","styled_content_excerpt_from_first_use"),
    		"text": "{:_^40} {:_^40}".format("styles-in_order_of_first_use","excerpt_from_first_use"),
    		"dict_category_name": "Macmillan_style_first_use",
    		"line_template": "{description:.<40} {para_string:50}",
    		"required": True,
            "errstring": "No Macmillan styled paragraphs were found in the manuscript.",
            "alternate_content": {
                "text": "No Macmillan styles detected."
            }
        },       #
    	"08_macmillan_character_style_1st_use": {
    		"text": "{:_^40}".format("character_styles_in_use"),
    		"dict_category_name": "Macmillan_charstyle_first_use",
    		"line_template": "{description}",
    		"required": "n-a",
            "alternate_content": {
                "text": "{:_^40}\nNo character styles detected.".format("character_styles_in_use")
            }
    	},
    	"90_non_macmillan_styles": {   # using high digits for "errror only" items; since they're order agnostic & we may have to renumber the others
    		"dict_category_name": "non-Macmillan_style_used",
    		"line_template": "",
    		"badnews": True,
            "errstring": "Non-Macmillan style '{description}' in {parent_section_start_type}: {parent_section_start_content}. (Paragraph {para_index})"
    	},
    	"91_non_bookmaker_style": {
    		"dict_category_name": "non_bookmaker_macmillan_style",
    		"line_template": "",
    		"badnews": True,
            "errstring": "Non-Bookmaker style: '{description}' in {parent_section_start_type}: {parent_section_start_content}. (Paragraph {para_index})"
    	},
    	"92_empty_section_start_para": {
    		"dict_category_name": "empty_section_start_para",
    		"line_template": "",
    		"badnews": True,
            "errstring": "Empty Section-Start paragraph: found a '{description}' para with no text. (Paragraph {para_index})"
    	}
    }
    # print report_recipe
    return report_recipe

# #---------------------  MAIN
# # only run if this script is being invoked directly (for testing)
if __name__ == '__main__':
    # hardcoding values, just for testing
    titlestyle = "Titlepage Book Title (tit)"
    isbnstyle = "span ISBN (isbn)"
    authorstyle = "Titlepage Author Name (au)"

    report_recipe = getReportRecipe(titlestyle, authorstyle, isbnstyle)
    print report_recipe["metadata_heading"]
