# report_generator.py
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from datetime import datetime
import os

# Import styling utilities
from report_styling import (
    set_document_styles, 
    format_table_text
)

# Import section generators
from sections.sections_header import setup_document_header_footer
from sections.title_page import add_title_page
from sections.table_of_contents import add_toc_section
from sections.executive_summary import add_executive_summary

# Import summary findings sections
from sections.summary_findings.accessible_names import add_accessible_names_section
from sections.summary_findings.animation import add_animation_section
from sections.summary_findings.color_contrast import add_color_contrast_section
from sections.summary_findings.color_as_indicator import add_color_as_indicator_section
from sections.summary_findings.dialogs import add_dialogs_section
from sections.summary_findings.event_handling import add_event_handling_section
from sections.summary_findings.floating_dialogs import add_floating_dialogs_section
from sections.summary_findings.focus_management import add_focus_management_section
from sections.summary_findings.forms import add_forms_section
from sections.summary_findings.headings import add_headings_section
from sections.summary_findings.images import add_images_section
from sections.summary_findings.landmarks import add_landmarks_section
from sections.summary_findings.language import add_language_section
from sections.summary_findings.lists import add_lists_section
from sections.summary_findings.maps import add_maps_section
from sections.summary_findings.media_queries import add_media_queries_section
from sections.summary_findings.menus import add_menus_section
from sections.summary_findings.more_controls import add_more_controls_section
from sections.summary_findings.responsive_accessibility import add_responsive_accessibility_summary
from sections.summary_findings.structure import add_structure_summary_section
from sections.summary_findings.tabindex import add_tabindex_section
from sections.summary_findings.title_attribute import add_title_attribute_section
from sections.summary_findings.tables import add_tables_section
from sections.summary_findings.timers import add_timers_section
from sections.summary_findings.videos import add_videos_section

# Import detailed findings sections
from sections.detailed_findings.accessible_names import add_detailed_accessible_names
from sections.detailed_findings.animation import add_detailed_animation
from sections.detailed_findings.color_contrast import add_detailed_color_contrast
from sections.detailed_findings.color_as_indicator import add_detailed_color_as_indicator
from sections.detailed_findings.dialogs import add_detailed_dialogs
from sections.detailed_findings.event_handling import add_detailed_event_handling
from sections.detailed_findings.forms import add_detailed_forms
from sections.detailed_findings.headings import add_detailed_headings
from sections.detailed_findings.images import add_detailed_images
from sections.detailed_findings.landmarks import add_detailed_landmarks
from sections.detailed_findings.language import add_detailed_language
from sections.detailed_findings.lists import add_detailed_lists
from sections.detailed_findings.media_queries import add_detailed_media_queries
from sections.detailed_findings.responsive_accessibility import add_responsive_accessibility_detailed
from sections.detailed_findings.structure import add_detailed_structure
# Import the new detailed sections
from sections.detailed_findings.maps import add_detailed_maps
from sections.detailed_findings.menus import add_detailed_menus
from sections.detailed_findings.more_controls import add_detailed_more_controls
from sections.detailed_findings.tabindex import add_detailed_tabindex
from sections.detailed_findings.title_attribute import add_detailed_title_attribute
from sections.detailed_findings.tables import add_detailed_tables
from sections.detailed_findings.timers import add_detailed_timers
from sections.detailed_findings.videos import add_detailed_videos

# Import appendices
from sections.appendices import add_appendices

def create_report_template(db_connection, title, author, date):
    print("Starting report creation...")

    ####################################################
    # Get list of URLs and domains used by the reporting
    # Gets used everywhere
    ####################################################

    all_test_runs = db_connection.get_all_test_runs()
    test_run_ids = [str(run['_id']) for run in all_test_runs]
    all_urls = db_connection.page_results.distinct('url', {'test_run_id': {'$in': test_run_ids}})    

    total_domains = set()
    for url in all_urls:
        domain = url.replace('http://', '').replace('https://', '').split('/')[0]
        total_domains.add(domain)

    doc = Document()
    
    # Set up document styles
    set_document_styles(doc)

    # Set up header and footer
    setup_document_header_footer(doc, title)

    # Add title page
    add_title_page(doc, db_connection, title, author, date)

    # Add table of contents
    doc.add_page_break()
    add_toc_section(doc)

    # Add executive summary
    doc.add_page_break()
    add_executive_summary(doc, db_connection, total_domains)

    #############################################
    # Summary Findings
    #############################################
    h1 = doc.add_heading('Summary findings', level=1)
    h1.style = doc.styles['Heading 1']

    # Add Media Queries Section (first, as it affects overall responsiveness)
    add_media_queries_section(doc, db_connection, total_domains)
    
    # Add Responsive Accessibility Section (right after media queries)
    add_responsive_accessibility_summary(doc, db_connection, total_domains)
    
    # Add Accessible Names Section
    add_accessible_names_section(doc, db_connection, total_domains)
    
    # Add Animation Section
    add_animation_section(doc, db_connection, total_domains)
    
    # Add Color Contrast Section
    add_color_contrast_section(doc, db_connection, total_domains)
    
    # Add Color as Indicator Section
    add_color_as_indicator_section(doc, db_connection, total_domains)
    
    # Add Dialogs Section
    add_dialogs_section(doc, db_connection, total_domains)
    
    # Add Event Handling Section
    add_event_handling_section(doc, db_connection, total_domains)
    
    # Add Floating Dialogs Section
    add_floating_dialogs_section(doc, db_connection, total_domains)
    
    # Add Focus Management Section
    add_focus_management_section(doc, db_connection, total_domains)
    
    # Add Forms Section
    add_forms_section(doc, db_connection, total_domains)
    
    # Add Headings Section
    add_headings_section(doc, db_connection, total_domains)
    
    # Add Images Section
    add_images_section(doc, db_connection, total_domains)
    
    # Add Landmarks Section
    add_landmarks_section(doc, db_connection, total_domains)
    
    # Add Language Section
    add_language_section(doc, db_connection, total_domains)
    
    # Add Lists Section
    add_lists_section(doc, db_connection, total_domains)
    
    # Add Maps Section
    add_maps_section(doc, db_connection, total_domains)
    
    # Add Menus Section
    add_menus_section(doc, db_connection, total_domains)
    
    # Add More Controls Section
    add_more_controls_section(doc, db_connection, total_domains)
    
    # Add Structure Section
    add_structure_summary_section(doc, db_connection, total_domains)
    
    # Add Tabindex Section
    add_tabindex_section(doc, db_connection, total_domains)
    
    # Add Title Attribute Section
    add_title_attribute_section(doc, db_connection, total_domains)
    
    # Add Tables Section
    add_tables_section(doc, db_connection, total_domains)
    
    # Add Timers Section
    add_timers_section(doc, db_connection, total_domains)
    
    # Add Videos Section
    add_videos_section(doc, db_connection, total_domains)

    #############################################
    # Detailed Findings
    #############################################
    doc.add_page_break()
    h1 = doc.add_heading('Detailed findings', level=1)
    h1.style = doc.styles['Heading 1']

    # Add Detailed Media Queries Section (first, as it affects overall responsiveness)
    add_detailed_media_queries(doc, db_connection, total_domains)
    
    # Add Detailed Responsive Accessibility Section (right after media queries)
    add_responsive_accessibility_detailed(doc, db_connection, total_domains)
    
    # Add Detailed Accessible Names Section
    add_detailed_accessible_names(doc, db_connection, total_domains)
    
    # Add Detailed Animation Section
    add_detailed_animation(doc, db_connection, total_domains)
    
    # Add Detailed Color Contrast Section
    add_detailed_color_contrast(doc, db_connection, total_domains)
    
    # Add Detailed Color as Indicator Section
    add_detailed_color_as_indicator(doc, db_connection, total_domains)
    
    # Add Detailed Dialogs Section
    add_detailed_dialogs(doc, db_connection, total_domains)
    
    # Add Detailed Event Handling Section
    add_detailed_event_handling(doc, db_connection, total_domains)
    
    # Add Detailed Forms Section
    add_detailed_forms(doc, db_connection, total_domains)
    
    # Add Detailed Headings Section
    add_detailed_headings(doc, db_connection, total_domains)
    
    # Add Detailed Images Section
    add_detailed_images(doc, db_connection, total_domains)
    
    # Add Detailed Landmarks Section
    add_detailed_landmarks(doc, db_connection, total_domains)
    
    # Add Detailed Language Section
    add_detailed_language(doc, db_connection, total_domains)
    
    # Add Detailed Lists Section
    add_detailed_lists(doc, db_connection, total_domains)
    
    # Add Detailed Structure Section
    add_detailed_structure(doc, db_connection, total_domains)
    
    # Add the new detailed sections 
    # Add Detailed Maps Section
    add_detailed_maps(doc, db_connection, total_domains)
    
    # Add Detailed Menus Section
    add_detailed_menus(doc, db_connection, total_domains)
    
    # Add Detailed More Controls Section
    add_detailed_more_controls(doc, db_connection, total_domains)
    
    # Add Detailed Tabindex Section
    add_detailed_tabindex(doc, db_connection, total_domains)
    
    # Add Detailed Title Attribute Section
    add_detailed_title_attribute(doc, db_connection, total_domains)
    
    # Add Detailed Tables Section
    add_detailed_tables(doc, db_connection, total_domains)
    
    # Add Detailed Timers Section
    add_detailed_timers(doc, db_connection, total_domains)
    
    # Add Detailed Videos Section
    add_detailed_videos(doc, db_connection, total_domains)

    #############################################
    # Appendices
    #############################################
    doc.add_page_break()
    add_appendices(doc, db_connection)

    return doc

def generate_report(db_connection, title, author, date, output_folder):
    try:
        doc = create_report_template(db_connection, title, author, date)
        output_filename = f'{output_folder}/accessibility_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx'
        doc.save(output_filename)
        return output_filename
    except Exception as e:
        print(f"Error generating report: {e}")
        return None