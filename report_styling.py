# report_styling.py
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn

def set_document_styles(doc):
    """Set up document styles with specified font sizes and spacing"""
    # Default paragraph text
    style = doc.styles['Normal']
    style.font.size = Pt(14)
    style.font.name = 'Arial'
    
    # Heading Styles
    title_style = doc.styles['Title']  # Level 0
    title_style.font.size = Pt(26)
    title_style.font.name = 'Arial'
    title_style.font.bold = True
    title_style.paragraph_format.space_after = Pt(14)  # One line height
    
    h1_style = doc.styles['Heading 1']
    h1_style.font.size = Pt(22)
    h1_style.font.name = 'Arial'
    h1_style.font.bold = True
    h1_style.paragraph_format.space_after = Pt(14)  # One line height
    
    h2_style = doc.styles['Heading 2']
    h2_style.font.size = Pt(18)
    h2_style.font.name = 'Arial'
    h2_style.font.bold = True
    h2_style.paragraph_format.space_after = Pt(14)  # One line height
    
    h3_style = doc.styles['Heading 3']
    h3_style.font.size = Pt(16)
    h3_style.font.name = 'Arial'
    h3_style.font.bold = True
    h3_style.paragraph_format.space_after = Pt(14)  # One line height
    
    # List styles
    list_style = doc.styles['List Bullet']
    list_style.font.size = Pt(14)
    list_style.font.name = 'Arial'
    
    # Caption style
    caption_style = doc.styles['Caption']
    caption_style.font.size = Pt(14)
    caption_style.font.name = 'Arial'
    caption_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # TOC styles are automatically created when the TOC is inserted

def add_table_of_contents(doc):
    """Add a table of contents to the document"""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')

    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    # Add hyperlink and web options to the TOC field code
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u \\w'  # \h adds hyperlinks, \w preserves tab spacing

    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

def format_toc_styles(doc):
    """Format TOC styles to properly handle numbered headings"""
    for i in range(1, 4):  # TOC levels 1-3
        style_name = f'TOC {i}'
        if style_name in doc.styles:
            style = doc.styles[style_name]
            style.paragraph_format.tab_stops.clear_all()
            
            # Calculate indentation based on level
            base_indent = Inches(0.25)
            level_indent = Inches(0.25 * (i-1))
            
            # Set paragraph indentation
            style.paragraph_format.left_indent = level_indent
            style.paragraph_format.first_line_indent = -base_indent
            
            # Add tab stops:
            # First tab for after the number
            style.paragraph_format.tab_stops.add_tab_stop(level_indent, WD_TAB_ALIGNMENT.LEFT)
            # Final tab for page number
            style.paragraph_format.tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
            
            # Set font
            style.font.name = 'Arial'
            style.font.size = Pt(14)

def format_table_text(table):
    """Format text within tables to ensure consistent font and size"""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(14)
                    run.font.name = 'Arial'

def create_element(name):
    """Create an XML element with the specified name"""
    return OxmlElement(name)

def create_attribute(element, name, value):
    """Add an attribute to an XML element"""
    element.set(qn(name), value)

def add_page_number(paragraph):
    """Add a page number field to a paragraph"""
    run = paragraph.add_run()
    fldChar1 = create_element('w:fldChar')
    create_attribute(fldChar1, 'w:fldCharType', 'begin')

    instrText = create_element('w:instrText')
    create_attribute(instrText, 'xml:space', 'preserve')
    instrText.text = "PAGE"

    fldChar2 = create_element('w:fldChar')
    create_attribute(fldChar2, 'w:fldCharType', 'end')

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)