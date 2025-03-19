from report_styling import add_table_of_contents, format_toc_styles

def add_toc_section(doc):
    """Add the table of contents section"""
    toc_heading = doc.add_heading('Table of Contents', level=1)
    toc_heading.style = doc.styles['Heading 1']
    add_table_of_contents(doc)
    format_toc_styles(doc)
    