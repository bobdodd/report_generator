from report_styling import format_table_text

def count_descendants(element):
    """Count the total number of descendant elements"""
    count = 0
    if element and 'children' in element:
        count += len(element['children'])
        for child in element['children']:
            count += count_descendants(child)
    return count

def count_element_type(element, tag_name):
    """Count elements of a specific tag type within a parent element"""
    count = 0
    if element:
        if element.get('tag', '').lower() == tag_name.lower():
            count += 1
        if 'children' in element:
            for child in element['children']:
                count += count_element_type(child, tag_name)
    return count

def element_contains_tag(element, tag_name):
    """Check if element contains a specific tag anywhere in its descendants"""
    if element:
        if element.get('tag', '').lower() == tag_name.lower():
            return True
        if 'children' in element:
            for child in element['children']:
                if element_contains_tag(child, tag_name):
                    return True
    return False

def add_structure_summary_section(doc, db_connection, total_domains):
    """Add the Page Structure section to the summary findings"""
    h2 = doc.add_heading('Page Structure Analysis', level=2)
    h2.style = doc.styles['Heading 2']

    # Add explanation
    doc.add_paragraph("""
    Understanding the structure of web pages is fundamental to accessibility. This section analyzes the common elements found across pages, such as headers, footers, and navigation components. Consistent structure helps users understand and navigate content efficiently.
    """.strip())

    # Query for structure analysis results
    structure_analysis = list(db_connection.db.structure_analysis.find(
        {},
        {"_id": 0}
    ).sort("timestamp", -1).limit(1))

    if structure_analysis:
        analysis = structure_analysis[0]
        
        # Get the overall summary and domain analyses
        overall_summary = analysis.get('overall_summary', {})
        domain_analyses = analysis.get('domain_analyses', {})
        
        # Calculate total pages from domain analyses if not in overall summary
        total_pages = overall_summary.get('total_pages', 0)
        if total_pages == 0 and domain_analyses:
            total_pages = sum(domain_data.get('page_count', 0) for domain_data in domain_analyses.values())
        
        # Add summary statistics
        doc.add_paragraph()
        doc.add_paragraph("Structure Analysis Summary:", style='Normal')
        
        stats_table = doc.add_table(rows=6, cols=2)  # Expanded to include main and complementary content
        stats_table.style = 'Table Grid'
        
        # Add summary data using overall_summary
        rows = stats_table.rows
        rows[0].cells[0].text = "Pages Analyzed"
        rows[0].cells[1].text = str(total_pages)
        
        rows[1].cells[0].text = "Header Consistency"
        header_score = overall_summary.get('average_header_score', 0) * 100
        rows[1].cells[1].text = f"{header_score:.1f}%"
        
        rows[2].cells[0].text = "Footer Consistency"
        footer_score = overall_summary.get('average_footer_score', 0) * 100
        rows[2].cells[1].text = f"{footer_score:.1f}%"
        
        rows[3].cells[0].text = "Navigation Consistency"
        nav_score = overall_summary.get('average_navigation_score', 0) * 100
        rows[3].cells[1].text = f"{nav_score:.1f}%"
        
        rows[4].cells[0].text = "Main Content Consistency"
        main_score = overall_summary.get('average_main_content_score', 0) * 100
        rows[4].cells[1].text = f"{main_score:.1f}%"
        
        rows[5].cells[0].text = "Complementary Content Consistency"
        comp_score = overall_summary.get('average_complementary_score', 0) * 100
        rows[5].cells[1].text = f"{comp_score:.1f}%"
        
        format_table_text(stats_table)
    else:
        doc.add_paragraph("""
    No structure analysis data was found. Please ensure the page structure analysis test is properly integrated and the analysis has been run after testing.
        """.strip())
        