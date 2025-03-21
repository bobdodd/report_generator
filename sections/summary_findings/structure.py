from report_styling import format_table_text
import json

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

    # Query to find latest test results that contain page structure data
    page_results = list(db_connection.page_results.find(
        {"results.accessibility.tests.page_structure": {"$exists": True}},
        {"url": 1, "results.accessibility.tests.page_structure": 1, "_id": 0}
    ).sort("timestamp", -1).limit(50))  # Get latest 50 pages to analyze

    if page_results:
        # Track statistics
        total_analyzed = len(page_results)
        pages_with_header = 0
        pages_with_footer = 0
        pages_with_nav = 0
        pages_with_main = 0
        pages_with_complementary = 0
        
        # Content block stats
        pages_with_hero = 0
        pages_with_cards = 0
        pages_with_features = 0
        pages_with_carousels = 0
        
        # Component stats
        pages_with_search = 0
        pages_with_cookie_notice = 0
        pages_with_popups = 0
        pages_with_forms = 0
        
        # Analyze each page's structure data
        for page in page_results:
            try:
                if 'results' in page and 'accessibility' in page['results'] and 'tests' in page['results']['accessibility'] \
                   and 'page_structure' in page['results']['accessibility']['tests']:
                    
                    structure_data = page['results']['accessibility']['tests']['page_structure']
                    
                    # The structure might be a string (JSON) or already an object
                    if isinstance(structure_data, str):
                        structure_data = json.loads(structure_data)
                    
                    # Check for page_structure data format
                    if isinstance(structure_data, dict) and 'page_structure' in structure_data:
                        page_flags = structure_data['page_structure'].get('pageFlags', {})
                        
                        # Count core structural elements
                        if page_flags.get('hasHeader', False): pages_with_header += 1
                        if page_flags.get('hasFooter', False): pages_with_footer += 1
                        if page_flags.get('hasMainNavigation', False): pages_with_nav += 1
                        if page_flags.get('hasMainContent', False): pages_with_main += 1
                        if page_flags.get('hasComplementaryContent', False): pages_with_complementary += 1
                        
                        # Count content blocks
                        if page_flags.get('hasHeroSection', False): pages_with_hero += 1
                        if page_flags.get('hasCardGrids', False): pages_with_cards += 1
                        if page_flags.get('hasFeatureSections', False): pages_with_features += 1
                        if page_flags.get('hasCarousels', False): pages_with_carousels += 1
                        
                        # Count UI components
                        if page_flags.get('hasSearchComponent', False): pages_with_search += 1
                        if page_flags.get('hasCookieNotice', False): pages_with_cookie_notice += 1
                        if page_flags.get('hasPopups', False): pages_with_popups += 1
                        if page_flags.get('hasForms', False): pages_with_forms += 1
            except Exception as e:
                # Skip problematic entries
                print(f"Error processing page structure data: {e}")
                continue
        
        # Create structure summary table
        doc.add_paragraph()
        doc.add_paragraph("Core Structure Elements:", style='Normal')
        
        stats_table = doc.add_table(rows=6, cols=3)
        stats_table.style = 'Table Grid'
        
        # Headers
        header_row = stats_table.rows[0].cells
        header_row[0].text = "Component"
        header_row[1].text = "Pages"
        header_row[2].text = "Percentage"
        
        # Structure data
        components = [
            ("Header", pages_with_header),
            ("Footer", pages_with_footer),
            ("Navigation", pages_with_nav),
            ("Main Content", pages_with_main),
            ("Complementary Content", pages_with_complementary)
        ]
        
        for i, (component, count) in enumerate(components, 1):
            row = stats_table.rows[i].cells
            row[0].text = component
            row[1].text = str(count)
            percentage = (count / total_analyzed) * 100 if total_analyzed > 0 else 0
            row[2].text = f"{percentage:.1f}%"
        
        format_table_text(stats_table)
        
        # Create content blocks table
        doc.add_paragraph()
        doc.add_paragraph("Common Content Blocks:", style='Normal')
        
        blocks_table = doc.add_table(rows=5, cols=3)
        blocks_table.style = 'Table Grid'
        
        # Headers
        header_row = blocks_table.rows[0].cells
        header_row[0].text = "Content Block"
        header_row[1].text = "Pages"
        header_row[2].text = "Percentage"
        
        # Content blocks data
        blocks = [
            ("Hero Sections", pages_with_hero),
            ("Card Grids", pages_with_cards),
            ("Feature Sections", pages_with_features),
            ("Carousels/Sliders", pages_with_carousels)
        ]
        
        for i, (block, count) in enumerate(blocks, 1):
            row = blocks_table.rows[i].cells
            row[0].text = block
            row[1].text = str(count)
            percentage = (count / total_analyzed) * 100 if total_analyzed > 0 else 0
            row[2].text = f"{percentage:.1f}%"
        
        format_table_text(blocks_table)
        
        # Add note about detection methodology
        doc.add_paragraph()
        doc.add_paragraph("""
        Note: Elements are detected using multiple strategies, including semantic tags, ARIA roles, class/ID naming patterns, and positional analysis. The percentages above reflect the presence of clearly identifiable structural elements.
        """.strip())
        
    else:
        doc.add_paragraph("""
    No structure analysis data was found. Please ensure the page structure analysis test is properly integrated and the analysis has been run after testing.
        """.strip())
        