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

def add_detailed_structure(doc, db_connection, total_domains):
    """Add the detailed Page Structure section"""
    doc.add_page_break()
    h2 = doc.add_heading('Page Structure Analysis', level=2)
    h2.style = doc.styles['Heading 2']

    # Add explanation with WCAG reference
    doc.add_paragraph("""
    Understanding the structure of web pages is fundamental to accessibility. This section analyzes the common elements found across pages, such as headers, footers, and navigation components. Consistent structure helps users understand and navigate content efficiently.
    
    WCAG Success Criteria 1.3.1 (Info and Relationships) requires that information, structure, and relationships conveyed through presentation can be programmatically determined. Additionally, WCAG 2.4.1 (Bypass Blocks) requires a mechanism to bypass blocks of content that are repeated on multiple pages.
    """.strip())

    # Query to find recent pages with page structure data
    page_results = list(db_connection.page_results.find(
        {"results.accessibility.tests.page_structure": {"$exists": True}},
        {"url": 1, "results.accessibility.tests.page_structure": 1, "_id": 0}
    ).sort("timestamp", -1).limit(30))  # Get latest 30 pages to analyze in detail

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
        
        # Detection methods - analyze how components were detected
        header_detection_methods = {}
        footer_detection_methods = {}
        nav_detection_methods = {}
        main_detection_methods = {}
        
        # Sample data
        sample_header = None
        sample_footer = None
        sample_nav = None
        sample_main = None
        sample_content_blocks = None

        # Analyze each page's structure data
        for page in page_results:
            try:
                url = page.get('url')
                if 'results' in page and 'accessibility' in page['results'] and 'tests' in page['results']['accessibility'] \
                   and 'page_structure' in page['results']['accessibility']['tests']:
                    
                    structure_data = page['results']['accessibility']['tests']['page_structure']
                    
                    # The structure might be a string (JSON) or already an object
                    if isinstance(structure_data, str):
                        structure_data = json.loads(structure_data)
                    
                    # Check for page_structure data format
                    if isinstance(structure_data, dict) and 'page_structure' in structure_data:
                        # Get the main page structure object
                        page_struct = structure_data['page_structure']
                        
                        # Get page flags
                        page_flags = page_struct.get('pageFlags', {})
                        
                        # Count core structural elements
                        has_header = page_flags.get('hasHeader', False)
                        has_footer = page_flags.get('hasFooter', False)
                        has_nav = page_flags.get('hasMainNavigation', False)
                        has_main = page_flags.get('hasMainContent', False)
                        has_complementary = page_flags.get('hasComplementaryContent', False)
                        
                        if has_header: pages_with_header += 1
                        if has_footer: pages_with_footer += 1
                        if has_nav: pages_with_nav += 1
                        if has_main: pages_with_main += 1
                        if has_complementary: pages_with_complementary += 1
                        
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
                        
                        # Track detection methods
                        structure_summary = page_struct.get('summary', {})
                        
                        # Header detection methods
                        if has_header and 'header' in structure_summary:
                            header_types = structure_summary['header'].get('types', [])
                            for method in header_types:
                                header_detection_methods[method] = header_detection_methods.get(method, 0) + 1
                        
                        # Footer detection methods
                        if has_footer and 'footer' in structure_summary:
                            footer_types = structure_summary['footer'].get('types', [])
                            for method in footer_types:
                                footer_detection_methods[method] = footer_detection_methods.get(method, 0) + 1
                        
                        # Navigation detection methods
                        if has_nav and 'navigation' in structure_summary:
                            nav_types = structure_summary['navigation'].get('types', [])
                            for method in nav_types:
                                nav_detection_methods[method] = nav_detection_methods.get(method, 0) + 1
                        
                        # Main content detection methods
                        if has_main and 'mainContent' in structure_summary:
                            main_types = structure_summary['mainContent'].get('types', [])
                            for method in main_types:
                                main_detection_methods[method] = main_detection_methods.get(method, 0) + 1
                        
                        # Store sample data for detailed analysis 
                        # (from the first page that has good data)
                        if not sample_header and has_header:
                            sample_header = page_struct.get('keyElements', {}).get('primaryHeader')
                            
                        if not sample_footer and has_footer:
                            sample_footer = page_struct.get('keyElements', {}).get('primaryFooter')
                            
                        if not sample_nav and has_nav and 'navigation' in structure_summary:
                            nav_details = structure_summary.get('navigation', {})
                            if nav_details.get('found', False):
                                sample_nav = page_struct.get('keyElements', {}).get('navigation')
                        
                        if not sample_main and has_main:
                            sample_main = page_struct.get('keyElements', {}).get('mainContent')
                        
                        if not sample_content_blocks and 'contentBlocks' in structure_summary:
                            sample_content_blocks = page_struct.get('fullStructure', {}).get('commonContentBlocks')
            except Exception as e:
                # Skip problematic entries
                print(f"Error processing page structure data for {url}: {e}")
                continue
        
        # Create the overview summary table
        doc.add_paragraph()
        doc.add_paragraph("Structure Analysis Overview:", style='Normal')
        
        overview_table = doc.add_table(rows=8, cols=3)
        overview_table.style = 'Table Grid'
        
        # Headers
        header_row = overview_table.rows[0].cells
        header_row[0].text = "Structure Element"
        header_row[1].text = "Pages"
        header_row[2].text = "Detection Method"
        
        # Structure data
        components = [
            ("Header", pages_with_header, get_top_detection_method(header_detection_methods)),
            ("Footer", pages_with_footer, get_top_detection_method(footer_detection_methods)),
            ("Navigation", pages_with_nav, get_top_detection_method(nav_detection_methods)),
            ("Main Content", pages_with_main, get_top_detection_method(main_detection_methods)),
            ("Complementary Content", pages_with_complementary, "Various methods"),
            ("Forms", pages_with_forms, "Form element detection"),
            ("UI Components", pages_with_search, "Search, popups, and notices")
        ]
        
        for i, (component, count, method) in enumerate(components, 1):
            row = overview_table.rows[i].cells
            row[0].text = component
            row[1].text = f"{count}/{total_analyzed} ({(count/total_analyzed)*100:.1f}%)"
            row[2].text = method
        
        format_table_text(overview_table)
        
        # UI Components Detail Table
        doc.add_paragraph()
        doc.add_paragraph("UI Components Breakdown:", style='Normal')
        
        ui_table = doc.add_table(rows=5, cols=2)
        ui_table.style = 'Table Grid'
        
        # Headers
        ui_headers = ui_table.rows[0].cells
        ui_headers[0].text = "Component Type"
        ui_headers[1].text = "Prevalence"
        
        # UI Component data
        ui_components = [
            ("Search Components", f"{pages_with_search}/{total_analyzed} ({(pages_with_search/total_analyzed)*100:.1f}%)"),
            ("Cookie Notices", f"{pages_with_cookie_notice}/{total_analyzed} ({(pages_with_cookie_notice/total_analyzed)*100:.1f}%)"),
            ("Popups/Modals", f"{pages_with_popups}/{total_analyzed} ({(pages_with_popups/total_analyzed)*100:.1f}%)"),
            ("Form Components", f"{pages_with_forms}/{total_analyzed} ({(pages_with_forms/total_analyzed)*100:.1f}%)")
        ]
        
        for i, (component, count) in enumerate(ui_components, 1):
            row = ui_table.rows[i].cells
            row[0].text = component
            row[1].text = count
        
        format_table_text(ui_table)
        
        # Header Analysis
        doc.add_paragraph()
        h3 = doc.add_heading('Header Analysis', level=3)
        h3.style = doc.styles['Heading 3']
        
        doc.add_paragraph("""
        Headers are crucial for establishing site identity and providing primary navigation. Well-structured headers should use semantic HTML (<header> element) with appropriate ARIA roles.
        """.strip())
        
        # Create detection methods table for headers
        doc.add_paragraph("Header Detection Methods:", style='Normal')
        create_detection_methods_table(doc, header_detection_methods, total_analyzed)
        
        # Header sample analysis (if available)
        if sample_header:
            doc.add_paragraph("Header Structure Analysis:", style='Normal')
            
            # Extract header details
            tag = sample_header.get('tag', 'Unknown')
            roles = sample_header.get('role', 'None')
            classes = sample_header.get('classArray', [])
            classes_str = ', '.join(classes) if classes else 'None'
            
            doc.add_paragraph(f"Tag: <{tag}>", style='List Bullet')
            doc.add_paragraph(f"ARIA Role: {roles}", style='List Bullet')
            doc.add_paragraph(f"Common CSS Classes: {classes_str}", style='List Bullet')
            
            # Get complexity info if available
            children = sample_header.get('children', [])
            doc.add_paragraph(f"Direct children: {len(children)}", style='List Bullet')
            
            # Check if header is fixed or sticky
            position = sample_header.get('position', 'static')
            doc.add_paragraph(f"Position style: {position}", style='List Bullet')
            
            # Show common child elements
            child_tags = {}
            for child in children:
                child_tag = child.get('tag', 'unknown')
                child_tags[child_tag] = child_tags.get(child_tag, 0) + 1
            
            if child_tags:
                doc.add_paragraph("Common child elements:")
                for tag, count in sorted(child_tags.items(), key=lambda x: x[1], reverse=True):
                    if count > 0:
                        doc.add_paragraph(f"<{tag}>: {count}", style='List Bullet')
        else:
            doc.add_paragraph("No detailed header analysis available.")
        
        # Footer Analysis
        doc.add_paragraph()
        h3 = doc.add_heading('Footer Analysis', level=3)
        h3.style = doc.styles['Heading 3']
        
        doc.add_paragraph("""
        Footers typically contain supplementary navigation, copyright information, and secondary links. Well-structured footers should use semantic HTML (<footer> element) with appropriate ARIA roles.
        """.strip())
        
        # Create detection methods table for footers
        doc.add_paragraph("Footer Detection Methods:", style='Normal')
        create_detection_methods_table(doc, footer_detection_methods, total_analyzed)
        
        # Footer sample analysis (if available)
        if sample_footer:
            doc.add_paragraph("Footer Structure Analysis:", style='Normal')
            
            # Extract footer details
            tag = sample_footer.get('tag', 'Unknown')
            roles = sample_footer.get('role', 'None')
            classes = sample_footer.get('classArray', [])
            classes_str = ', '.join(classes) if classes else 'None'
            
            doc.add_paragraph(f"Tag: <{tag}>", style='List Bullet')
            doc.add_paragraph(f"ARIA Role: {roles}", style='List Bullet')
            doc.add_paragraph(f"Common CSS Classes: {classes_str}", style='List Bullet')
            
            # Get complexity info if available
            children = sample_footer.get('children', [])
            doc.add_paragraph(f"Direct children: {len(children)}", style='List Bullet')
            
            # Show common child elements
            child_tags = {}
            for child in children:
                child_tag = child.get('tag', 'unknown')
                child_tags[child_tag] = child_tags.get(child_tag, 0) + 1
            
            if child_tags:
                doc.add_paragraph("Common child elements:")
                for tag, count in sorted(child_tags.items(), key=lambda x: x[1], reverse=True):
                    if count > 0:
                        doc.add_paragraph(f"<{tag}>: {count}", style='List Bullet')
        else:
            doc.add_paragraph("No detailed footer analysis available.")
        
        # Navigation Analysis
        doc.add_paragraph()
        h3 = doc.add_heading('Navigation Analysis', level=3)
        h3.style = doc.styles['Heading 3']
        
        doc.add_paragraph("""
        Navigation is essential for users to move around a site. Well-structured navigation should use semantic HTML (<nav> element) with appropriate ARIA roles and be keyboard accessible.
        """.strip())
        
        # Create detection methods table for navigation
        doc.add_paragraph("Navigation Detection Methods:", style='Normal')
        create_detection_methods_table(doc, nav_detection_methods, total_analyzed)
        
        # Navigation sample analysis (if available)
        if sample_nav:
            doc.add_paragraph("Navigation Structure Analysis:", style='Normal')
            
            # Extract navigation details
            tag = sample_nav.get('tag', 'Unknown')
            roles = sample_nav.get('role', 'None')
            classes = sample_nav.get('classArray', [])
            classes_str = ', '.join(classes) if classes else 'None'
            
            doc.add_paragraph(f"Tag: <{tag}>", style='List Bullet')
            doc.add_paragraph(f"ARIA Role: {roles}", style='List Bullet')
            doc.add_paragraph(f"Common CSS Classes: {classes_str}", style='List Bullet')
            
            # Count links and list items
            link_count = count_element_type(sample_nav, 'a')
            list_item_count = count_element_type(sample_nav, 'li')
            
            doc.add_paragraph(f"Links: {link_count}", style='List Bullet')
            doc.add_paragraph(f"List items: {list_item_count}", style='List Bullet')
            
            # Is the navigation within a list structure?
            has_ul = element_contains_tag(sample_nav, 'ul')
            has_ol = element_contains_tag(sample_nav, 'ol')
            
            doc.add_paragraph(f"Uses list structure: {'Yes' if (has_ul or has_ol) else 'No'}", style='List Bullet')
        else:
            doc.add_paragraph("No detailed navigation analysis available.")
        
        # Content Blocks Analysis
        doc.add_paragraph()
        h3 = doc.add_heading('Content Blocks Analysis', level=3)
        h3.style = doc.styles['Heading 3']
        
        doc.add_paragraph("""
        Modern websites typically use common content block patterns. These patterns should maintain accessibility regardless of their visual presentation.
        """.strip())
        
        # Content blocks table
        content_blocks_table = doc.add_table(rows=5, cols=2)
        content_blocks_table.style = 'Table Grid'
        
        # Headers
        cb_header = content_blocks_table.rows[0].cells
        cb_header[0].text = "Content Block Type"
        cb_header[1].text = "Prevalence"
        
        # Content blocks data
        cb_data = [
            ("Hero Sections", f"{pages_with_hero}/{total_analyzed} ({(pages_with_hero/total_analyzed)*100:.1f}%)"),
            ("Card Grids/Layouts", f"{pages_with_cards}/{total_analyzed} ({(pages_with_cards/total_analyzed)*100:.1f}%)"),
            ("Feature Sections", f"{pages_with_features}/{total_analyzed} ({(pages_with_features/total_analyzed)*100:.1f}%)"),
            ("Carousels/Sliders", f"{pages_with_carousels}/{total_analyzed} ({(pages_with_carousels/total_analyzed)*100:.1f}%)")
        ]
        
        for i, (block_type, prevalence) in enumerate(cb_data, 1):
            row = content_blocks_table.rows[i].cells
            row[0].text = block_type
            row[1].text = prevalence
        
        format_table_text(content_blocks_table)
        
        # Sample content blocks analysis
        if sample_content_blocks:
            doc.add_paragraph()
            doc.add_paragraph("Common Patterns in Content Blocks:", style='Normal')
            
            # Analyze carousels
            if 'carousels' in sample_content_blocks and sample_content_blocks['carousels']:
                carousel_example = sample_content_blocks['carousels'][0]
                doc.add_paragraph("Carousel/Slider Structure:", style='List Bullet')
                
                carousel_details = carousel_example.get('details', {})
                
                # Extract class/id patterns
                carousel_id = carousel_details.get('id', 'None')
                carousel_classes = carousel_details.get('classArray', [])
                carousel_classes_str = ', '.join(carousel_classes) if carousel_classes else 'None'
                
                doc.add_paragraph(f"Common ID pattern: {carousel_id}", style='List Bullet')
                doc.add_paragraph(f"Common class patterns: {carousel_classes_str}", style='List Bullet')
                
                # Extract slide count if available
                slide_count = carousel_example.get('slideCount', 'Unknown')
                doc.add_paragraph(f"Typical slides per carousel: {slide_count}", style='List Bullet')
            
            # Analyze card grids
            if 'cardGrids' in sample_content_blocks and sample_content_blocks['cardGrids']:
                card_example = sample_content_blocks['cardGrids'][0]
                doc.add_paragraph("Card Grid Structure:", style='List Bullet')
                
                card_count = card_example.get('cardCount', 'Unknown')
                card_consistency = card_example.get('cardPatternConsistency', 0) * 100
                
                doc.add_paragraph(f"Typical cards per grid: {card_count}", style='List Bullet')
                doc.add_paragraph(f"Card consistency: {card_consistency:.1f}%", style='List Bullet')
                
                # Check if we have a sample card to analyze
                sample_card = card_example.get('sampleCard', {})
                if sample_card:
                    doc.add_paragraph("Typical card contents:", style='List Bullet')
                    
                    has_image = element_contains_tag(sample_card, 'img')
                    has_heading = any(element_contains_tag(sample_card, f'h{i}') for i in range(1, 7))
                    has_link = element_contains_tag(sample_card, 'a')
                    
                    if has_image:
                        doc.add_paragraph("• Contains images", style='List Bullet')
                    if has_heading:
                        doc.add_paragraph("• Contains headings", style='List Bullet')
                    if has_link:
                        doc.add_paragraph("• Contains links", style='List Bullet')
        
        # Accessibility Implications Section
        doc.add_paragraph()
        h3 = doc.add_heading('Accessibility Implications', level=3)
        h3.style = doc.styles['Heading 3']
        
        doc.add_paragraph("""
        The structure of web pages has significant accessibility implications:
        
        1. Semantic Structure: Pages should use semantic HTML elements (<header>, <nav>, <main>, <footer>, etc.) to convey structure to assistive technologies.
        
        2. ARIA Roles: When semantic HTML isn't used, appropriate ARIA roles should be applied (role="banner", role="navigation", role="main", role="contentinfo").
        
        3. Skip Links: Sites should provide a mechanism to bypass repeated blocks of content (usually a "skip to main content" link).
        
        4. Consistent Navigation: Navigation mechanisms should be consistent across the site to aid users' mental model.
        
        5. Keyboard Navigation: All interface components should be navigable via keyboard, with a logical tab order.
        
        6. Visible Focus: Interactive elements should have a visible focus indicator.
        
        7. Responsive Structure: Page structure should adapt appropriately for different viewport sizes while maintaining accessibility.
        """.strip())
        
        # Recommendations
        doc.add_paragraph()
        h3 = doc.add_heading('Recommendations', level=3)
        h3.style = doc.styles['Heading 3']
        
        doc.add_paragraph("""
        Based on the structure analysis, consider these recommendations:
        
        1. Use semantic HTML elements properly for all major page components.
        
        2. Ensure all navigation elements are keyboard accessible.
        
        3. Provide a mechanism to bypass repeated content blocks.
        
        4. Maintain consistent structure across pages to aid user orientation.
        
        5. For dynamic content like carousels, ensure they are operable via keyboard and have appropriate ARIA attributes.
        
        6. For card layouts, ensure adequate color contrast and meaningful link text.
        
        7. Implement responsive design that maintains accessibility at all viewport sizes.
        """.strip())
    else:
        doc.add_paragraph("""
    No structure analysis data was found. Please ensure the page structure analysis test is properly integrated and the analysis has been run after testing.
        """.strip())

def get_top_detection_method(methods_dict):
    """Get the most common detection method from a dictionary of methods -> counts"""
    if not methods_dict:
        return "None"
    
    sorted_methods = sorted(methods_dict.items(), key=lambda x: x[1], reverse=True)
    if sorted_methods:
        return sorted_methods[0][0]
    return "Unknown"

def create_detection_methods_table(doc, methods_dict, total_pages):
    """Create a table showing detection methods and their frequency"""
    if not methods_dict:
        doc.add_paragraph("No detection methods data available.")
        return
    
    # Sort methods by frequency
    sorted_methods = sorted(methods_dict.items(), key=lambda x: x[1], reverse=True)
    
    # Create table
    table = doc.add_table(rows=len(sorted_methods) + 1, cols=2)
    table.style = 'Table Grid'
    
    # Headers
    headers = table.rows[0].cells
    headers[0].text = "Detection Method"
    headers[1].text = "Pages"
    
    # Add data
    for i, (method, count) in enumerate(sorted_methods, 1):
        row = table.rows[i].cells
        row[0].text = method
        percentage = (count / total_pages) * 100 if total_pages > 0 else 0
        row[1].text = f"{count} ({percentage:.1f}%)"
    
    format_table_text(table)
        