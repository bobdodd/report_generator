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

def add_detailed_structure(doc, db_connection, total_domains):
    """Add the detailed Page Structure section"""
    doc.add_page_break()
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
        
        # Component Presence
        doc.add_paragraph()
        doc.add_paragraph("Common UI Components Presence:", style='Normal')
        
        # Calculate component presence across all domains
        header_pages = 0
        footer_pages = 0
        nav_pages = 0
        main_pages = 0
        complementary_pages = 0
        search_pages = 0
        
        for domain_data in domain_analyses.values():
            header_pages += domain_data.get('header_analysis', {}).get('pages_with_component', 0)
            footer_pages += domain_data.get('footer_analysis', {}).get('pages_with_component', 0)
            nav_pages += domain_data.get('navigation_analysis', {}).get('pages_with_component', 0)
            main_pages += domain_data.get('main_content_analysis', {}).get('pages_with_component', 0)
            complementary_pages += domain_data.get('complementary_analysis', {}).get('pages_with_component', 0)
            
            # Search might be in component_presence or directly in the domain data
            search_count = domain_data.get('component_presence', {}).get('search', 0)
            if search_count == 0:  # Try alternate location
                search_count = domain_data.get('search_components', 0)
            search_pages += search_count
        
        component_table = doc.add_table(rows=7, cols=3)  # Expanded to include main and complementary content
        component_table.style = 'Table Grid'
        
        # Headers
        component_headers = component_table.rows[0].cells
        component_headers[0].text = "Component"
        component_headers[1].text = "Pages"
        component_headers[2].text = "% of Total"
        
        # Component data
        components = [
            ("Header", header_pages),
            ("Footer", footer_pages),
            ("Navigation", nav_pages),
            ("Main Content", main_pages),
            ("Complementary Content", complementary_pages),
            ("Search", search_pages)
        ]
        
        for i, (component, count) in enumerate(components, 1):
            row = component_table.rows[i].cells
            row[0].text = component
            row[1].text = str(count)
            percentage = (count / total_pages) * 100 if total_pages > 0 else 0
            row[2].text = f"{percentage:.1f}%"
        
        format_table_text(component_table)
        
        # Header Analysis
        doc.add_paragraph()
        h3 = doc.add_heading('Header Analysis', level=3)
        h3.style = doc.styles['Heading 3']
        
        # Find a domain with header data to use as an example
        example_domain = None
        for domain, data in domain_analyses.items():
            if data.get('header_analysis', {}).get('pages_with_component', 0) > 0:
                example_domain = domain
                break
        
        if example_domain:
            header_analysis = domain_analyses[example_domain]['header_analysis']
            
            # Common patterns
            patterns = header_analysis.get('common_patterns', {})
            
            doc.add_paragraph(f"Most common header tag: <{patterns.get('tag', 'unknown')}>")
            
            # Header structure details
            doc.add_paragraph("Header structure details:")
            
            # If we have sample data, use it to provide more information
            sample_pages = domain_analyses[example_domain].get('sample_pages', {})
            if sample_pages:
                sample_url = next(iter(sample_pages.keys()))
                sample_data = sample_pages[sample_url]
                
                header_element = sample_data.get('keyElements', {}).get('header', {})
                if header_element:
                    # Count total descendants (not just direct children)
                    descendants = count_descendants(header_element)
                    doc.add_paragraph(f"Average header complexity: {descendants} total elements", style='List Bullet')
                    
                    # Look for specific elements within header
                    links = count_element_type(header_element, 'a')
                    buttons = count_element_type(header_element, 'button')
                    images = count_element_type(header_element, 'img')
                    
                    doc.add_paragraph(f"Typical header contains: {links} links, {buttons} buttons, {images} images", style='List Bullet')
                    
                    # Determine if header likely contains site navigation
                    has_nav = element_contains_tag(header_element, 'nav')
                    doc.add_paragraph(f"Header contains navigation menu: {'Yes' if has_nav else 'No'}", style='List Bullet')
            
            if patterns.get('common_classes'):
                doc.add_paragraph("Common header CSS classes:")
                for cls in patterns.get('common_classes', []):
                    doc.add_paragraph(cls, style='List Bullet')
        else:
            doc.add_paragraph("No consistent header structure was identified across pages.")
        
        # Footer Analysis
        doc.add_paragraph()
        h3 = doc.add_heading('Footer Analysis', level=3)
        h3.style = doc.styles['Heading 3']
        
        # Find a domain with footer data to use as an example
        example_domain = None
        for domain, data in domain_analyses.items():
            if data.get('footer_analysis', {}).get('pages_with_component', 0) > 0:
                example_domain = domain
                break
        
        if example_domain:
            footer_analysis = domain_analyses[example_domain]['footer_analysis']
            
            # Common patterns
            patterns = footer_analysis.get('common_patterns', {})
            
            doc.add_paragraph(f"Most common footer tag: <{patterns.get('tag', 'unknown')}>")
            
            # Footer structure details
            doc.add_paragraph("Footer structure details:")
            
            # If we have sample data, use it to provide more information
            sample_pages = domain_analyses[example_domain].get('sample_pages', {})
            if sample_pages:
                sample_url = next(iter(sample_pages.keys()))
                sample_data = sample_pages[sample_url]
                
                footer_element = sample_data.get('keyElements', {}).get('footer', {})
                if footer_element:
                    # Count total descendants (not just direct children)
                    descendants = count_descendants(footer_element)
                    doc.add_paragraph(f"Average footer complexity: {descendants} total elements", style='List Bullet')
                    
                    # Look for specific elements within footer
                    links = count_element_type(footer_element, 'a')
                    buttons = count_element_type(footer_element, 'button')
                    images = count_element_type(footer_element, 'img')
                    
                    doc.add_paragraph(f"Typical footer contains: {links} links, {buttons} buttons, {images} images", style='List Bullet')
            
            if patterns.get('common_classes'):
                doc.add_paragraph("Common footer CSS classes:")
                for cls in patterns.get('common_classes', []):
                    doc.add_paragraph(cls, style='List Bullet')
        else:
            doc.add_paragraph("No consistent footer structure was identified across pages.")
        
        # Navigation Analysis
        doc.add_paragraph()
        h3 = doc.add_heading('Navigation Analysis', level=3)
        h3.style = doc.styles['Heading 3']
        
        # Find a domain with navigation data to use as an example
        example_domain = None
        for domain, data in domain_analyses.items():
            if data.get('navigation_analysis', {}).get('pages_with_component', 0) > 0:
                example_domain = domain
                break
        
        if example_domain:
            navigation_analysis = domain_analyses[example_domain]['navigation_analysis']
            
            # Common patterns
            patterns = navigation_analysis.get('common_patterns', {})
            
            doc.add_paragraph(f"Most common navigation tag: <{patterns.get('tag', 'unknown')}>")
            
            # Navigation structure details
            doc.add_paragraph("Navigation structure details:")
            
            # If we have sample data, use it to provide more information
            sample_pages = domain_analyses[example_domain].get('sample_pages', {})
            if sample_pages:
                sample_url = next(iter(sample_pages.keys()))
                sample_data = sample_pages[sample_url]
                
                nav_element = sample_data.get('keyElements', {}).get('navigation', {})
                if nav_element:
                    # Count total links
                    links = count_element_type(nav_element, 'a')
                    doc.add_paragraph(f"Average navigation contains: {links} links", style='List Bullet')
            
            if patterns.get('common_classes'):
                doc.add_paragraph("Common navigation CSS classes:")
                for cls in patterns.get('common_classes', []):
                    doc.add_paragraph(cls, style='List Bullet')
        else:
            doc.add_paragraph("No consistent navigation structure was identified across pages.")
    else:
        doc.add_paragraph("""
    No structure analysis data was found. Please ensure the page structure analysis test is properly integrated and the analysis has been run after testing.
        """.strip())
        