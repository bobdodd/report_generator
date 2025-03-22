from report_styling import format_table_text

def add_site_specific_reports(doc, db_connection, total_domains):
    """Add site-specific report sections to the document"""

    # Get all test runs to find relevant page results
    all_test_runs = db_connection.get_all_test_runs()
    test_run_ids = [str(run['_id']) for run in all_test_runs]
    
    # Get all unique URLs from the database
    all_urls = db_connection.page_results.distinct('url', {'test_run_id': {'$in': test_run_ids}})
    
    # Group URLs by domain
    domain_urls = {}
    for url in all_urls:
        domain = url.replace('http://', '').replace('https://', '').split('/')[0]
        if domain not in domain_urls:
            domain_urls[domain] = []
        domain_urls[domain].append(url)
    
    # For each domain, create a specific report section
    for domain in sorted(domain_urls.keys()):
        add_domain_specific_section(doc, db_connection, domain, domain_urls[domain])
        
        # Add a page break between domains (except after the last one)
        if domain != sorted(domain_urls.keys())[-1]:
            doc.add_page_break()

def add_domain_specific_section(doc, db_connection, domain, urls):
    """Add a domain-specific section to the report"""
    # Add domain heading
    h1 = doc.add_heading(f"{domain} Site Specific Report", level=1)
    h1.style = doc.styles['Heading 1']
    
    # Introduction
    doc.add_paragraph(f"""
    This section provides a detailed analysis specific to the {domain} website. The findings 
    presented here are based on the accessibility tests conducted on {len(urls)} pages from this domain.
    """.strip())
    
    doc.add_paragraph()
    
    # Add overview of pages tested
    h2 = doc.add_heading('Pages Tested', level=2)
    h2.style = doc.styles['Heading 2']
    
    # Create a table for all tested pages
    pages_table = doc.add_table(rows=len(urls) + 1, cols=2)
    pages_table.style = 'Table Grid'
    
    # Add headers
    headers = pages_table.rows[0].cells
    headers[0].text = "URL"
    headers[1].text = "Page Title"
    
    # Get all test runs
    all_test_runs = db_connection.get_all_test_runs()
    test_run_ids = [str(run['_id']) for run in all_test_runs]
    
    # Add page data with titles where available
    for i, url in enumerate(sorted(urls), 1):
        # Find the page result to get the title - use more comprehensive query projection
        page_result = db_connection.page_results.find_one(
            {'url': url, 'test_run_id': {'$in': test_run_ids}},
            {
                'page_title': 1, 
                'accessibility.tests.html_structure.details.title.analysis.text': 1,
                'accessibility.title': 1,  # Some might have this structure
                'results.accessibility.tests.html_structure.details.title.analysis.text': 1,
                'results.accessibility.title': 1,
                'results.tests.html_structure.details.title.analysis.text': 1,
                'url': 1
            }
        )
        
        row = pages_table.rows[i].cells
        row[0].text = url
        
        # Extract the base URL for a more user-friendly presentation
        base_url = url.replace('https://', '').replace('http://', '')
        
        # Extract a title from the URL if we can't find a proper title
        url_title = url.split('/')[-1].replace('.aspx', '').replace('.html', '').replace('-', ' ').replace('_', ' ').title()
        if url_title.lower() == 'default' or url_title.lower() == 'index':
            # Try to get a better title from the parent directory
            parts = url.split('/')
            if len(parts) > 4:  # At least has a parent directory
                url_title = parts[-2].replace('-', ' ').replace('_', ' ').title()
        
        # Get page title if available, with multiple fallback options
        try:
            # Try all possible paths to find a title
            title = None
            
            # Check all possible paths where title might be stored
            paths = [
                lambda d: d.get('page_title'),
                lambda d: d.get('accessibility', {}).get('tests', {}).get('html_structure', {}).get('details', {}).get('title', {}).get('analysis', {}).get('text'),
                lambda d: d.get('accessibility', {}).get('title'),
                lambda d: d.get('results', {}).get('accessibility', {}).get('tests', {}).get('html_structure', {}).get('details', {}).get('title', {}).get('analysis', {}).get('text'),
                lambda d: d.get('results', {}).get('accessibility', {}).get('title'),
                lambda d: d.get('results', {}).get('tests', {}).get('html_structure', {}).get('details', {}).get('title', {}).get('analysis', {}).get('text')
            ]
            
            # Try each path until we find a title
            for path_func in paths:
                if page_result:
                    potential_title = path_func(page_result)
                    if potential_title and len(potential_title.strip()) > 0:
                        title = potential_title
                        break
            
            # If no title found, use the URL-derived title as fallback
            if not title or len(title.strip()) == 0:
                title = f"{url_title} (URL-derived)"
                
            row[1].text = title
        except (AttributeError, KeyError, TypeError) as e:
            # Fallback to URL-derived title
            row[1].text = f"{url_title} (URL-derived)"
    
    format_table_text(pages_table)
    
    # Query for issues specific to this domain
    issue_categories = [
        ('Accessible Names', 'accessible_names'),
        ('Color Contrast', 'color_contrast'),
        ('Forms', 'forms'),
        ('Headings', 'headings'),
        ('Images', 'images'),
        ('Landmarks', 'landmarks'),
        ('Language', 'language')
    ]
    
    # Loop through each issue category
    for display_name, category_key in issue_categories:
        # Query the database for issues in this category for this domain
        has_issues = False
        
        # Check if any pages for this domain have issues in this category
        for url in urls:
            query = {
                'url': url,
                'test_run_id': {'$in': test_run_ids},
                f'results.accessibility.tests.{category_key}.has_issues': True
            }
            
            if db_connection.page_results.find_one(query):
                has_issues = True
                break
        
        # Add a subheading for this category
        h3 = doc.add_heading(display_name, level=3)
        h3.style = doc.styles['Heading 3']
        
        # Add content based on whether issues were found
        if has_issues:
            doc.add_paragraph(f"Issues related to {display_name.lower()} were found on this domain.")
            
            # Count pages with issues
            issue_count = 0
            for url in urls:
                query = {
                    'url': url,
                    'test_run_id': {'$in': test_run_ids},
                    f'results.accessibility.tests.{category_key}.has_issues': True
                }
                
                if db_connection.page_results.find_one(query):
                    issue_count += 1
            
            doc.add_paragraph(f"Issues were found on {issue_count} of {len(urls)} pages ({(issue_count/len(urls)*100):.1f}%).")
        else:
            doc.add_paragraph(f"No issues related to {display_name.lower()} were found on this domain.")
    
    # Add summary and recommendations
    h2 = doc.add_heading('Summary and Recommendations', level=2)
    h2.style = doc.styles['Heading 2']
    
    doc.add_paragraph("""
    Based on the accessibility tests conducted, this section summarizes the findings specific 
    to this domain and provides recommendations for improvement.
    """.strip())
    
    # Count total issues by severity
    doc.add_paragraph()
    
    # Placeholder for recommendations - in a real implementation, 
    # this would be dynamically generated based on the actual findings
    doc.add_paragraph(f"Recommendations for {domain}:", style='Normal')
    doc.add_paragraph("Review and implement WCAG 2.1 Level AA compliance across the site", style='List Bullet')
    doc.add_paragraph("Ensure all interactive elements are keyboard accessible", style='List Bullet')
    doc.add_paragraph("Provide appropriate alternative text for all images", style='List Bullet')
    doc.add_paragraph("Ensure proper heading structure throughout all pages", style='List Bullet')
    doc.add_paragraph("Maintain sufficient color contrast for text elements", style='List Bullet')