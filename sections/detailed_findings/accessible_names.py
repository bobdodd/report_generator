import json
from report_styling import format_table_text
from section_aware_reporting import get_unique_section_issues, extract_domain_from_url

def add_detailed_accessible_names(doc, db_connection, total_domains):
    """Add the detailed Accessible Names section"""
    h2 = doc.add_heading('Accessible Names', level=2)
    h2.style = doc.styles['Heading 2']

    # Add explanation first (in case there's no data)
    doc.add_paragraph("""
    Interactive elements such as links, buttons, form fields etc. must have an accessible name that can be programmatically determined. This name is what will be announced by screen readers and other assistive technologies when the user encounters the element. Without an accessible name, users will not know the purpose or function of the element.
    """.strip())
    
    doc.add_paragraph()
    
    try:
        # Get section-aware issue statistics
        section_data = get_unique_section_issues(db_connection, 'accessible_names', issue_identifier='element')
        
        # If no issues found, show a placeholder message
        if not section_data:
            doc.add_paragraph("No accessibility issues with accessible names were found in this test run.")
            
            # Add section on WCAG requirements anyway
            doc.add_heading('WCAG Requirements', level=3)
            doc.add_paragraph("""
            The Web Content Accessibility Guidelines (WCAG) require that all interactive elements have names that can be programmatically determined:
            
            • WCAG 2.1 Success Criterion 1.1.1 Non-text Content (Level A): All non-text content that is presented to the user has a text alternative that serves the equivalent purpose.
            
            • WCAG 2.1 Success Criterion 4.1.2 Name, Role, Value (Level A): For all user interface components, the name and role can be programmatically determined.
            """.strip())
            
            return
    except Exception as e:
        # Handle any errors gracefully
        doc.add_paragraph(f"Note: Could not retrieve section-aware data for accessible names. Section data may not be available in this test run.")
        
        # Add section on WCAG requirements anyway
        doc.add_heading('WCAG Requirements', level=3)
        doc.add_paragraph("""
        The Web Content Accessibility Guidelines (WCAG) require that all interactive elements have names that can be programmatically determined:
        
        • WCAG 2.1 Success Criterion 1.1.1 Non-text Content (Level A): All non-text content that is presented to the user has a text alternative that serves the equivalent purpose.
        
        • WCAG 2.1 Success Criterion 4.1.2 Name, Role, Value (Level A): For all user interface components, the name and role can be programmatically determined.
        """.strip())
        
        return
    
    # Calculate total violations and unique URLs
    total_violations = 0
    unique_urls = set()
    domains = set()
    
    for section_type, violations in section_data.items():
        total_violations += len(violations)
        for violation in violations:
            if 'page_url' in violation:
                unique_urls.add(violation['page_url'])
                # Extract domain from URL
                domain = extract_domain_from_url(violation['page_url'])
                domains.add(domain)
    
    # Add total statistics paragraph
    total_text = f"Found {total_violations} instances of missing accessible names across {len(domains)} domains."
    total_text += f" Issues appeared on {len(unique_urls)} unique URLs."
    doc.add_paragraph(total_text)
    
    # Add section on WCAG requirements
    doc.add_heading('WCAG Requirements', level=3)
    doc.add_paragraph("""
    The Web Content Accessibility Guidelines (WCAG) require that all interactive elements have names that can be programmatically determined:
    
    • WCAG 2.1 Success Criterion 1.1.1 Non-text Content (Level A): All non-text content that is presented to the user has a text alternative that serves the equivalent purpose.
    
    • WCAG 2.1 Success Criterion 4.1.2 Name, Role, Value (Level A): For all user interface components, the name and role can be programmatically determined.
    """.strip())
    
    doc.add_paragraph()
    
    # Show issues by section
    doc.add_heading('Issues by Page Section', level=3)
    
    # Map section types to friendly names
    section_names = {
        'header': 'Header',
        'footer': 'Footer',
        'navigation': 'Navigation Menu',
        'mainContent': 'Main Content Area',
        'complementaryContent': 'Sidebar Content',
        'search': 'Search Component',
        'cookie': 'Cookie Notice',
        'heroSection': 'Hero Section',
        'form': 'Form Section',
        'topArea': 'Top of Page',
        'middleArea': 'Middle of Page',
        'bottomArea': 'Bottom of Page',
        'unknown': 'Unknown Section'
    }
    
    # Process each section
    for section_type, violations in section_data.items():
        if not violations:
            continue
            
        section_name = section_names.get(section_type, section_type)
        doc.add_heading(f"Issues in {section_name}", level=4)
        
        # Add a list of example issues in this section
        for violation in violations[:5]:  # Show up to 5 examples
            p = doc.add_paragraph(style='List Bullet')
            
            # Try to create a descriptive message
            if 'message' in violation:
                p.add_run(violation['message'])
            elif 'element' in violation and isinstance(violation['element'], dict) and 'html' in violation['element']:
                html = violation['element']['html']
                p.add_run(f"Element missing accessible name: {html[:100]}")
            else:
                p.add_run("Missing accessible name on element")
            
            # Add URL if available
            if 'page_url' in violation:
                p.add_run(f" (found on {violation['page_url']})")
        
        doc.add_paragraph()
    
    # Add recommendations
    doc.add_heading('Recommendations', level=3)
    doc.add_paragraph("""
    To ensure all interactive elements have accessible names:
    
    1. Ensure all images have appropriate alt text
    2. Add labels to all form controls
    3. Make sure buttons and links have descriptive text
    4. Use aria-label or aria-labelledby when visible text is not sufficient
    5. Test with screen readers to verify that names are announced correctly
    """.strip())