import json
from report_styling import format_table_text
from section_aware_reporting import get_unique_section_issues

def add_detailed_accessible_names(doc, db_connection, total_domains):
    """Add the detailed Accessible Names section"""
    h2 = doc.add_heading('Accessible Names', level=2)
    h2.style = doc.styles['Heading 2']

    # Get section-aware issue statistics
    issue_data = get_unique_section_issues(db_connection, 'accessible_names', issue_identifier='element')
    
    # Add explanation
    doc.add_paragraph("""
    Interactive elements such as links, buttons, form fields etc. must have an accessible name that can be programmatically determined. This name is what will be announced by screen readers and other assistive technologies when the user encounters the element. Without an accessible name, users will not know the purpose or function of the element.
    """.strip())
    
    doc.add_paragraph()
    
    # Add total statistics paragraph
    total_text = f"Found {issue_data['total_issues']} instances of missing accessible names across {len(issue_data['domain_unique_issues'])} domains."
    if issue_data['has_section_data']:
        total_text += f" Based on our analysis, there are {issue_data['unique_issues']} unique issues when accounting for repeating page sections."
    doc.add_paragraph(total_text)
    
    # Add section on WCAG requirements
    doc.add_heading('WCAG Requirements', level=3)
    doc.add_paragraph("""
    The Web Content Accessibility Guidelines (WCAG) require that all interactive elements have names that can be programmatically determined:
    
    • WCAG 2.1 Success Criterion 1.1.1 Non-text Content (Level A): All non-text content that is presented to the user has a text alternative that serves the equivalent purpose.
    
    • WCAG 2.1 Success Criterion 4.1.2 Name, Role, Value (Level A): For all user interface components, the name and role can be programmatically determined.
    """.strip())
    
    doc.add_paragraph()
    
    # If we have section data, show issues by section with more details
    if issue_data['has_section_data']:
        doc.add_heading('Issues by Page Section', level=3)
        
        # Get relevant sections with issues
        sections_with_issues = {
            section_type: section_data 
            for section_type, section_data in issue_data['section_statistics'].items()
            if section_data['total_count'] > 0
        }
        
        # Create table for section statistics
        section_table = doc.add_table(rows=len(sections_with_issues) + 1, cols=4)
        section_table.style = 'Table Grid'
        
        # Set column headers
        headers = section_table.rows[0].cells
        headers[0].text = "Page Section"
        headers[1].text = "# of issues"
        headers[2].text = "# of sites"
        headers[3].text = "% of sites"
        
        # Add data for each section
        for i, (section_type, section_data) in enumerate(sorted(sections_with_issues.items()), 1):
            row = section_table.rows[i].cells
            section_name = section_data['name']
            issue_count = section_data['total_count']
            domain_count = len(section_data['domains'])
            percentage = (domain_count / len(total_domains)) * 100 if total_domains else 0
            
            row[0].text = section_name
            row[1].text = str(issue_count)
            row[2].text = str(domain_count)
            row[3].text = f"{percentage:.1f}%"
        
        # Format the table text
        format_table_text(section_table)
        
        # Add some space after the table
        doc.add_paragraph()
        
        # Add detailed breakdown for each section
        doc.add_heading('Detailed Breakdown by Section', level=3)
        
        # Process each section
        for section_type, section_data in sorted(sections_with_issues.items()):
            section_name = section_data['name']
            
            # Add section header
            doc.add_paragraph(f"Section: {section_name}", style='Heading 4')
            
            # Add brief description of issues in this section
            section_desc = f"Found {section_data['total_count']} instances of accessibility issues in {section_name} sections "
            section_desc += f"across {len(section_data['domains'])} domains."
            doc.add_paragraph(section_desc)
            
            # Get all issues for this section
            issues = section_data['issues']
            
            # Create table for element breakdown in this section
            if issues:
                element_table = doc.add_table(rows=len(issues) + 1, cols=3)
                element_table.style = 'Table Grid'
                
                # Set headers
                headers = element_table.rows[0].cells
                headers[0].text = "Element"
                headers[1].text = "# of instances"
                headers[2].text = "# of sites affected"
                
                # Add issue data
                for i, (element, element_data) in enumerate(sorted(issues.items()), 1):
                    row = element_table.rows[i].cells
                    row[0].text = f"<{element}>"
                    row[1].text = str(element_data['count'])
                    row[2].text = str(len(element_data['domains']))
                
                # Format the table text
                format_table_text(element_table)
            
            # Add space after each section's breakdown
            doc.add_paragraph()
    
    # If we don't have section data, show by element type
    else:
        # Create element type table similar to the summary section
        doc.add_heading('Issues by Element Type', level=3)
        
        tag_statistics = {}
        for tag, tag_data in issue_data['issue_statistics'].items():
            tag_statistics[tag] = {
                'count': tag_data['count'],
                'domains': tag_data['domains'],
                'pages': tag_data['pages']
            }
        
        # Create results table
        table = doc.add_table(rows=len(tag_statistics) + 1, cols=4)
        table.style = 'Table Grid'

        # Set column headers
        headers = table.rows[0].cells
        headers[0].text = "Tag name"
        headers[1].text = "# of instances"
        headers[2].text = "# of sites"
        headers[3].text = "% of sites"

        # Add data for each tag
        for i, (tag, stats) in enumerate(sorted(tag_statistics.items()), 1):
            row = table.rows[i].cells
            percentage = (len(stats['domains']) / len(total_domains)) * 100 if total_domains else 0
            
            row[0].text = f"<{tag}>"
            row[1].text = str(stats['count'])
            row[2].text = str(len(stats['domains']))
            row[3].text = f"{percentage:.1f}%"

        # Format the table text
        format_table_text(table)
        
        # Add some space after the table
        doc.add_paragraph()
        
        # Add detailed breakdown by domain (original style)
        doc.add_heading('Detailed Breakdown by Domain', level=3)
        
        # Create domain breakdown
        for domain, domain_data in sorted(issue_data['issues_by_domain'].items()):
            if not domain_data:
                continue
                
            # Add domain header
            doc.add_paragraph(f"Domain: {domain}", style='Heading 4')
            
            # Create table for this domain's element breakdown
            domain_table = doc.add_table(rows=len(domain_data) + 1, cols=2)
            domain_table.style = 'Table Grid'
            
            # Set headers
            headers = domain_table.rows[0].cells
            headers[0].text = "Element"
            headers[1].text = "# of instances"
            
            # Add element data
            for i, (element, element_data) in enumerate(sorted(domain_data.items()), 1):
                row = domain_table.rows[i].cells
                row[0].text = f"<{element}>"
                row[1].text = str(element_data['count'])
            
            # Format the table text
            format_table_text(domain_table)
            
            # Add space after each domain's breakdown
            doc.add_paragraph()
    
    # Add remediation guidance section
    doc.add_heading('How to Fix', level=3)
    doc.add_paragraph("""
    To fix missing accessible names, follow these guidelines for different elements:
    
    • Images: Add meaningful alt text that describes the purpose or content of the image. Use empty alt text (alt="") for decorative images.
    
    • Buttons and links: Ensure they have clear text content. For icon-only buttons or links, add aria-label attributes.
    
    • Form controls: Use properly associated <label> elements, or aria-label/aria-labelledby attributes.
    
    • Iframes: Always include a title attribute that describes the iframe's purpose or content.
    
    • ARIA components: Use aria-label or aria-labelledby to provide an accessible name.
    """.strip())
    
    # Add impact section
    doc.add_heading('Impact on Users', level=3)
    doc.add_paragraph("""
    Missing accessible names make it impossible for screen reader users to understand the purpose of interactive elements. This results in:
    
    • Screen readers announcing generic text like "button" or "image" without any context
    • Users unable to understand what will happen when they activate a control
    • Form fields that can't be understood or completed properly
    
    This creates significant barriers for people who rely on screen readers or other assistive technologies.
    """.strip())