import json
from report_styling import format_table_text
from section_aware_reporting import get_unique_section_issues

def add_headings_section(doc, db_connection, total_domains):
    """Add the Headings section to the summary findings"""
    h2 = doc.add_heading('Headings', level=2)
    h2.style = doc.styles['Heading 2']
    
    # Get section-aware issue statistics
    issue_data = get_unique_section_issues(db_connection, 'headings', issue_identifier='type')
    
    # Add explanation paragraph
    doc.add_paragraph("""
    Headings provide structure to web content and are essential for screen reader users to navigate and understand the organization of the page. 
    Proper heading structure follows a hierarchical pattern, starting with a single H1 that represents the page title, followed by H2 sections, 
    and then H3 subsections within those.
    """.strip())
    
    # Add total statistics paragraph
    doc.add_paragraph()
    total_text = ""
    if issue_data['has_section_data']:
        total_text = f"Found {issue_data['total_issues']} heading structure issues, representing {issue_data['unique_issues']} unique issues when accounting for repeating page sections."
    else:
        total_text = f"Found {issue_data['total_issues']} heading structure issues across {len(issue_data['domain_unique_issues'])} domains."
    doc.add_paragraph(total_text)
    
    # Process issues by type for the table
    issue_statistics = {}
    
    if issue_data['has_section_data']:
        # Extract issue statistics from section data
        for section_data in issue_data['section_statistics'].values():
            for issue_type, issue_data in section_data['issues'].items():
                if issue_type not in issue_statistics:
                    issue_statistics[issue_type] = {
                        'count': 0,
                        'domains': set(),
                        'pages': set()
                    }
                
                issue_statistics[issue_type]['count'] += issue_data['count']
                issue_statistics[issue_type]['domains'].update(issue_data['domains'])
                issue_statistics[issue_type]['pages'].update(issue_data['pages'])
    else:
        # Use the issue statistics directly
        for issue_type, issue_data in issue_data['issue_statistics'].items():
            issue_statistics[issue_type] = {
                'count': issue_data['count'],
                'domains': set(issue_data['domains']),
                'pages': set(issue_data['pages'])
            }
    
    # Create issue type descriptions for better readability
    issue_descriptions = {
        'missing-h1': 'Missing H1 heading',
        'multiple-h1s': 'Multiple H1 headings',
        'empty-heading': 'Empty heading',
        'invalid-level-before-main': 'Invalid heading level before main content',
        'contentinfo-wrong-heading-level': 'Wrong heading level in footer',
        'hierarchy-gap': 'Gap in heading hierarchy (e.g., H2 to H4)',
        'visual-hierarchy-issue': 'Visual hierarchy doesn\'t match semantic hierarchy'
    }
    
    # Create results table for issue types
    table = doc.add_table(rows=len(issue_statistics) + 1, cols=4)
    table.style = 'Table Grid'

    # Set column headers
    headers = table.rows[0].cells
    headers[0].text = "Issue"
    headers[1].text = "# of instances"
    headers[2].text = "# of sites"
    headers[3].text = "% of sites"

    # Add data for each issue type
    for i, (issue_type, stats) in enumerate(sorted(issue_statistics.items()), 1):
        row = table.rows[i].cells
        percentage = (len(stats['domains']) / len(total_domains)) * 100 if total_domains else 0
        
        # Use friendly description if available, otherwise use the raw issue type
        issue_desc = issue_descriptions.get(issue_type, issue_type)
        
        row[0].text = issue_desc
        row[1].text = str(stats['count'])
        row[2].text = str(len(stats['domains']))
        row[3].text = f"{percentage:.1f}%"

    # Format the table text
    format_table_text(table)

    # Add some space after the table
    doc.add_paragraph()
    
    # If we have section data, show issues by section
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