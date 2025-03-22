import json
from report_styling import format_table_text
from section_aware_reporting import get_unique_section_issues

def add_accessible_names_section(doc, db_connection, total_domains):
    """Add the Accessible Names section to the summary findings"""
    h2 = doc.add_heading('Accessible names', level=2)
    h2.style = doc.styles['Heading 2']
    
    # Get section-aware issue statistics
    issue_data = get_unique_section_issues(db_connection, 'accessible_names', issue_identifier='element')
    
    # Add explanation paragraph
    doc.add_paragraph("""
    Interactive elements such as links, buttons, form fields etc. must have an accessible name that can be programmatically determined. 
    This name is what will be announced by screen readers and other assistive technologies when the user encounters the element.
    Without an accessible name, users will not know the purpose or function of the element.
    """.strip())
    
    # Add total statistics paragraph
    doc.add_paragraph()
    total_text = f"Found {issue_data['total_issues']} instances of missing accessible names across {len(issue_data['domain_unique_issues'])} domains."
    if issue_data['has_section_data']:
        total_text += f" Based on our analysis, there are {issue_data['unique_issues']} unique issues when accounting for repeating page sections."
    doc.add_paragraph(total_text)
    
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
        
        # Add issues by tag (element) paragraph 
        doc.add_heading('Issues by Element Type', level=3)
    
    # If we don't have section data, just show by element type
    elif not issue_data['has_section_data']:
        doc.add_heading('Issues by Element Type', level=3)
    
    # Process and display issues by element type (tag)
    tag_statistics = {}
    
    if issue_data['has_section_data']:
        # Extract tag statistics from section data
        for section_data in issue_data['section_statistics'].values():
            for tag, tag_data in section_data['issues'].items():
                if tag not in tag_statistics:
                    tag_statistics[tag] = {
                        'count': 0,
                        'domains': set(),
                        'pages': set()
                    }
                
                tag_statistics[tag]['count'] += tag_data['count']
                tag_statistics[tag]['domains'].update(tag_data['domains'])
                tag_statistics[tag]['pages'].update(tag_data['pages'])
    else:
        # Use the issue statistics directly
        for tag, tag_data in issue_data['issue_statistics'].items():
            tag_statistics[tag] = {
                'count': tag_data['count'],
                'domains': set(tag_data['domains']),
                'pages': set(tag_data['pages'])
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