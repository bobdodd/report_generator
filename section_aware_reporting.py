"""
Section-aware report generation utilities.
"""
from pymongo import MongoClient

def get_unique_section_issues(db_connection, issue_type, domain, issue_identifier=None):
    """
    Get unique issue instances from the database, organized by page section.
    
    Args:
        db_connection: MongoDB connection
        issue_type: Type of issue (e.g., 'accessible_names', 'headings')
        domain: Domain to filter by
        issue_identifier: Optional identifier to further filter issues
        
    Returns:
        A dictionary of section-organized issues
    """
    try:
        # Find all page results for the domain
        domain_filter = {'url': {'$regex': domain}}
        page_results = list(db_connection.page_results.find(domain_filter))
        
        # Collect all issues with section information
        sections = {}
        
        for page in page_results:
            url = page.get('url', '')
            
            # Navigate to the test results
            if 'results' in page and 'accessibility' in page['results']:
                accessibility = page['results']['accessibility']
            elif 'accessibility' in page:
                # Fallback for older data structure
                accessibility = page['accessibility']
                
                if 'tests' in accessibility and issue_type in accessibility['tests']:
                    test_data = accessibility['tests'][issue_type]
                    
                    # Check for details and violations
                    if 'details' in test_data and 'violations' in test_data['details']:
                        violations = test_data['details']['violations']
                        
                        # Filter by issue_identifier if provided
                        if issue_identifier:
                            violations = [v for v in violations if v.get('issue') == issue_identifier]
                        
                        # Organize violations by section
                        for violation in violations:
                            if 'section' in violation:
                                section_type = violation['section'].get('section_type', 'unknown')
                                section_name = violation['section'].get('section_name', 'Unknown Section')
                                
                                # Create section entry if it doesn't exist
                                if section_type not in sections:
                                    sections[section_type] = {
                                        'name': section_name,
                                        'issues': []
                                    }
                                
                                # Add issue with page URL
                                issue_copy = violation.copy()
                                issue_copy['page_url'] = url
                                sections[section_type]['issues'].append(issue_copy)
        
        return sections
    
    except Exception as e:
        print(f"Error retrieving section issues: {e}")
        return {}

def process_section_statistics(violations):
    """
    Process violations to generate section statistics.
    
    Args:
        violations: List of violations with section information
        
    Returns:
        Dictionary of section statistics
    """
    section_stats = {}
    
    for violation in violations:
        if 'section' in violation:
            section_type = violation['section'].get('section_type', 'unknown')
            
            if section_type not in section_stats:
                section_stats[section_type] = {
                    'name': violation['section'].get('section_name', 'Unknown Section'),
                    'count': 0,
                    'elements': []
                }
            
            section_stats[section_type]['count'] += 1
            if 'element' in violation:
                section_stats[section_type]['elements'].append(violation['element'])
    
    # Calculate percentages
    total_violations = sum(s['count'] for s in section_stats.values())
    if total_violations > 0:
        for section in section_stats.values():
            section['percentage'] = round((section['count'] / total_violations) * 100, 1)
    
    return section_stats

def format_section_table(section_data, title):
    """
    Format section data into an HTML table for reporting.
    
    Args:
        section_data: Dictionary of section statistics
        title: Title for the table
        
    Returns:
        HTML table as a string
    """
    if not section_data:
        return f"<p>No section data available for {title}</p>"
    
    html = f"""
    <div class="section-table">
        <h3>{title} by Page Section</h3>
        <table border="1" cellpadding="4" cellspacing="0">
            <thead>
                <tr>
                    <th>Page Section</th>
                    <th>Count</th>
                    <th>Percentage</th>
                </tr>
            </thead>
            <tbody>
    """
    
    for section_type, section in section_data.items():
        html += f"""
            <tr>
                <td>{section['name']}</td>
                <td>{section['count']}</td>
                <td>{section.get('percentage', 0)}%</td>
            </tr>
        """
    
    html += """
            </tbody>
        </table>
    </div>
    """
    
    return html