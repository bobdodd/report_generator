"""
Summary findings section for headings issues
"""
import os
from ...section_aware_reporting import process_section_statistics, format_section_table

def generate_headings_summary(db, domain):
    """
    Generate the headings summary section.
    
    Args:
        db: MongoDB database connection
        domain: Domain to generate report for
        
    Returns:
        String containing the HTML content for the section
    """
    # Analysis of headings section
    domain_pattern = domain.replace('.', '\\.')
    domain_filter = {'url': {'$regex': domain_pattern}}
    
    # Collect heading issues across all pages in the domain
    all_issues = []
    
    # Get all page results for the domain
    page_results = list(db.page_results.find(domain_filter))
    
    for page in page_results:
        url = page.get('url', '')
        
        # Navigate to the headings test results
        if 'results' in page and 'accessibility' in page['results']:
            accessibility = page['results']['accessibility']
            
            if 'tests' in accessibility and 'headings' in accessibility['tests']:
                test_data = accessibility['tests']['headings']
                
                # Check for details and violations
                if 'details' in test_data and 'violations' in test_data['details']:
                    violations = test_data['details']['violations']
                    
                    # Add page URL to each violation for tracking
                    for violation in violations:
                        violation_copy = violation.copy()
                        violation_copy['page_url'] = url
                        all_issues.append(violation_copy)
    
    # Count issues by type
    issue_counts = {}
    for issue in all_issues:
        issue_type = issue.get('element', 'unknown')
        if issue_type not in issue_counts:
            issue_counts[issue_type] = 0
        issue_counts[issue_type] += 1
    
    # Generate section statistics if section data is available
    section_table_html = ""
    section_stats = {}
    
    # First try to get section statistics directly from the stored data
    for page in page_results:
        # Navigate to the headings test results
        if 'results' in page and 'accessibility' in page['results']:
            accessibility = page['results']['accessibility']
            
            if 'tests' in accessibility and 'headings' in accessibility['tests']:
                test_data = accessibility['tests']['headings']
                
                # Check for section statistics
                if 'details' in test_data and 'section_statistics' in test_data['details']:
                    stored_stats = test_data['details']['section_statistics']
                    for section_type, count in stored_stats.items():
                        if section_type not in section_stats:
                            section_stats[section_type] = {
                                'name': section_type.capitalize(),
                                'count': 0,
                                'percentage': 0
                            }
                        section_stats[section_type]['count'] += count
    
    # Calculate percentages
    total_count = sum(s['count'] for s in section_stats.values())
    if total_count > 0:
        for section in section_stats.values():
            section['percentage'] = round((section['count'] / total_count) * 100, 1)
    
    # If we found section statistics, generate the table
    if section_stats:
        section_table_html = format_section_table(section_stats, "Heading Issues")
    # Otherwise, try to generate them from the issues
    else:
        section_aware_issues = [issue for issue in all_issues if 'section' in issue]
        if section_aware_issues:
            computed_stats = process_section_statistics(section_aware_issues)
            section_table_html = format_section_table(computed_stats, "Heading Issues")
    
    # Create the summary HTML
    html = f"""
    <div class="summary-section">
        <h2>Headings Structure Issues Summary</h2>
        
        <p>
            Headings provide structure to web content and are essential for screen reader users to navigate and understand 
            the organization of the page. Proper heading structure follows a hierarchical pattern, starting with a single 
            H1 that represents the page title, followed by H2 sections, and then H3 subsections within those.
        </p>
        
        <p>
            <strong>Total issues found: {len(all_issues)}</strong> across {len(page_results)} pages on {domain}.
        </p>
    """
    
    if issue_counts:
        html += """
        <h3>Issues by Element Type</h3>
        <ul>
        """
        
        # Add individual issue types
        for element_type, count in sorted(issue_counts.items(), key=lambda x: x[1], reverse=True):
            html += f"<li><strong>{element_type}</strong>: {count} issues</li>\n"
        
        html += "</ul>"
    
    html += """
        <h3>Impact on Users</h3>
        <p>
            When headings are not properly structured:
        </p>
        <ul>
            <li>Screen reader users cannot effectively navigate the page</li>
            <li>Users cannot understand the page organization</li>
            <li>Users with cognitive disabilities may have difficulty parsing content</li>
        </ul>
        
        <h3>Key WCAG Success Criteria</h3>
        <ul>
            <li><strong>1.3.1 Info and Relationships</strong> (Level A): Information, structure, and relationships can be programmatically determined</li>
            <li><strong>2.4.1 Bypass Blocks</strong> (Level A): Proper headings facilitate bypassing blocks of content</li>
            <li><strong>2.4.6 Headings and Labels</strong> (Level AA): Headings describe topic or purpose</li>
            <li><strong>2.4.10 Section Headings</strong> (Level AAA): Section headings are used to organize content</li>
        </ul>
    """
    
    # Add section-aware table if available
    html += section_table_html
    
    html += """
    </div>
    """
    
    return html