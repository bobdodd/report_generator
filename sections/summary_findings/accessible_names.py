"""
Summary findings section for accessible names issues
"""
import os
from ...section_aware_reporting import process_section_statistics, format_section_table

def generate_accessible_names_summary(db, domain):
    """
    Generate the accessible names summary section.
    
    Args:
        db: MongoDB database connection
        domain: Domain to generate report for
        
    Returns:
        String containing the HTML content for the section
    """
    # Analysis of accessible names section
    domain_pattern = domain.replace('.', '\\.')
    domain_filter = {'url': {'$regex': domain_pattern}}
    
    # Collect accessible name issues across all pages in the domain
    all_issues = []
    
    # Get all page results for the domain
    page_results = list(db.page_results.find(domain_filter))
    
    for page in page_results:
        url = page.get('url', '')
        
        # Navigate to the accessible_names test results
        if 'results' in page and 'accessibility' in page['results']:
            accessibility = page['results']['accessibility']
            
            if 'tests' in accessibility and 'accessible_names' in accessibility['tests']:
                test_data = accessibility['tests']['accessible_names']
                
                # Check for details and violations
                if 'details' in test_data and 'violations' in test_data['details']:
                    violations = test_data['details']['violations']
                    
                    # Add page URL to each violation for tracking
                    for violation in violations:
                        violation_copy = violation.copy()
                        violation_copy['page_url'] = url
                        all_issues.append(violation_copy)
    
    # Print debug info for violations
    print(f"\nFound {len(all_issues)} issues for {domain}")
    for page in page_results:
        print(f"Checking page: {page.get('url', 'unknown')}")
        if 'results' in page:
            print(f"Results keys: {list(page['results'].keys())}")
            if 'accessibility' in page['results']:
                accessibility = page['results']['accessibility']
                print(f"Accessibility keys: {list(accessibility.keys())}")
                if 'tests' in accessibility and 'accessible_names' in accessibility['tests']:
                    test_data = accessibility['tests']['accessible_names']
                    print(f"Test data keys: {list(test_data.keys())}")
                    if 'details' in test_data:
                        details = test_data['details']
                        print(f"Details keys: {list(details.keys())}")
                        if 'section_statistics' in details:
                            print(f"Section statistics: {details['section_statistics']}")
                        if 'violations' in details:
                            violations = details['violations']
                            print(f"Found {len(violations)} violations")
                            # Print first violation
                            if violations:
                                print(f"First violation: {violations[0]}")
    
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
        # Navigate to the accessible_names test results
        if 'results' in page and 'accessibility' in page['results']:
            accessibility = page['results']['accessibility']
            
            if 'tests' in accessibility and 'accessible_names' in accessibility['tests']:
                test_data = accessibility['tests']['accessible_names']
                
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
        section_table_html = format_section_table(section_stats, "Accessible Name Issues")
    # Otherwise, try to generate them from the issues
    else:
        section_aware_issues = [issue for issue in all_issues if 'section' in issue]
        if section_aware_issues:
            computed_stats = process_section_statistics(section_aware_issues)
            section_table_html = format_section_table(computed_stats, "Accessible Name Issues")
    
    # Create the summary HTML
    html = f"""
    <div class="summary-section">
        <h2>Accessible Names Issues Summary</h2>
        
        <p>
            Accessible names are text alternatives that identify interactive elements for users of assistive technologies.
            These names are essential for screen reader users to understand the purpose of buttons, links, form controls, and images.
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
            When interactive elements lack proper accessible names:
        </p>
        <ul>
            <li>Screen reader users cannot determine the purpose of controls</li>
            <li>Voice control users cannot activate elements by speaking their names</li>
            <li>Users with cognitive disabilities may have difficulty understanding element functions</li>
        </ul>
        
        <h3>Key WCAG Success Criteria</h3>
        <ul>
            <li><strong>1.1.1 Non-text Content</strong> (Level A): All non-text content has a text alternative</li>
            <li><strong>2.4.4 Link Purpose</strong> (Level A): The purpose of links can be determined from the link text</li>
            <li><strong>3.3.2 Labels or Instructions</strong> (Level A): Form controls have associated labels</li>
            <li><strong>4.1.2 Name, Role, Value</strong> (Level A): UI components have accessible names</li>
        </ul>
    """
    
    # Add section-aware table if available
    html += section_table_html
    
    html += """
    </div>
    """
    
    return html