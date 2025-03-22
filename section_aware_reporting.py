"""
Section-aware reporting utilities for the report generator.
These functions help process and aggregate accessibility issues by page section.
"""
import json

def get_unique_section_issues(db_connection, test_name, issue_identifier="element", extra_fields=None):
    """
    Get issues grouped by section and count unique issues per section type.
    This handles repeating sections (header, footer, etc.) properly by counting
    each issue type only once per section type per domain.
    
    Args:
        db_connection: Database connection object
        test_name: Name of the test (e.g., 'accessible_names')
        issue_identifier: Field to use for issue identification (default: 'element')
        extra_fields: Additional fields to include in results (optional)
        
    Returns:
        Dictionary with section statistics and issues
    """
    # Query for pages with violations that include section information
    section_query_path = f"results.accessibility.tests.{test_name}.{test_name}.details.section_statistics"
    violations_path = f"results.accessibility.tests.{test_name}.{test_name}.details.violations"
    
    # Check if any pages have section statistics
    has_section_data = db_connection.page_results.find_one({section_query_path: {"$exists": True}})
    
    if not has_section_data:
        # Fallback to regular issue counting if no section data available
        print(f"No section data found for {test_name}, falling back to regular issue counting")
        return get_regular_issues(db_connection, test_name, issue_identifier, extra_fields)
    
    # Query for pages with violations
    projection = {
        "url": 1,
        violations_path: 1,
        "_id": 0
    }
    
    pages_with_violations = list(db_connection.page_results.find(
        {violations_path: {"$exists": True}},
        projection
    ))
    
    # Process violations by section
    section_statistics = {}
    sections_by_domain = {}
    
    for page in pages_with_violations:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        
        # Initialize domain data
        if domain not in sections_by_domain:
            sections_by_domain[domain] = {}
        
        # Get violations array and parse if needed
        violations_path_parts = violations_path.split('.')
        violations = page
        for part in violations_path_parts:
            if part in violations:
                violations = violations[part]
            else:
                violations = []
                break
                
        if isinstance(violations, str):
            try:
                violations = json.loads(violations)
            except:
                violations = []
        
        # Process each violation
        for violation in violations:
            if not isinstance(violation, dict):
                continue
                
            # Get the section info if available
            section_info = violation.get('section', {})
            section_type = section_info.get('section_type', 'unknown')
            section_name = section_info.get('section_name', 'Unknown Section')
            is_primary = section_info.get('primary', False)
            
            # Get the issue identifier
            issue_key = violation.get(issue_identifier, 'unknown')
            
            # Create a unique key for this issue+section type
            issue_section_key = f"{issue_key}|{section_type}"
            
            # Initialize section in statistics if needed
            if section_type not in section_statistics:
                section_statistics[section_type] = {
                    'name': section_name,
                    'total_count': 0,
                    'is_primary': is_primary,
                    'issues': {},
                    'domains': set(),
                    'pages': set()
                }
            
            # Update global section statistics
            section_statistics[section_type]['total_count'] += 1
            section_statistics[section_type]['domains'].add(domain)
            section_statistics[section_type]['pages'].add(page['url'])
            
            # Initialize issue in this section if needed
            if issue_key not in section_statistics[section_type]['issues']:
                section_statistics[section_type]['issues'][issue_key] = {
                    'count': 0,
                    'domains': set(),
                    'pages': set(),
                    'sample': violation  # Store a sample violation for reference
                }
            
            # Update issue statistics for this section
            section_statistics[section_type]['issues'][issue_key]['count'] += 1
            section_statistics[section_type]['issues'][issue_key]['domains'].add(domain)
            section_statistics[section_type]['issues'][issue_key]['pages'].add(page['url'])
            
            # Initialize section type in domain if needed
            if section_type not in sections_by_domain[domain]:
                sections_by_domain[domain][section_type] = {
                    'name': section_name,
                    'is_primary': is_primary,
                    'count': 0,
                    'issues': set(),
                    'pages': set()
                }
            
            # Update domain-specific section statistics
            sections_by_domain[domain][section_type]['count'] += 1
            sections_by_domain[domain][section_type]['issues'].add(issue_key)
            sections_by_domain[domain][section_type]['pages'].add(page['url'])
    
    # Convert set to list for serialization
    for section in section_statistics.values():
        section['domains'] = list(section['domains'])
        section['pages'] = list(section['pages'])
        
        for issue in section['issues'].values():
            issue['domains'] = list(issue['domains'])
            issue['pages'] = list(issue['pages'])
    
    for domain in sections_by_domain.values():
        for section in domain.values():
            section['issues'] = list(section['issues'])
            section['pages'] = list(section['pages'])
    
    # Calculate unique issues vs total instances
    total_issues = sum(section['total_count'] for section in section_statistics.values())
    unique_issues = sum(len(section['issues']) for section in section_statistics.values())
    
    # Calculate unique issues per domain
    domain_unique_issues = {}
    for domain, sections in sections_by_domain.items():
        domain_unique_issues[domain] = sum(len(section['issues']) for section in sections.values())
    
    return {
        'section_statistics': section_statistics,
        'sections_by_domain': sections_by_domain,
        'total_issues': total_issues,
        'unique_issues': unique_issues,
        'domain_unique_issues': domain_unique_issues,
        'has_section_data': True
    }

def get_regular_issues(db_connection, test_name, issue_identifier="element", extra_fields=None):
    """
    Fallback function for tests without section data
    """
    violations_path = f"results.accessibility.tests.{test_name}.{test_name}.details.violations"
    
    projection = {
        "url": 1,
        violations_path: 1,
        "_id": 0
    }
    
    pages_with_violations = list(db_connection.page_results.find(
        {violations_path: {"$exists": True}},
        projection
    ))
    
    # Process violations
    issue_statistics = {}
    issues_by_domain = {}
    
    for page in pages_with_violations:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        
        # Initialize domain data
        if domain not in issues_by_domain:
            issues_by_domain[domain] = {}
        
        # Get violations array and parse if needed
        violations_path_parts = violations_path.split('.')
        violations = page
        for part in violations_path_parts:
            if part in violations:
                violations = violations[part]
            else:
                violations = []
                break
                
        if isinstance(violations, str):
            try:
                violations = json.loads(violations)
            except:
                violations = []
        
        # Process each violation
        for violation in violations:
            if not isinstance(violation, dict):
                continue
                
            # Get the issue identifier
            issue_key = violation.get(issue_identifier, 'unknown')
            
            # Initialize issue in statistics if needed
            if issue_key not in issue_statistics:
                issue_statistics[issue_key] = {
                    'count': 0,
                    'domains': set(),
                    'pages': set(),
                    'sample': violation  # Store a sample violation for reference
                }
            
            # Update issue statistics
            issue_statistics[issue_key]['count'] += 1
            issue_statistics[issue_key]['domains'].add(domain)
            issue_statistics[issue_key]['pages'].add(page['url'])
            
            # Initialize issue in domain if needed
            if issue_key not in issues_by_domain[domain]:
                issues_by_domain[domain][issue_key] = {
                    'count': 0,
                    'pages': set()
                }
            
            # Update domain-specific issue statistics
            issues_by_domain[domain][issue_key]['count'] += 1
            issues_by_domain[domain][issue_key]['pages'].add(page['url'])
    
    # Convert set to list for serialization
    for issue in issue_statistics.values():
        issue['domains'] = list(issue['domains'])
        issue['pages'] = list(issue['pages'])
    
    for domain in issues_by_domain.values():
        for issue in domain.values():
            issue['pages'] = list(issue['pages'])
    
    # Calculate statistics
    total_issues = sum(issue['count'] for issue in issue_statistics.values())
    unique_issues = len(issue_statistics)
    
    # Calculate unique issues per domain
    domain_unique_issues = {domain: len(issues) for domain, issues in issues_by_domain.items()}
    
    return {
        'issue_statistics': issue_statistics,
        'issues_by_domain': issues_by_domain,
        'total_issues': total_issues,
        'unique_issues': unique_issues,
        'domain_unique_issues': domain_unique_issues,
        'has_section_data': False
    }