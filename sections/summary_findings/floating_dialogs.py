from report_styling import format_table_text

def add_floating_dialogs_section(doc, db_connection, total_domains):
    """Add the Floating Dialogs section to the summary findings"""
    h3 = doc.add_heading('Floating Dialogs', level=2)
    h3.style = doc.styles['Heading 2']

    # Query for pages with dialog issues - using the consolidated results field
    pages_with_dialog_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.floating_dialogs.dialogs.consolidated": {"$exists": True},
            "results.accessibility.tests.floating_dialogs.dialogs.consolidated.summary.totalIssues": {"$gt": 0}
        },
        {
            "url": 1,
            "results.accessibility.tests.floating_dialogs.dialogs.consolidated": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for each issue type by severity
    dialog_issues = {
        "violations": {
            "hiddenInteractiveContent": {"name": "Hidden interactive content", "pages": set(), "domains": set(), "severity": "critical"},
            "incorrectHeadingLevel": {"name": "Incorrect heading structure", "pages": set(), "domains": set(), "severity": "high"},
            "missingCloseButton": {"name": "Missing close button", "pages": set(), "domains": set(), "severity": "high"},
            "improperFocusManagement": {"name": "Improper focus management", "pages": set(), "domains": set(), "severity": "high"}
        },
        "warnings": {
            "contentOverlap": {"name": "Content overlap issues", "pages": set(), "domains": set(), "severity": "moderate"}
        }
    }

    # Count issues and store URLs by domain
    domain_to_urls = {}

    for page in pages_with_dialog_issues:
        url = page['url']
        domain = url.replace('http://', '').replace('https://', '').split('/')[0]
        consolidated = page['results']['accessibility']['tests']['floating_dialogs']['dialogs']['consolidated']
        
        # Initialize domain entry if it doesn't exist
        if domain not in domain_to_urls:
            domain_to_urls[domain] = {}
        
        # Process violations
        if 'issuesByType' in consolidated:
            issues_by_type = consolidated['issuesByType']
            
            # Process violations
            for violation_type, violation_data in issues_by_type.get('violations', {}).items():
                if violation_type in dialog_issues['violations'] and violation_data.get('count', 0) > 0:
                    dialog_issues['violations'][violation_type]['pages'].add(url)
                    dialog_issues['violations'][violation_type]['domains'].add(domain)
                    
                    # Store the severity if available
                    if 'severity' in violation_data:
                        dialog_issues['violations'][violation_type]['severity'] = violation_data['severity']
                    
                    # Store URL by issue type for this domain
                    if violation_type not in domain_to_urls[domain]:
                        domain_to_urls[domain][violation_type] = []
                    domain_to_urls[domain][violation_type].append(url)
            
            # Process warnings
            for warning_type, warning_data in issues_by_type.get('warnings', {}).items():
                if warning_type in dialog_issues['warnings'] and warning_data.get('count', 0) > 0:
                    dialog_issues['warnings'][warning_type]['pages'].add(url)
                    dialog_issues['warnings'][warning_type]['domains'].add(domain)
                    
                    # Store the severity if available
                    if 'severity' in warning_data:
                        dialog_issues['warnings'][warning_type]['severity'] = warning_data['severity']
                    
                    # Store URL by issue type for this domain
                    if warning_type not in domain_to_urls[domain]:
                        domain_to_urls[domain][warning_type] = []
                    domain_to_urls[domain][warning_type].append(url)

    # Create filtered list of issues that have affected pages
    all_active_issues = []

    for category in ['violations', 'warnings']:
        for issue_type, data in dialog_issues[category].items():
            if len(data['pages']) > 0:
                all_active_issues.append({
                    'category': category,
                    'type': issue_type,
                    'name': data['name'],
                    'severity': data['severity'],
                    'pages': data['pages'],
                    'domains': data['domains']
                })

    # Sort issues by severity - critical first, then high, then moderate
    severity_order = {'critical': 0, 'high': 1, 'moderate': 2, 'low': 3}
    all_active_issues.sort(key=lambda x: severity_order.get(x['severity'], 4))

    if all_active_issues:
        # Create summary table
        summary_table = doc.add_table(rows=len(all_active_issues) + 1, cols=5)
        summary_table.style = 'Table Grid'

        # Set column headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Issue"
        headers[1].text = "Severity"
        headers[2].text = "Pages Affected"
        headers[3].text = "Sites Affected"
        headers[4].text = "% of Total Sites"

        # Add data
        for i, issue in enumerate(all_active_issues, 1):
            row = summary_table.rows[i].cells
            row[0].text = issue['name']
            row[1].text = issue['severity'].capitalize()
            row[2].text = str(len(issue['pages']))
            row[3].text = str(len(issue['domains']))
            row[4].text = f"{(len(issue['domains']) / len(total_domains) * 100):.1f}%"

        # Format the table text
        format_table_text(summary_table)

    else:
        doc.add_paragraph("No floating dialog accessibility issues were found.")
        