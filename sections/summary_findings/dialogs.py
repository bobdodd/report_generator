from report_styling import format_table_text

def add_dialogs_section(doc, db_connection, total_domains):
    """Add the Dialogs section to the summary findings"""
    h2 = doc.add_heading('Dialogs', level=2)
    h2.style = doc.styles['Heading 2']
    
    # Query for pages with modal issues
    pages_with_modal_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.modals.modals.pageFlags.hasModals": True,
            "results.accessibility.tests.modals.modals.pageFlags.hasModalViolations": True
        },
        {
            "url": 1,
            "results.accessibility.tests.modals.modals.pageFlags": 1,
            "results.accessibility.tests.modals.modals.details.summary": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for each issue type
    modal_issues = {
        "modalsWithoutClose": {"name": "Missing close mechanism", "pages": set(), "domains": set()},
        "modalsWithoutFocusManagement": {"name": "Improper focus management", "pages": set(), "domains": set()},
        "modalsWithoutProperHeading": {"name": "Missing/improper heading", "pages": set(), "domains": set()},
        "modalsWithoutTriggers": {"name": "Missing/improper triggers", "pages": set(), "domains": set()}
    }

    # Count issues
    for page in pages_with_modal_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        summary = page['results']['accessibility']['tests']['modals']['modals']['details']['summary']
        
        for flag in modal_issues:
            if summary.get(flag, 0) > 0:
                modal_issues[flag]['pages'].add(page['url'])
                modal_issues[flag]['domains'].add(domain)

    # Create filtered list of issues that have affected pages
    active_issues = {flag: data for flag, data in modal_issues.items() 
                    if len(data['pages']) > 0}

    if active_issues:
        # Create summary table
        summary_table = doc.add_table(rows=len(active_issues) + 1, cols=4)
        summary_table.style = 'Table Grid'

        # Set column headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Modal Issue"
        headers[1].text = "Pages Affected"
        headers[2].text = "Sites Affected"
        headers[3].text = "% of Total Sites"

        # Add data
        for i, (flag, data) in enumerate(active_issues.items(), 1):
            row = summary_table.rows[i].cells
            row[0].text = data['name']
            row[1].text = str(len(data['pages']))
            row[2].text = str(len(data['domains']))
            row[3].text = f"{(len(data['domains']) / len(total_domains) * 100):.1f}%"

        # Format the table text
        format_table_text(summary_table)

    else:
        doc.add_paragraph("No dialog accessibility issues were found.")
        