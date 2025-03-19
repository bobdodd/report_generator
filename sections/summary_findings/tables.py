from report_styling import format_table_text

def add_tables_section(doc, db_connection, total_domains):
    """Add the Tables section to the summary findings"""
    h2 = doc.add_heading('Tables', level=2)
    h2.style = doc.styles['Heading 2']
 
    # Query for pages with table issues
    pages_with_table_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.tables.tables.pageFlags": {"$exists": True},
            "$or": [
                {"results.accessibility.tests.tables.tables.pageFlags.hasMissingHeaders": True},
                {"results.accessibility.tests.tables.tables.pageFlags.hasNoScope": True},
                {"results.accessibility.tests.tables.tables.pageFlags.hasMissingCaption": True},
                {"results.accessibility.tests.tables.tables.pageFlags.hasLayoutTables": True},
                {"results.accessibility.tests.tables.tables.pageFlags.hasComplexTables": True}
            ]
        },
        {
            "url": 1,
            "results.accessibility.tests.tables.tables.pageFlags": 1,
            "results.accessibility.tests.tables.tables.details": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for each issue type
    table_issues = {
        "hasMissingHeaders": {"name": "Missing table headers", "pages": set(), "domains": set()},
        "hasNoScope": {"name": "Missing scope attributes", "pages": set(), "domains": set()},
        "hasMissingCaption": {"name": "Missing table captions", "pages": set(), "domains": set()},
        "hasLayoutTables": {"name": "Layout tables", "pages": set(), "domains": set()},
        "hasComplexTables": {"name": "Complex tables without proper structure", "pages": set(), "domains": set()}
    }

    # Count issues
    for page in pages_with_table_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        flags = page['results']['accessibility']['tests']['tables']['tables']['pageFlags']
        
        for flag in table_issues:
            if flags.get(flag, False):
                table_issues[flag]['pages'].add(page['url'])
                table_issues[flag]['domains'].add(domain)

    # Create filtered list of issues that have affected pages
    active_issues = {flag: data for flag, data in table_issues.items() 
                    if len(data['pages']) > 0}

    if active_issues:
        # Create summary table
        summary_table = doc.add_table(rows=len(active_issues) + 1, cols=4)
        summary_table.style = 'Table Grid'

        # Set column headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Table Issue"
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
        doc.add_paragraph("No table issues found.")