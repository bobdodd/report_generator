from report_styling import format_table_text

def add_maps_section(doc, db_connection, total_domains):
    """Add the Maps section to the summary findings"""
    h2 = doc.add_heading('Maps', level=2)
    h2.style = doc.styles['Heading 2']

    # Query for pages with map issues
    pages_with_map_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.maps.maps.pageFlags": {"$exists": True},
            "$or": [
                {"results.accessibility.tests.maps.maps.pageFlags.hasMaps": True},
                {"results.accessibility.tests.maps.maps.pageFlags.hasMapsWithoutTitle": True},
                {"results.accessibility.tests.maps.maps.pageFlags.hasMapsWithAriaHidden": True}
            ]
        },
        {
            "url": 1,
            "results.accessibility.tests.maps.maps.pageFlags": 1,
            "results.accessibility.tests.maps.maps.details": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for each issue type
    map_issues = {
        "hasMaps": {"name": "Pages containing maps", "pages": set(), "domains": set()},
        "hasMapsWithoutTitle": {"name": "Maps without proper titles", "pages": set(), "domains": set()},
        "hasMapsWithAriaHidden": {"name": "Maps hidden from screen readers", "pages": set(), "domains": set()}
    }

    # Count issues
    for page in pages_with_map_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        flags = page['results']['accessibility']['tests']['maps']['maps']['pageFlags']
        
        for flag in map_issues:
            if flags.get(flag, False):
                map_issues[flag]['pages'].add(page['url'])
                map_issues[flag]['domains'].add(domain)

    # Create filtered list of issues that have affected pages
    active_issues = {flag: data for flag, data in map_issues.items() 
                    if len(data['pages']) > 0}

    if active_issues:
        # Create summary table
        summary_table = doc.add_table(rows=len(active_issues) + 1, cols=4)
        summary_table.style = 'Table Grid'

        # Set column headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Issue"
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
        doc.add_paragraph("No interactive maps found.")
        