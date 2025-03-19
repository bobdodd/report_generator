from report_styling import format_table_text

def add_timers_section(doc, db_connection, total_domains):
    """Add the Timers section to the summary findings"""
    h2 = doc.add_heading('Timers', level=2)
    h2.style = doc.styles['Heading 2']

    # Query for pages with timer issues
    pages_with_timer_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.timers.timers.pageFlags": {"$exists": True},
            "$or": [
                {"results.accessibility.tests.timers.timers.pageFlags.hasTimers": True},
                {"results.accessibility.tests.timers.timers.pageFlags.hasAutoStartTimers": True},
                {"results.accessibility.tests.timers.timers.pageFlags.hasTimersWithoutControls": True}
            ]
        },
        {
            "url": 1,
            "results.accessibility.tests.timers.timers.pageFlags": 1,
            "results.accessibility.tests.timers.timers.details": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for each issue type
    timer_issues = {
        "hasTimers": {"name": "Pages with timers", "pages": set(), "domains": set()},
        "hasAutoStartTimers": {"name": "Auto-starting timers", "pages": set(), "domains": set()},
        "hasTimersWithoutControls": {"name": "Timers without adequate controls", "pages": set(), "domains": set()}
    }

    # Count issues
    for page in pages_with_timer_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        flags = page['results']['accessibility']['tests']['timers']['timers']['pageFlags']
        
        for flag in timer_issues:
            if flags.get(flag, False):
                timer_issues[flag]['pages'].add(page['url'])
                timer_issues[flag]['domains'].add(domain)

    # Create filtered list of issues that have affected pages
    active_issues = {flag: data for flag, data in timer_issues.items() 
                    if len(data['pages']) > 0}

    if active_issues:
        # Create summary table
        summary_table = doc.add_table(rows=len(active_issues) + 1, cols=4)
        summary_table.style = 'Table Grid'

        # Set column headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Timer Issue"
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
        doc.add_paragraph("No timer-related issues were found.")
        