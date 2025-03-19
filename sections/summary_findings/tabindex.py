# sections/summary_findings/tabindex.py
from report_styling import format_table_text

def add_tabindex_section(doc, db_connection, total_domains):
    """Add the summary Tabindex section"""
    doc.add_paragraph()
    h3 = doc.add_heading('Tabindex', level=2)
    h3.style = doc.styles['Heading 2']

    # Query for pages with tabindex issues
    pages_with_tabindex_issues = list(db_connection.page_results.find(
        {"results.accessibility.tests.tabindex.tabindex.pageFlags": {"$exists": True}},
        {
            "url": 1,
            "results.accessibility.tests.tabindex.tabindex.pageFlags": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for each issue type
    tabindex_issues = {
        "hasPositiveTabindex": {"name": "Elements with positive tabindex", "pages": set(), "domains": set()},
        "hasNonInteractiveZeroTabindex": {"name": "Non-interactive elements with tabindex=0", "pages": set(), "domains": set()},
        "hasMissingRequiredTabindex": {"name": "Interactive elements missing required tabindex", "pages": set(), "domains": set()},
        "hasSvgTabindexWarnings": {"name": "SVG elements with tabindex warnings", "pages": set(), "domains": set()}
    }    

    # Count issues
    for page in pages_with_tabindex_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        flags = page['results']['accessibility']['tests']['tabindex']['tabindex']['pageFlags']
        
        for flag in tabindex_issues:
            if flags.get(flag, False):  # If issue exists (True)
                tabindex_issues[flag]['pages'].add(page['url'])
                tabindex_issues[flag]['domains'].add(domain)

    # Create filtered list of issues that have affected pages
    active_issues = {flag: data for flag, data in tabindex_issues.items() 
                    if len(data['pages']) > 0}

    if active_issues:
        # Create summary table
        summary_table = doc.add_table(rows=len(active_issues) + 1, cols=4)
        summary_table.style = 'Table Grid'

        # Set column headers
        tiheaders = summary_table.rows[0].cells
        tiheaders[0].text = "Issue"
        tiheaders[1].text = "Pages Affected"
        tiheaders[2].text = "Sites Affected"
        tiheaders[3].text = "% of Total Sites"

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
        doc.add_paragraph("No tabindex accessibility issues were found.")