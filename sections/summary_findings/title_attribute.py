from report_styling import format_table_text

def add_title_attribute_section(doc, db_connection, total_domains):
    """Add the Title Attribute section to the summary findings"""
    h2 = doc.add_heading('Title Attribute', level=2)
    h2.style = doc.styles['Heading 2']

    # Query for pages with title attribute issues
    pages_with_title_issues = list(db_connection.page_results.find(
        {"results.accessibility.tests.title.titleAttribute.pageFlags.hasImproperTitleAttributes": True},
        {
            "url": 1,
            "results.accessibility.tests.title.titleAttribute.details": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Count affected domains
    affected_domains = set()
    total_improper_uses = 0
    domain_counts = {}

    for page in pages_with_title_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        affected_domains.add(domain)
        
        # Count improper uses from the details
        improper_uses = len(page['results']['accessibility']['tests']['title']['titleAttribute']['details']['improperUse'])
        total_improper_uses += improper_uses
        
        # Track counts by domain
        if domain not in domain_counts:
            domain_counts[domain] = 0
        domain_counts[domain] += improper_uses

    # Calculate percentage
    percentage = (len(affected_domains) / len(total_domains)) * 100 if total_domains else 0

    # Create summary table
    summary_table = doc.add_table(rows=2, cols=4)
    summary_table.style = 'Table Grid'

    # Set column headers
    headers = summary_table.rows[0].cells
    headers[0].text = "Issue"
    headers[1].text = "Total Occurrences"
    headers[2].text = "Sites Affected"
    headers[3].text = "% of Total Sites"

    # Add data
    row = summary_table.rows[1].cells
    row[0].text = "Improper use of title attribute"
    row[1].text = str(total_improper_uses)
    row[2].text = str(len(affected_domains))
    row[3].text = f"{percentage:.1f}%"

    # Format the table text
    format_table_text(summary_table)

    # Add some space after the table
    doc.add_paragraph()
    