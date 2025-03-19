import json
from report_styling import format_table_text

def add_accessible_names_section(doc, db_connection, total_domains):
    """Add the Accessible Names section to the summary findings"""
    h2 = doc.add_heading('Accessible names', level=2)
    h2.style = doc.styles['Heading 2']
    
    # Original query to get pages with missing accessible names (for total counts)
    pages_with_name_issues = list(db_connection.page_results.find(
        {"results.accessibility.tests.accessible_names.accessible_names.details.summary.missingNames": {"$gt": 0}},
        {
            "url": 1,
            "results.accessibility.tests.accessible_names.accessible_names.details.summary.missingNames": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Count affected domains (needed for overall statistics)
    affected_domains = set()
    total_missing_names = 0
    for page in pages_with_name_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        affected_domains.add(domain)
        total_missing_names += page['results']['accessibility']['tests']['accessible_names']['accessible_names']['details']['summary']['missingNames']

    # Query for pages with violations for tag-specific analysis
    pages_with_violations = list(db_connection.page_results.find(
        {"results.accessibility.tests.accessible_names.accessible_names.details.violations": {"$exists": True}},
        {
            "url": 1,
            "results.accessibility.tests.accessible_names.accessible_names.details.violations": 1,
            "_id": 0
        }
    ))

    # Process violations to count by tag
    tag_statistics = {}
    
    for page in pages_with_violations:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        
        # Get the violations array and parse it if it's a string
        violations = page['results']['accessibility']['tests']['accessible_names']['accessible_names']['details']['violations']
        if isinstance(violations, str):
            violations = json.loads(violations)
        
        # Track unique tags for this page
        page_tags = set()
        
        for violation in violations:
            tag = violation['element']
            
            if tag not in tag_statistics:
                tag_statistics[tag] = {
                    'count': 0,
                    'pages': set(),
                    'domains': set()
                }
            
            tag_statistics[tag]['count'] += 1
            tag_statistics[tag]['pages'].add(page['url'])
            tag_statistics[tag]['domains'].add(domain)

    # Create results table
    table = doc.add_table(rows=len(tag_statistics) + 1, cols=4)
    table.style = 'Table Grid'

    # Set column headers
    headers = table.rows[0].cells
    headers[0].text = "Tag name"
    headers[1].text = "# of instances"
    headers[2].text = "# of sites"
    headers[3].text = "% of sites"

    # Add data for each tag
    for i, (tag, stats) in enumerate(sorted(tag_statistics.items()), 1):
        row = table.rows[i].cells
        percentage = (len(stats['domains']) / len(total_domains)) * 100 if total_domains else 0
        
        row[0].text = f"<{tag}>"
        row[1].text = str(stats['count'])
        row[2].text = str(len(stats['domains']))
        row[3].text = f"{percentage:.1f}%"

    # Format the table text
    format_table_text(table)

    # Add some space after the table
    doc.add_paragraph()