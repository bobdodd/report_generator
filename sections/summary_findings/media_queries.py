from docx.oxml import parse_xml
from report_styling import format_table_text

def add_media_queries_section(doc, db_connection, total_domains):
    """Add the Media Queries section to the summary findings"""
    doc.add_paragraph()
    h2 = doc.add_heading('Media Queries and Responsive Design', level=2)
    h2.style = doc.styles['Heading 2']

    # Query for pages with responsive design issues
    pages_with_media_query_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.media_queries.media_queries.pageFlags": {"$exists": True},
            "$or": [
                {"results.accessibility.tests.media_queries.media_queries.pageFlags.hasResponsiveBreakpoints": False},
                {"results.accessibility.tests.media_queries.media_queries.pageFlags.hasPrintStyles": False},
                {"results.accessibility.tests.media_queries.media_queries.pageFlags.hasReducedMotionSupport": False},
                {"results.accessibility.tests.media_queries.media_queries.pageFlags.hasDarkModeSupport": False}
            ]
        },
        {
            "url": 1,
            "results.accessibility.tests.media_queries.media_queries.pageFlags": 1,
            "results.accessibility.tests.media_queries.media_queries.details.summary": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Count affected domains for each issue
    issue_counts = {
        "no_responsive": {"name": "No responsive breakpoints", "pages": set(), "domains": set()},
        "no_print": {"name": "No print stylesheets", "pages": set(), "domains": set()},
        "no_reduced_motion": {"name": "No reduced motion support", "pages": set(), "domains": set()},
        "no_dark_mode": {"name": "No dark mode support", "pages": set(), "domains": set()}
    }

    for page in pages_with_media_query_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        flags = page['results']['accessibility']['tests']['media_queries']['media_queries']['pageFlags']
        
        if not flags.get('hasResponsiveBreakpoints', True):
            issue_counts["no_responsive"]["pages"].add(page['url'])
            issue_counts["no_responsive"]["domains"].add(domain)
            
        if not flags.get('hasPrintStyles', True):
            issue_counts["no_print"]["pages"].add(page['url'])
            issue_counts["no_print"]["domains"].add(domain)
            
        if not flags.get('hasReducedMotionSupport', True):
            issue_counts["no_reduced_motion"]["pages"].add(page['url'])
            issue_counts["no_reduced_motion"]["domains"].add(domain)
            
        if not flags.get('hasDarkModeSupport', True):
            issue_counts["no_dark_mode"]["pages"].add(page['url'])
            issue_counts["no_dark_mode"]["domains"].add(domain)

    # Create summary table
    table = doc.add_table(rows=5, cols=4)
    table.style = 'Table Grid'

    # Set column headers
    headers = table.rows[0].cells
    headers[0].text = "Issue"
    headers[1].text = "# of pages"
    headers[2].text = "# of sites affected"
    headers[3].text = "% of sites"

    # Add data for each issue
    for i, (issue_key, issue_data) in enumerate(issue_counts.items(), 1):
        row = table.rows[i].cells
        row[0].text = issue_data["name"]
        row[1].text = str(len(issue_data["pages"]))
        row[2].text = str(len(issue_data["domains"]))
        percentage = (len(issue_data["domains"]) / len(total_domains)) * 100 if total_domains else 0
        row[3].text = f"{percentage:.1f}%"

    # Format the table text
    format_table_text(table)

    # Add some explanation
    doc.add_paragraph()
    p = doc.add_paragraph("Media queries are essential for implementing responsive design and honoring user preferences. Their proper implementation affects several WCAG criteria:")
    
    doc.add_paragraph("1.4.4 Resize text - Content can be resized up to 200% without loss of functionality", style='List Bullet')
    doc.add_paragraph("1.4.10 Reflow - Content can be presented without horizontal scrolling at widths up to 320px", style='List Bullet')
    doc.add_paragraph("2.3.3 Animation from Interactions - Motion animation triggered by interaction can be disabled", style='List Bullet')
    doc.add_paragraph("1.4.8 Visual Presentation - Users can select foreground and background colors (e.g., dark mode)", style='List Bullet')
    
    # Add space after the section
    doc.add_paragraph()