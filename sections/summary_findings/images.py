from report_styling import format_table_text

def add_images_section(doc, db_connection, total_domains):
    """Add the Images section to the summary findings"""
    h2 = doc.add_heading('Images', level=2)
    h2.style = doc.styles['Heading 2']
    doc.add_paragraph()

    # Query for pages with image issues
    pages_with_image_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.images.images.pageFlags": {"$exists": True},
            "$or": [
                {"results.accessibility.tests.images.images.pageFlags.hasImagesWithoutAlt": True},
                {"results.accessibility.tests.images.images.pageFlags.hasImagesWithInvalidAlt": True},
                {"results.accessibility.tests.images.images.pageFlags.hasSVGWithoutRole": True}
            ]
        },
        {
            "url": 1,
            "results.accessibility.tests.images.images": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for different image issues
    image_issues = {
        "missing_alt": {
            "name": "Missing alt text",
            "pages": set(),
            "domains": set(),
            "count": 0
        },
        "invalid_alt": {
            "name": "Invalid alt text",
            "pages": set(),
            "domains": set(),
            "count": 0
        },
        "missing_role": {
            "name": "SVGs missing role",
            "pages": set(),
            "domains": set(),
            "count": 0
        }
    }

    # Process each page
    total_images = 0
    total_decorative = 0

    for page in pages_with_image_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        image_data = page['results']['accessibility']['tests']['images']['images']
        flags = image_data['pageFlags']
        details = flags['details']
        
        # Count total and decorative images
        total_images += details.get('totalImages', 0)
        total_decorative += details.get('decorativeImages', 0)
        
        # Check missing alt text
        if flags.get('hasImagesWithoutAlt'):
            image_issues['missing_alt']['pages'].add(page['url'])
            image_issues['missing_alt']['domains'].add(domain)
            image_issues['missing_alt']['count'] += details.get('missingAlt', 0)
        
        # Check invalid alt text
        if flags.get('hasImagesWithInvalidAlt'):
            image_issues['invalid_alt']['pages'].add(page['url'])
            image_issues['invalid_alt']['domains'].add(domain)
            image_issues['invalid_alt']['count'] += details.get('invalidAlt', 0)
        
        # Check missing SVG roles
        if flags.get('hasSVGWithoutRole'):
            image_issues['missing_role']['pages'].add(page['url'])
            image_issues['missing_role']['domains'].add(domain)
            image_issues['missing_role']['count'] += details.get('missingRole', 0)

    # Create filtered list of issues that have affected pages
    active_issues = {flag: data for flag, data in image_issues.items() 
                    if len(data['pages']) > 0}

    if active_issues:
        # Create summary table
        summary_table = doc.add_table(rows=len(active_issues) + 1, cols=5)
        summary_table.style = 'Table Grid'

        # Set column headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Issue Type"
        headers[1].text = "Number of Images"
        headers[2].text = "Pages Affected"
        headers[3].text = "Sites Affected"
        headers[4].text = "% of Total Sites"

        # Add data
        for i, (flag, data) in enumerate(active_issues.items(), 1):
            row = summary_table.rows[i].cells
            row[0].text = data['name']
            row[1].text = str(data['count'])
            row[2].text = str(len(data['pages']))
            row[3].text = str(len(data['domains']))
            row[4].text = f"{(len(data['domains']) / len(total_domains) * 100):.1f}%"

        # Format the table text
        format_table_text(summary_table)

    else:
        doc.add_paragraph("No image accessibility issues were found.")
        