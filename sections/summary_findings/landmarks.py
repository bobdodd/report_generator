from report_styling import format_table_text

def add_landmarks_section(doc, db_connection, total_domains):
    """Add the Landmarks section to the summary findings"""
    h2 = doc.add_heading('Landmarks', level=2)
    h2.style = doc.styles['Heading 2']

    # Query for pages with landmark issues
    pages_with_landmark_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.landmarks.landmarks.pageFlags": {"$exists": True},
            "$or": [
                {"results.accessibility.tests.landmarks.landmarks.pageFlags.missingRequiredLandmarks": True},
                {"results.accessibility.tests.landmarks.landmarks.pageFlags.hasDuplicateLandmarksWithoutNames": True},
                {"results.accessibility.tests.landmarks.landmarks.pageFlags.hasNestedTopLevelLandmarks": True},
                {"results.accessibility.tests.landmarks.landmarks.pageFlags.hasContentOutsideLandmarks": True}
            ]
        },
        {
            "url": 1,
            "results.accessibility.tests.landmarks.landmarks": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for different landmark issues
    landmark_issues = {
        "missing": {
            "name": "Missing required landmarks",
            "pages": set(),
            "domains": set(),
            "details": {
                "banner": 0,
                "main": 0,
                "contentinfo": 0,
                "search": 0
            }
        },
        "duplicate": {
            "name": "Duplicate landmarks without unique names",
            "pages": set(),
            "domains": set(),
            "details": {
                "banner": 0,
                "main": 0,
                "navigation": 0,
                "complementary": 0,
                "contentinfo": 0,
                "search": 0,
                "form": 0,
                "region": 0
            }
        },
        "nested": {
            "name": "Nested top-level landmarks",
            "pages": set(),
            "domains": set()
        },
        "outside": {
            "name": "Content outside landmarks",
            "pages": set(),
            "domains": set(),
            "count": 0
        }
    }

    # Process each page
    total_landmarks = 0
    for page in pages_with_landmark_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        landmark_data = page['results']['accessibility']['tests']['landmarks']['landmarks']
        flags = landmark_data['pageFlags']
        details = flags['details']
        
        # Count total landmarks
        if 'totalLandmarks' in landmark_data.get('details', {}).get('summary', {}):
            total_landmarks += landmark_data['details']['summary']['totalLandmarks']
        
        # Check missing landmarks
        if flags.get('missingRequiredLandmarks'):
            landmark_issues['missing']['pages'].add(page['url'])
            landmark_issues['missing']['domains'].add(domain)
            missing = details.get('missingLandmarks', {})
            for landmark in ['banner', 'main', 'contentinfo', 'search']:
                if missing.get(landmark):
                    landmark_issues['missing']['details'][landmark] += 1

        # Check duplicate landmarks
        if flags.get('hasDuplicateLandmarksWithoutNames'):
            landmark_issues['duplicate']['pages'].add(page['url'])
            landmark_issues['duplicate']['domains'].add(domain)
            duplicates = details.get('duplicateLandmarks', {})
            for landmark in landmark_issues['duplicate']['details'].keys():
                if landmark in duplicates:
                    landmark_issues['duplicate']['details'][landmark] += duplicates[landmark].get('count', 0)

        # Check nested landmarks
        if flags.get('hasNestedTopLevelLandmarks'):
            landmark_issues['nested']['pages'].add(page['url'])
            landmark_issues['nested']['domains'].add(domain)

        # Check content outside landmarks
        if flags.get('hasContentOutsideLandmarks'):
            landmark_issues['outside']['pages'].add(page['url'])
            landmark_issues['outside']['domains'].add(domain)
            landmark_issues['outside']['count'] += details.get('contentOutsideLandmarksCount', 0)

    # Create summary table
    if any(len(issue['pages']) > 0 for issue in landmark_issues.values()):
        # Create main issues summary table
        summary_table = doc.add_table(rows=len(landmark_issues) + 1, cols=4)
        summary_table.style = 'Table Grid'

        # Set column headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Issue Type"
        headers[1].text = "Pages Affected"
        headers[2].text = "Sites Affected"
        headers[3].text = "% of Total Sites"

        # Add data
        row_idx = 1
        for issue_type, data in landmark_issues.items():
            if len(data['pages']) > 0:
                row = summary_table.rows[row_idx].cells
                row[0].text = data['name']
                row[1].text = str(len(data['pages']))
                row[2].text = str(len(data['domains']))
                row[3].text = f"{(len(data['domains']) / len(total_domains) * 100):.1f}%"
                row_idx += 1

        # Format the table text
        format_table_text(summary_table)

    else:
        doc.add_paragraph("No landmark structure issues were found.")
        