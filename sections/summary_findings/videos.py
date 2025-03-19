from report_styling import format_table_text

def add_videos_section(doc, db_connection, total_domains):
    """Add the Videos section to the summary findings"""
    h2 = doc.add_heading('Videos', level=2)
    h2.style = doc.styles['Heading 2']

    # Query for pages with video issues
    pages_with_video_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.video.video.pageFlags": {"$exists": True},
            "$or": [
                {"results.accessibility.tests.video.video.pageFlags.missingCaptions": True},
                {"results.accessibility.tests.video.video.pageFlags.missingAudioDescription": True},
                {"results.accessibility.tests.video.video.pageFlags.inaccessibleControls": True},
                {"results.accessibility.tests.video.video.pageFlags.missingTranscript": True},
                {"results.accessibility.tests.video.video.pageFlags.hasAutoplay": True},
                {"results.accessibility.tests.video.video.pageFlags.missingLabels": True}
            ]
        },
        {
            "url": 1,
            "results.accessibility.tests.video.video.pageFlags": 1,
            "results.accessibility.tests.video.video.details": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for each issue type
    video_issues = {
        "missingCaptions": {"name": "Missing closed captions", "pages": set(), "domains": set()},
        "missingAudioDescription": {"name": "Missing audio descriptions", "pages": set(), "domains": set()},
        "inaccessibleControls": {"name": "Inaccessible video controls", "pages": set(), "domains": set()},
        "missingTranscript": {"name": "Missing transcripts", "pages": set(), "domains": set()},
        "hasAutoplay": {"name": "Autoplay without user control", "pages": set(), "domains": set()},
        "missingLabels": {"name": "Missing video labels/titles", "pages": set(), "domains": set()}
    }

    # Count issues
    if (len(pages_with_video_issues) > 0):
        for page in pages_with_video_issues:
            domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
            flags = page['results']['accessibility']['tests']['video']['video']['pageFlags']
            
            for flag in video_issues:
                if flags.get(flag, False):
                    video_issues[flag]['pages'].add(page['url'])
                    video_issues[flag]['domains'].add(domain)

        # Create filtered list of issues that have affected pages
        active_issues = {flag: data for flag, data in video_issues.items() 
                        if len(data['pages']) > 0}

        if active_issues:
            # Create summary table
            summary_table = doc.add_table(rows=len(active_issues) + 1, cols=4)
            summary_table.style = 'Table Grid'

            # Set column headers
            headers = summary_table.rows[0].cells
            headers[0].text = "Video Issue"
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
            doc.add_paragraph("No video accessibility issues were found.")
    else:
        doc.add_paragraph("No videos were found.")