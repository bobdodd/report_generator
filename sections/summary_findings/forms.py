from report_styling import format_table_text

def add_forms_section(doc, db_connection, total_domains):
    """Add the Forms section to the summary findings"""
    h2 = doc.add_heading('Forms', level=2)
    h2.style = doc.styles['Heading 2']

    doc.add_paragraph()

    # Query for pages with form issues
    pages_with_form_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.forms.forms.pageFlags": {"$exists": True},
            "$or": [
                {"results.accessibility.tests.forms.forms.pageFlags.hasInputsWithoutLabels": True},
                {"results.accessibility.tests.forms.forms.pageFlags.hasPlaceholderOnlyInputs": True},
                {"results.accessibility.tests.forms.forms.pageFlags.hasFormsWithoutHeadings": True},
                {"results.accessibility.tests.forms.forms.pageFlags.hasFormsOutsideLandmarks": True},
                {"results.accessibility.tests.forms.forms.pageFlags.hasContrastIssues": True},
                {"results.accessibility.tests.forms.forms.pageFlags.hasLayoutIssues": True}
            ]
        },
        {
            "url": 1,
            "results.accessibility.tests.forms.forms": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for different form issues
    form_issues = {
        "missing_labels": {
            "name": "Inputs without labels",
            "pages": set(),
            "domains": set(),
            "count": 0
        },
        "placeholder_only": {
            "name": "Placeholder-only inputs",
            "pages": set(),
            "domains": set(),
            "count": 0
        },
        "no_headings": {
            "name": "Forms without headings",
            "pages": set(),
            "domains": set(),
            "count": 0
        },
        "outside_landmarks": {
            "name": "Forms outside landmarks",
            "pages": set(),
            "domains": set(),
            "count": 0
        },
        "contrast_issues": {
            "name": "Input contrast issues",
            "pages": set(),
            "domains": set(),
            "count": 0
        },
        "layout_issues": {
            "name": "Form layout issues",
            "pages": set(),
            "domains": set(),
            "count": 0
        }
    }

    # Process each page
    total_forms = 0
    for page in pages_with_form_issues:
        try:
            domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
            form_data = page['results']['accessibility']['tests']['forms']['forms']
            flags = form_data.get('pageFlags', {})
            summary = form_data.get('details', {}).get('summary', {})
            
            # Update total forms count
            total_forms += summary.get('totalForms', 0)
            
            # Check inputs without labels
            if flags.get('hasInputsWithoutLabels'):
                form_issues['missing_labels']['pages'].add(page['url'])
                form_issues['missing_labels']['domains'].add(domain)
                form_issues['missing_labels']['count'] += summary.get('inputsWithoutLabels', 0)
            
            # Check placeholder-only inputs
            if flags.get('hasPlaceholderOnlyInputs'):
                form_issues['placeholder_only']['pages'].add(page['url'])
                form_issues['placeholder_only']['domains'].add(domain)
                form_issues['placeholder_only']['count'] += summary.get('inputsWithPlaceholderOnly', 0)
            
            # Check forms without headings
            if flags.get('hasFormsWithoutHeadings'):
                form_issues['no_headings']['pages'].add(page['url'])
                form_issues['no_headings']['domains'].add(domain)
                form_issues['no_headings']['count'] += summary.get('formsWithoutHeadings', 0)
            
            # Check forms outside landmarks
            if flags.get('hasFormsOutsideLandmarks'):
                form_issues['outside_landmarks']['pages'].add(page['url'])
                form_issues['outside_landmarks']['domains'].add(domain)
                form_issues['outside_landmarks']['count'] += summary.get('formsOutsideLandmarks', 0)
            
            # Check contrast issues
            if flags.get('hasContrastIssues'):
                form_issues['contrast_issues']['pages'].add(page['url'])
                form_issues['contrast_issues']['domains'].add(domain)
                form_issues['contrast_issues']['count'] += summary.get('inputsWithContrastIssues', 0)
            
            # Check layout issues
            if flags.get('hasLayoutIssues'):
                form_issues['layout_issues']['pages'].add(page['url'])
                form_issues['layout_issues']['domains'].add(domain)
                form_issues['layout_issues']['count'] += summary.get('inputsWithLayoutIssues', 0)
                
        except Exception as e:
            print(f"Error processing page {page.get('url', 'unknown')}: {str(e)}")
            continue

    # Create filtered list of issues that have affected pages
    active_issues = {flag: data for flag, data in form_issues.items() 
                    if len(data['pages']) > 0}

    if active_issues:
        # Create summary table
        summary_table = doc.add_table(rows=len(active_issues) + 1, cols=5)
        summary_table.style = 'Table Grid'

        # Set column headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Issue Type"
        headers[1].text = "Number of Occurrences"
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

        # Add statistics
        doc.add_paragraph()
        doc.add_paragraph("Form Statistics:", style='Normal')
        doc.add_paragraph(f"Total number of forms across all pages: {total_forms}")

    else:
        doc.add_paragraph("No form accessibility issues were found.")