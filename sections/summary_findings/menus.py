from report_styling import format_table_text

def add_menus_section(doc, db_connection, total_domains):
    """Add the Menus section to the summary findings"""
    h2 = doc.add_heading('Menus', level=2)
    h2.style = doc.styles['Heading 2']

    # Query for pages with menu issues
    pages_with_menu_issues = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.menus.menus.pageFlags": {"$exists": True},
            "$or": [
                {"results.accessibility.tests.menus.menus.pageFlags.hasInvalidMenuRoles": True},
                {"results.accessibility.tests.menus.menus.pageFlags.hasMenusWithoutCurrent": True},
                {"results.accessibility.tests.menus.menus.pageFlags.hasUnnamedMenus": True},
                {"results.accessibility.tests.menus.menus.pageFlags.hasDuplicateMenuNames": True}
            ]
        },
        {
            "url": 1,
            "results.accessibility.tests.menus.menus": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize counters for each issue type
    menu_issues = {
        "invalidRoles": {"name": "Invalid menu roles", "pages": set(), "domains": set(), "count": 0},
        "menusWithoutCurrent": {"name": "Missing current page indicators", "pages": set(), "domains": set(), "count": 0},
        "unnamedMenus": {"name": "Unnamed menus", "pages": set(), "domains": set(), "count": 0},
        "duplicateNames": {"name": "Duplicate menu names", "pages": set(), "domains": set(), "count": 0}
    }

    # Count issues
    total_menus = 0
    for page in pages_with_menu_issues:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        menu_data = page['results']['accessibility']['tests']['menus']['menus']
        flags = menu_data['pageFlags']
        details = menu_data['pageFlags']['details']
        
        total_menus += details.get('totalMenus', 0)
        
        # Check each type of issue
        if flags.get('hasInvalidMenuRoles'):
            menu_issues['invalidRoles']['pages'].add(page['url'])
            menu_issues['invalidRoles']['domains'].add(domain)
            menu_issues['invalidRoles']['count'] += details.get('invalidRoles', 0)
            
        if flags.get('hasMenusWithoutCurrent'):
            menu_issues['menusWithoutCurrent']['pages'].add(page['url'])
            menu_issues['menusWithoutCurrent']['domains'].add(domain)
            menu_issues['menusWithoutCurrent']['count'] += details.get('menusWithoutCurrent', 0)
            
        if flags.get('hasUnnamedMenus'):
            menu_issues['unnamedMenus']['pages'].add(page['url'])
            menu_issues['unnamedMenus']['domains'].add(domain)
            menu_issues['unnamedMenus']['count'] += details.get('unnamedMenus', 0)
            
        if flags.get('hasDuplicateMenuNames'):
            menu_issues['duplicateNames']['pages'].add(page['url'])
            menu_issues['duplicateNames']['domains'].add(domain)

    # Create filtered list of issues that have affected pages
    active_issues = {flag: data for flag, data in menu_issues.items() 
                    if len(data['pages']) > 0}

    if active_issues:
        # Create summary table
        summary_table = doc.add_table(rows=len(active_issues) + 1, cols=5)
        summary_table.style = 'Table Grid'

        # Set column headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Menu Issue"
        headers[1].text = "Number of Occurrences"
        headers[2].text = "Pages Affected"
        headers[3].text = "Sites Affected"
        headers[4].text = "% of Total Sites"

        # Add data
        for i, (flag, data) in enumerate(active_issues.items(), 1):
            row = summary_table.rows[i].cells
            row[0].text = data['name']
            row[1].text = str(data['count']) if flag != 'duplicateNames' else 'N/A'
            row[2].text = str(len(data['pages']))
            row[3].text = str(len(data['domains']))
            row[4].text = f"{(len(data['domains']) / len(total_domains) * 100):.1f}%"

        # Format the table text
        format_table_text(summary_table)

    else:
        doc.add_paragraph("No navigation menu accessibility issues were found.")
        