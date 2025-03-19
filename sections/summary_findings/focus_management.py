from report_styling import format_table_text

def add_focus_management_section(doc, db_connection, total_domains):
    """Add the Focus Management (General) section to the summary findings"""
    h2 = doc.add_heading('Focus Management (General)', level=2)
    h2.style = doc.styles['Heading 2']

    # Query for pages with focus management information
    pages_with_focus = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.focus_management.focus_management": {"$exists": True}
        },
        {
            "url": 1,
            "results.accessibility.tests.focus_management.focus_management": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize tracking
    site_data = {}
    url_data = {}  # Track data by individual URL
    total_interactive_elements = 0
    total_violations = 0
    total_breakpoints_tested = 0

    # Process each page
    for page in pages_with_focus:
        try:
            url = page['url']
            domain = url.replace('http://', '').replace('https://', '').split('/')[0]
            focus_data = page['results']['accessibility']['tests']['focus_management']['focus_management']
            
            # Initialize domain tracking
            if domain not in site_data:
                site_data[domain] = {
                    "total_violations": 0,
                    "breakpoints_tested": 0,
                    "urls": set(),
                    "tests": {
                        "focus_outline_presence": {"violations": 0, "elements": set()},
                        "focus_outline_contrast": {"violations": 0, "elements": set()},
                        "focus_outline_offset": {"violations": 0, "elements": set()},
                        "hover_feedback": {"violations": 0, "elements": set()},
                        "focus_obscurement": {"violations": 0, "elements": set()},
                        "anchor_target_tabindex": {"violations": 0, "elements": set()}
                    }
                }
            
            # Initialize URL tracking
            if url not in url_data:
                url_data[url] = {
                    "domain": domain,
                    "total_violations": 0,
                    "breakpoints_tested": 0,
                    "tests": {
                        "focus_outline_presence": {"violations": 0, "elements": set()},
                        "focus_outline_contrast": {"violations": 0, "elements": set()},
                        "focus_outline_offset": {"violations": 0, "elements": set()},
                        "hover_feedback": {"violations": 0, "elements": set()},
                        "focus_obscurement": {"violations": 0, "elements": set()},
                        "anchor_target_tabindex": {"violations": 0, "elements": set()}
                    }
                }

            # Get metadata
            metadata = focus_data.get('metadata', {})
            url_violations = metadata.get('total_violations_found', 0)
            
            total_violations += url_violations
            site_data[domain]["total_violations"] += url_violations
            url_data[url]["total_violations"] = url_violations
            
            breakpoints_tested = metadata.get('total_breakpoints_tested', 0)
            total_breakpoints_tested += breakpoints_tested
            site_data[domain]["breakpoints_tested"] = max(site_data[domain]["breakpoints_tested"], breakpoints_tested)
            url_data[url]["breakpoints_tested"] = breakpoints_tested
            
            # Add URL to domain list
            site_data[domain]["urls"].add(url)

            # Process each test
            tests = focus_data.get('tests', {})
            for test_name, test_data in tests.items():
                if test_name in site_data[domain]["tests"]:
                    # Get summary data
                    summary = test_data.get('summary', {})
                    violations = summary.get('total_violations', 0)
                    
                    # Update site data
                    site_data[domain]["tests"][test_name]["violations"] += violations
                    
                    # Update URL data
                    url_data[url]["tests"][test_name]["violations"] = violations
                    
                    # Track affected elements
                    elements = test_data.get('elements_affected', [])
                    if isinstance(elements, list):
                        site_data[domain]["tests"][test_name]["elements"].update(elements)
                        url_data[url]["tests"][test_name]["elements"].update(elements)

        except Exception as e:
            print(f"Error processing page {page.get('url', 'unknown')}: {str(e)}")
            continue

    if pages_with_focus:
        # Map test IDs to more readable names
        test_name_map = {
            "focus_outline_presence": "Missing Focus Outlines",
            "focus_outline_contrast": "Insufficient Outline Contrast",
            "focus_outline_offset": "Insufficient Outline Offset/Width",
            "hover_feedback": "Insufficient Hover Feedback",
            "focus_obscurement": "Obscured Focus Outlines",
            "anchor_target_tabindex": "Improper Local Target Configuration"
        }

        # Calculate totals for each test type
        test_totals = {}
        for test_id, display_name in test_name_map.items():
            total_violations = sum(site["tests"][test_id]["violations"] for site in site_data.values())
            affected_sites = sum(1 for site in site_data.values() if site["tests"][test_id]["violations"] > 0)
            affected_pages = sum(1 for url in url_data.values() if url["tests"][test_id]["violations"] > 0)
            
            # Count unique elements across all sites
            all_elements = set()
            for site in site_data.values():
                all_elements.update(site["tests"][test_id]["elements"])
            
            test_totals[test_id] = {
                "violations": total_violations,
                "affected_sites": affected_sites,
                "affected_pages": affected_pages,
                "unique_elements": len(all_elements)
            }
        
        # Create table for test summaries
        summary_table = doc.add_table(rows=len(test_name_map) + 1, cols=5)
        summary_table.style = 'Table Grid'
        
        # Add headers
        headers = summary_table.rows[0].cells
        headers[0].text = "Issue Type"
        headers[1].text = "Total Violations"
        headers[2].text = "Pages Affected"
        headers[3].text = "Sites Affected"
        headers[4].text = "% of Total Sites"
        
        # Add data for each test type
        row_idx = 1
        for test_id, display_name in test_name_map.items():
            totals = test_totals[test_id]
            row = summary_table.rows[row_idx].cells
            row[0].text = display_name
            row[1].text = str(totals["violations"])
            row[2].text = str(totals["affected_pages"])
            row[3].text = str(totals["affected_sites"])
            row[4].text = f"{(totals['affected_sites'] / len(site_data) * 100):.1f}%" if site_data else "0%"
            row_idx += 1
        
        format_table_text(summary_table)
    else:
        doc.add_paragraph("No focus management data available in the database.", style='Normal')
        