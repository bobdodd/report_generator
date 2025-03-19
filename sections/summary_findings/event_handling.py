import traceback
from report_styling import format_table_text

def add_event_handling_section(doc, db_connection, total_domains):
    """Add the Event Handling section to the summary findings"""
    h2 = doc.add_heading('Event Handling and Keyboard Interaction', level=2)
    h2.style = doc.styles['Heading 2']

    # Query for pages with event information
    pages_with_events = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.events.events": {"$exists": True}
        },
        {
            "url": 1,
            "results.accessibility.tests.events.events": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Initialize tracking structures
    property_data = {
        # Event Types
        "event_mouse": {"name": "Mouse Events", "pages": set(), "domains": set(), "count": 0},
        "event_keyboard": {"name": "Keyboard Events", "pages": set(), "domains": set(), "count": 0},
        "event_focus": {"name": "Focus Events", "pages": set(), "domains": set(), "count": 0},
        "event_touch": {"name": "Touch Events", "pages": set(), "domains": set(), "count": 0},
        "event_timer": {"name": "Timer Events", "pages": set(), "domains": set(), "count": 0},
        "event_lifecycle": {"name": "Lifecycle Events", "pages": set(), "domains": set(), "count": 0},
        "event_other": {"name": "Other Events", "pages": set(), "domains": set(), "count": 0},
        
        # Tab Order
        "explicit_tabindex": {"name": "Explicit tabindex Usage", "pages": set(), "domains": set(), "count": 0},
        "visual_violations": {"name": "Visual Order Violations", "pages": set(), "domains": set(), "count": 0},
        "column_violations": {"name": "Column Order Violations", "pages": set(), "domains": set(), "count": 0},
        "negative_tabindex": {"name": "Negative Tabindex", "pages": set(), "domains": set(), "count": 0},
        "high_tabindex": {"name": "High Tabindex Values", "pages": set(), "domains": set(), "count": 0},
        
        # Interactive Elements
        "mouse_only": {"name": "Mouse-only Elements", "pages": set(), "domains": set(), "count": 0},
        "missing_tabindex": {"name": "Missing tabindex", "pages": set(), "domains": set(), "count": 0},
        "non_interactive": {"name": "Non-interactive with Handlers", "pages": set(), "domains": set(), "count": 0},
        
        # Modal Support
        "modals_no_escape": {"name": "Modals Missing Escape", "pages": set(), "domains": set(), "count": 0}
    }

    # Create detailed violation tracking organized by domain and URL
    domain_data = {}

    # Process each page
    for page in pages_with_events:
        try:
            url = page['url']
            
            domain = url.replace('http://', '').replace('https://', '').split('/')[0]
            event_data = page['results']['accessibility']['tests']['events']['events']
            
            # Initialize domain and URL tracking if needed
            if domain not in domain_data:
                domain_data[domain] = {
                    'urls': {}
                }
            
            # Initialize URL data
            domain_data[domain]['urls'][url] = {
                'event_types': {},
                'violations': {},
                'handlers_count': 0,
                'focusable_elements': 0,
                'total_violations': 0
            }
            
            # Get pageFlags data for the most reliable summary information
            pageFlags = event_data.get('pageFlags', {})
            details = pageFlags.get('details', {})
            
            # Track total handlers and violations
            total_handlers = details.get('totalHandlers', 0)
            total_violations = details.get('totalViolations', 0)
            domain_data[domain]['urls'][url]['handlers_count'] = total_handlers
            domain_data[domain]['urls'][url]['total_violations'] = total_violations

            # Track total focusable elements
            tab_order_data = details.get('tabOrder', {})
            focusable_elements = tab_order_data.get('totalFocusableElements', 0)
            domain_data[domain]['urls'][url]['focusable_elements'] = focusable_elements
            
            # Process event types using the updated structure
            by_type = details.get('byType', {})
            
            for event_type in ['mouse', 'keyboard', 'focus', 'touch', 'timer', 'lifecycle', 'other']:
                count = by_type.get(event_type, 0)
                
                if isinstance(count, list):
                    count = len(count)
                elif not isinstance(count, (int, float)):
                    try:
                        count = int(count or 0)
                    except (ValueError, TypeError):
                        count = 0
                
                # Track event type for this URL
                domain_data[domain]['urls'][url]['event_types'][event_type] = count
                
                if count > 0:
                    key = f"event_{event_type}"
                    property_data[key]['pages'].add(url)
                    property_data[key]['domains'].add(domain)
                    property_data[key]['count'] += count

            # Process violation counts by type
            violation_counts = details.get('violationCounts', {})
            
            # Tab order violations
            explicit_count = tab_order_data.get('elementsWithExplicitTabIndex', 0)
            visual_violations = violation_counts.get('visual-order', 0) or tab_order_data.get('visualOrderViolations', 0)
            column_violations = violation_counts.get('column-order', 0) or tab_order_data.get('columnOrderViolations', 0)
            
            # Track negative and high tabindex
            negative_tabindex = 1 if pageFlags.get('hasNegativeTabindex', False) else 0
            high_tabindex = 1 if pageFlags.get('hasHighTabindex', False) else 0
            
            # Track violations for this URL
            domain_data[domain]['urls'][url]['violations']['explicit_tabindex'] = explicit_count
            domain_data[domain]['urls'][url]['violations']['visual_order'] = visual_violations
            domain_data[domain]['urls'][url]['violations']['column_order'] = column_violations
            domain_data[domain]['urls'][url]['violations']['negative_tabindex'] = negative_tabindex
            domain_data[domain]['urls'][url]['violations']['high_tabindex'] = high_tabindex
            
            if explicit_count > 0:
                property_data['explicit_tabindex']['pages'].add(url)
                property_data['explicit_tabindex']['domains'].add(domain)
                property_data['explicit_tabindex']['count'] += explicit_count
                
            if visual_violations > 0:
                property_data['visual_violations']['pages'].add(url)
                property_data['visual_violations']['domains'].add(domain)
                property_data['visual_violations']['count'] += visual_violations
                
            if column_violations > 0:
                property_data['column_violations']['pages'].add(url)
                property_data['column_violations']['domains'].add(domain)
                property_data['column_violations']['count'] += column_violations
                
            if negative_tabindex > 0:
                property_data['negative_tabindex']['pages'].add(url)
                property_data['negative_tabindex']['domains'].add(domain)
                property_data['negative_tabindex']['count'] += negative_tabindex
                
            if high_tabindex > 0:
                property_data['high_tabindex']['pages'].add(url)
                property_data['high_tabindex']['domains'].add(domain)
                property_data['high_tabindex']['count'] += high_tabindex

            # Process element violations
            mouse_only = violation_counts.get('mouse-only', 0) or details.get('mouseOnlyElements', {}).get('count', 0)
            missing_tabindex = violation_counts.get('missing-tabindex', 0) or details.get('missingTabindex', 0)
            non_interactive = details.get('nonInteractiveWithHandlers', 0)
            modals_without_escape = violation_counts.get('modal-without-escape', 0)
            
            # Track violations for this URL
            domain_data[domain]['urls'][url]['violations']['mouse_only'] = mouse_only
            domain_data[domain]['urls'][url]['violations']['missing_tabindex'] = missing_tabindex
            domain_data[domain]['urls'][url]['violations']['non_interactive'] = non_interactive
            domain_data[domain]['urls'][url]['violations']['modals_no_escape'] = modals_without_escape
            
            if mouse_only > 0:
                property_data['mouse_only']['pages'].add(url)
                property_data['mouse_only']['domains'].add(domain)
                property_data['mouse_only']['count'] += mouse_only
                
            if missing_tabindex > 0:
                property_data['missing_tabindex']['pages'].add(url)
                property_data['missing_tabindex']['domains'].add(domain)
                property_data['missing_tabindex']['count'] += missing_tabindex
                
            if non_interactive > 0:
                property_data['non_interactive']['pages'].add(url)
                property_data['non_interactive']['domains'].add(domain)
                property_data['non_interactive']['count'] += non_interactive
                
            if modals_without_escape > 0:
                property_data['modals_no_escape']['pages'].add(url)
                property_data['modals_no_escape']['domains'].add(domain)
                property_data['modals_no_escape']['count'] += modals_without_escape

        except Exception as e:
            print(f"Error processing page {url}:")
            print("Exception:", str(e))
            traceback.print_exc()
            continue

    if pages_with_events:
        # Create summary table
        doc.add_heading('Event Handling Summary', level=3)
        
        # Calculate number of rows needed
        rows_needed = 1  # Header row
        
        # Event Types section (header + all event types)
        rows_needed += 1  # Section header
        rows_needed += len([k for k in property_data.keys() if k.startswith('event_')])
        
        # Tab Order section (header + 5 items)
        rows_needed += 1  # Section header
        rows_needed += 5  # explicit_tabindex, visual_violations, column_violations, negative_tabindex, high_tabindex
        
        # Interactive Elements section (header + 3 items)
        rows_needed += 1  # Section header
        rows_needed += 3  # mouse_only, missing_tabindex, non_interactive
        
        # Modal Support section (header + 1 item)
        rows_needed += 1  # Section header
        rows_needed += 1  # modals_no_escape

        # Create table with correct number of rows
        table = doc.add_table(rows=rows_needed, cols=4)
        table.style = 'Table Grid'
        
        # Add headers
        headers = table.rows[0].cells
        headers[0].text = "Property"
        headers[1].text = "Occurrences"
        headers[2].text = "Pages Affected"
        headers[3].text = "% of Sites"
        
        current_row = 1
        
        # Add Event Types section
        row = table.rows[current_row].cells
        row[0].text = "Event Types:"
        current_row += 1
        
        for key, data in sorted([(k, v) for k, v in property_data.items() if k.startswith('event_')], 
                            key=lambda x: x[1]['count'], reverse=True):
            row = table.rows[current_row].cells
            row[0].text = "  " + data['name']
            row[1].text = str(data['count'])
            row[2].text = str(len(data['pages']))
            row[3].text = f"{(len(data['domains']) / len(total_domains) * 100):.1f}%"
            current_row += 1
        
        # Add Tab Order section
        row = table.rows[current_row].cells
        row[0].text = "Tab Order:"
        current_row += 1
        
        for key in ['explicit_tabindex', 'visual_violations', 'column_violations', 'negative_tabindex', 'high_tabindex']:
            row = table.rows[current_row].cells
            row[0].text = "  " + property_data[key]['name']
            row[1].text = str(property_data[key]['count'])
            row[2].text = str(len(property_data[key]['pages']))
            row[3].text = f"{(len(property_data[key]['domains']) / len(total_domains) * 100):.1f}%"
            current_row += 1
        
        # Add Interactive Elements section
        row = table.rows[current_row].cells
        row[0].text = "Interactive Elements:"
        current_row += 1
        
        for key in ['mouse_only', 'missing_tabindex', 'non_interactive']:
            row = table.rows[current_row].cells
            row[0].text = "  " + property_data[key]['name']
            row[1].text = str(property_data[key]['count'])
            row[2].text = str(len(property_data[key]['pages']))
            row[3].text = f"{(len(property_data[key]['domains']) / len(total_domains) * 100):.1f}%"
            current_row += 1
        
        # Add Modal Support section
        row = table.rows[current_row].cells
        row[0].text = "Modal Support:"
        current_row += 1
        
        key = 'modals_no_escape'
        row = table.rows[current_row].cells
        row[0].text = "  " + property_data[key]['name']
        row[1].text = str(property_data[key]['count'])
        row[2].text = str(len(property_data[key]['pages']))
        row[3].text = f"{(len(property_data[key]['domains']) / len(total_domains) * 100):.1f}%"

        format_table_text(table)
        
    else:
        doc.add_paragraph("No event handling data was found.")
        