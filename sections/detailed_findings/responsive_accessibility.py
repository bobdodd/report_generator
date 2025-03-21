"""
Detailed findings for responsive accessibility tests across breakpoints
"""
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import logging
from typing import Dict, Any, List

from ...report_styling import (
    add_list_item, add_paragraph, add_subheading, add_subheading_h3, 
    add_subheading_h4, format_severity, add_table, add_hyperlink, 
    add_code_block, add_image_if_exists
)

def get_breakpoint_category(width: int) -> str:
    """
    Categorize a breakpoint width into a device category
    """
    if width <= 480:
        return "Mobile (Small)"
    elif width <= 768:
        return "Mobile (Large)/Tablet (Small)"
    elif width <= 1024:
        return "Tablet (Large)"
    elif width <= 1280:
        return "Desktop (Small)"
    else:
        return "Desktop (Large)"

def add_responsive_accessibility_detailed(document, db_connection, total_domains, screenshots_dir: str = None) -> None:
    """
    Add responsive accessibility detailed findings to the report
    
    Args:
        document: Document object to add content to
        db_connection: Database connection
        total_domains: Set of all domains analyzed
        screenshots_dir: Directory containing screenshots
    """
    # Create section heading
    document.add_page_break()
    h2 = document.add_heading('Responsive Accessibility Analysis', level=2)
    h2.style = document.styles['Heading 2']
    
    # Add subtitle
    sub_para = document.add_paragraph("Detailed evaluation of accessibility at different viewport sizes")
    sub_para.style = document.styles['Normal']
    for run in sub_para.runs:
        run.italic = True
    
    # Query for pages with responsive testing results
    pages_with_responsive_testing = list(db_connection.page_results.find(
        {"results.accessibility.responsive_testing": {"$exists": True}},
        {
            "url": 1,
            "results.accessibility.responsive_testing": 1,
            "_id": 0
        }
    ).sort("url", 1))
    
    if not pages_with_responsive_testing:
        add_paragraph(document, "No responsive testing data available across any pages.")
        return
    
    # Collect all breakpoints across all pages
    all_breakpoints = set()
    all_pages_data = []
    
    for page in pages_with_responsive_testing:
        url = page.get('url', 'Unknown URL')
        responsive_testing = page.get('results', {}).get('accessibility', {}).get('responsive_testing', {})
        
        if not responsive_testing:
            continue
            
        breakpoints = responsive_testing.get('breakpoints', [])
        all_breakpoints.update(breakpoints)
        
        # Store this page data for later use
        all_pages_data.append({
            'url': url,
            'responsive_testing': responsive_testing
        })
    
    sorted_breakpoints = sorted(list(all_breakpoints))
    
    # Add test methodology and documentation
    add_subheading(document, "Test Methodology")
    add_paragraph(
        document,
        "Responsive accessibility testing evaluates how well a site maintains accessibility across "
        "different viewport sizes. The tests adjust the browser viewport to match each breakpoint "
        "detected in the site's CSS media queries, then perform specific accessibility checks at each size."
    )
    
    # Add WCAG references
    add_paragraph(document, "Key WCAG Success Criteria evaluated:")
    wcag_criteria = [
        ("1.4.10 Reflow", "Content can be presented without loss of information or functionality, and without requiring scrolling in two dimensions"),
        ("1.4.4 Resize Text", "Text can be resized up to 200 percent without loss of content or functionality"),
        ("2.5.5 Target Size", "The size of the target for pointer inputs is at least 44 by 44 CSS pixels (Level AAA)"),
        ("1.3.2 Meaningful Sequence", "When the sequence in which content is presented affects its meaning, a correct reading sequence can be programmatically determined"),
        ("2.4.7 Focus Visible", "Any keyboard operable user interface has a mode of operation where the keyboard focus indicator is visible")
    ]
    
    for criterion, description in wcag_criteria:
        add_list_item(document, f"{criterion}: {description}")
    
    # Add breakpoints analysis
    add_subheading(document, "Breakpoints Analysis")
    
    # Format breakpoints into a table with categories and issues count
    headers = ["Breakpoint Width", "Device Category", "Issues Found", "Pages Affected"]
    rows = []
    
    # Count issues at each breakpoint across all pages
    breakpoint_issue_counts = {}
    breakpoint_page_counts = {}
    
    for page_data in all_pages_data:
        responsive_testing = page_data['responsive_testing']
        breakpoint_results = responsive_testing.get('breakpoint_results', {})
        
        for bp_str, bp_data in breakpoint_results.items():
            try:
                bp = int(bp_str)
            except ValueError:
                continue
                
            if bp not in breakpoint_issue_counts:
                breakpoint_issue_counts[bp] = 0
                breakpoint_page_counts[bp] = 0
            
            # Count issues in this breakpoint for this page
            issues_count = 0
            for test_name, test_data in bp_data.get('tests', {}).items():
                issues_count += len(test_data.get('issues', []))
            
            if issues_count > 0:
                breakpoint_issue_counts[bp] += issues_count
                breakpoint_page_counts[bp] += 1
    
    # Build the table rows
    for bp in sorted(list(all_breakpoints)):
        category = get_breakpoint_category(bp)
        issues_count = breakpoint_issue_counts.get(bp, 0)
        pages_count = breakpoint_page_counts.get(bp, 0)
        
        rows.append([
            f"{bp}px",
            category,
            str(issues_count),
            f"{pages_count} of {len(all_pages_data)}"
        ])
    
    if rows:
        add_table(document, headers, rows)
    else:
        add_paragraph(document, "No breakpoints were analyzed.")
    
    # Process detailed results for each test type
    # Build a summary of issues across all pages by test type
    test_summaries = {
        'overflow': {'issueCount': 0, 'affectedBreakpoints': set(), 'affectedPages': set()},
        'touchTargets': {'issueCount': 0, 'affectedBreakpoints': set(), 'affectedPages': set()},
        'fontScaling': {'issueCount': 0, 'affectedBreakpoints': set(), 'affectedPages': set()},
        'fixedPosition': {'issueCount': 0, 'affectedBreakpoints': set(), 'affectedPages': set()},
        'contentStacking': {'issueCount': 0, 'affectedBreakpoints': set(), 'affectedPages': set()}
    }
    
    # Collect examples of issues for each test type from all pages
    test_examples = {}
    
    # Process all pages to aggregate test data
    for page_data in all_pages_data:
        url = page_data['url']
        responsive_testing = page_data['responsive_testing']
        breakpoint_results = responsive_testing.get('breakpoint_results', {})
        consolidated = responsive_testing.get('consolidated', {})
        tests_summary = consolidated.get('testsSummary', {})
        
        # Aggregate from consolidated summaries
        for test_key in test_summaries.keys():
            if test_key in tests_summary:
                test_data = tests_summary[test_key]
                issue_count = test_data.get('issueCount', 0)
                
                if issue_count > 0:
                    test_summaries[test_key]['issueCount'] += issue_count
                    test_summaries[test_key]['affectedPages'].add(url)
                    test_summaries[test_key]['affectedBreakpoints'].update(
                        test_data.get('affectedBreakpoints', [])
                    )
        
        # Collect detailed examples from breakpoint results
        for bp_str, bp_data in breakpoint_results.items():
            try:
                bp = int(bp_str)
            except ValueError:
                continue
                
            tests = bp_data.get('tests', {})
            
            for test_key in test_summaries.keys():
                if test_key not in tests:
                    continue
                    
                test_result = tests[test_key]
                issues = test_result.get('issues', [])
                
                if issues and len(issues) > 0:
                    # Store the first new example we find for each test type
                    if test_key not in test_examples:
                        test_examples[test_key] = {
                            'url': url,
                            'breakpoint': bp,
                            'issues': issues[:3]  # Store up to 3 examples
                        }
    
    # Define test categories in a specific order with better names
    test_categories = {
        'overflow': {
            'name': 'Content Overflow',
            'description': 'Elements that overflow the viewport at specific breakpoints',
            'wcag': '1.4.10 (Reflow)'
        },
        'touchTargets': {
            'name': 'Touch Target Size',
            'description': 'Interactive elements that are too small for touch interaction',
            'wcag': '2.5.5 (Target Size)'
        },
        'fontScaling': {
            'name': 'Font Scaling',
            'description': 'Text that becomes too small at certain viewport sizes',
            'wcag': '1.4.4 (Resize Text)'
        },
        'fixedPosition': {
            'name': 'Fixed Position Elements',
            'description': 'Fixed elements that obscure content at certain viewport sizes',
            'wcag': '1.4.10 (Reflow), 2.4.7 (Focus Visible)'
        },
        'contentStacking': {
            'name': 'Content Stacking Order',
            'description': 'Issues with content reflow and reading order at different breakpoints',
            'wcag': '1.3.2 (Meaningful Sequence)'
        }
    }
    
    for test_key, test_info in test_categories.items():
        if test_key not in test_summaries or test_summaries[test_key]['issueCount'] == 0:
            continue
        
        test_data = test_summaries[test_key]
        issue_count = test_data['issueCount']
        affected_bps = sorted(list(test_data['affectedBreakpoints']))
        affected_pages = sorted(list(test_data['affectedPages']))
        
        # Create section for this test type
        add_subheading(document, test_info['name'])
        
        # Add description and WCAG references
        add_paragraph(
            document,
            f"{test_info['description']}. "
            f"This test evaluates compliance with WCAG {test_info['wcag']}."
        )
        
        # Detail affected pages and breakpoints
        add_paragraph(
            document,
            f"Found {issue_count} issues across {len(affected_pages)} pages and {len(affected_bps)} breakpoints."
        )
        
        # Show affected pages as list
        if affected_pages:
            add_subheading_h3(document, "Affected Pages")
            for url in affected_pages[:10]:  # Limit to 10 pages to avoid overwhelming the report
                add_list_item(document, url)
            
            if len(affected_pages) > 10:
                add_paragraph(
                    document,
                    f"... and {len(affected_pages) - 10} more pages affected."
                )
        
        # Detail affected breakpoints
        if affected_bps:
            add_subheading_h3(document, "Affected Breakpoints")
            
            # Create a table of breakpoints with their categories
            bp_headers = ["Breakpoint", "Device Category", "Issues"]
            bp_rows = []
            
            for bp in sorted([int(bp) for bp in affected_bps]):
                category = get_breakpoint_category(bp)
                issues_count = breakpoint_issue_counts.get(bp, 0)
                
                bp_rows.append([
                    f"{bp}px",
                    category,
                    str(issues_count)
                ])
            
            if bp_rows:
                add_table(document, bp_headers, bp_rows)
            
            # Show examples from this test type
            if test_key in test_examples:
                example_data = test_examples[test_key]
                issues = example_data['issues']
                url = example_data['url']
                bp = example_data['breakpoint']
                
                add_subheading_h3(document, f"Examples at {bp}px breakpoint from {url}")
                
                for issue in issues:
                    element_type = issue.get('element', 'Unknown element')
                    element_id = f" (id: {issue['id']})" if issue.get('id') else ""
                    details = issue.get('details', 'No details available')
                    severity = issue.get('severity', 'medium')
                    
                    # Format issue item with severity color
                    para = document.add_paragraph(style='List Bullet')
                    run = para.add_run(f"{element_type}{element_id}: ")
                    run.bold = True
                    
                    # Add severity indicator
                    severity_run = para.add_run(f"[{severity.upper()}] ")
                    format_severity(severity_run, severity)
                    
                    # Add issue details
                    para.add_run(details)
        
        # Add recommendations for this test type
        add_subheading_h3(document, "Recommendations")
        
        # Specific recommendations based on test type
        if test_key == 'overflow':
            add_list_item(
                document,
                "Use responsive design techniques like flexbox, CSS grid, or percentage-based widths to ensure content "
                "properly adjusts to different viewport sizes."
            )
            add_list_item(
                document,
                "Ensure images and media are responsive with max-width: 100% and height: auto."
            )
        elif test_key == 'touchTargets':
            add_list_item(
                document,
                "Increase touch target sizes to at least 44x44 pixels on mobile breakpoints."
            )
            add_list_item(
                document,
                "Ensure sufficient spacing between interactive elements (at least 8px)."
            )
            add_list_item(
                document,
                "Use CSS padding to increase the interactive area without changing the visual size if necessary."
            )
        elif test_key == 'fontScaling':
            add_list_item(
                document,
                "Set a minimum font size of at least 12px for all text."
            )
            add_list_item(
                document,
                "Use relative units like em or rem instead of px for text."
            )
            add_list_item(
                document,
                "Test with browser text zoom set to 200% to ensure content remains readable."
            )
        elif test_key == 'fixedPosition':
            add_list_item(
                document,
                "Consider removing fixed position elements at mobile breakpoints or ensure they take up minimal screen space."
            )
            add_list_item(
                document,
                "Verify that fixed elements don't obscure important content when the viewport size changes."
            )
            add_list_item(
                document,
                "For sticky headers or footers, ensure they don't take up more than 20% of the viewport height."
            )
        elif test_key == 'contentStacking':
            add_list_item(
                document,
                "Maintain a logical reading order when content reflows at different viewport sizes."
            )
            add_list_item(
                document,
                "Avoid using CSS order properties that change the visual presentation from the DOM order."
            )
            add_list_item(
                document,
                "Ensure that the responsive design maintains a coherent experience across all breakpoints."
            )
    
    # Find elements with issues across multiple breakpoints in all pages
    add_subheading(document, "Elements with Issues Across Multiple Breakpoints")
    
    # This requires more complex aggregation - we'll need to check actual issue details
    # For now, we'll just summarize the most problematic pages
    problem_pages = {}
    
    for page_data in all_pages_data:
        url = page_data['url']
        responsive_testing = page_data['responsive_testing']
        consolidated = responsive_testing.get('consolidated', {})
        
        if 'elements' in consolidated:
            elements = consolidated['elements']
            multi_breakpoint_elements = {
                k: v for k, v in elements.items() 
                if len(v.get('breakpoints', [])) > 1
            }
            
            if multi_breakpoint_elements:
                problem_pages[url] = len(multi_breakpoint_elements)
    
    if problem_pages:
        # Sort pages by number of problematic elements
        sorted_pages = sorted(problem_pages.items(), key=lambda x: x[1], reverse=True)
        
        headers = ["Page URL", "Elements with Cross-Breakpoint Issues"]
        rows = []
        
        for url, count in sorted_pages[:10]:  # Limit to top 10
            rows.append([url, str(count)])
        
        if rows:
            add_table(document, headers, rows)
        
        add_paragraph(
            document,
            "Pages listed above have elements with issues that persist across multiple breakpoints. "
            "These should be prioritized for remediation as they impact users on multiple device types."
        )
    else:
        add_paragraph(
            document,
            "No elements with issues across multiple breakpoints were identified."
        )
    
    # Add technical notes section
    add_subheading(document, "Technical Notes")
    add_paragraph(
        document,
        "Responsive accessibility testing was performed using automated viewport resizing to match "
        "the breakpoints detected in the site's CSS media queries. Each breakpoint was tested for "
        "specific issues that commonly affect users at different viewport sizes."
    )
    
    add_paragraph(
        document,
        "The testing process extracts breakpoints from the site's CSS, then systematically resizes "
        "the browser viewport to each width while maintaining the same height. At each breakpoint, "
        "the page is analyzed for overflow issues, touch target sizes, font scaling problems, fixed "
        "position elements, and content stacking order."
    )
    
    # Include reference to test documentation if available
    documentation = results.get('tests', {}).get('responsive_accessibility', {}).get('documentation', {})
    if documentation:
        test_name = documentation.get('testName', 'Responsive Accessibility Analysis')
        description = documentation.get('description', '')
        
        if description:
            add_paragraph(document)
            add_paragraph(document, f"Reference: {test_name}")
            add_paragraph(document, description)