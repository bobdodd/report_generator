"""
Summary findings for responsive accessibility tests across breakpoints
"""
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import logging
from typing import Dict, Any, List

from report_styling import (
    add_list_item, add_paragraph, add_subheading, add_subheading_h3, 
    add_subheading_h4, format_severity, add_table, add_hyperlink
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

def add_responsive_accessibility_summary(document, db_connection, total_domains):
    """
    Add responsive accessibility summary to the report
    
    Args:
        document: Document object to add content to
        db_connection: Database connection
        total_domains: Set of all domains analyzed
    """
    # Create section heading
    document.add_paragraph()
    h2 = document.add_heading('Responsive Accessibility Analysis Summary', level=2)
    h2.style = document.styles['Heading 2']
    
    # Add subtitle
    sub_para = document.add_paragraph("Evaluation of accessibility at different viewport sizes")
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
    
    # Collect breakpoints and issues across all pages
    all_breakpoints = set()
    total_issues = 0
    affected_breakpoints = 0
    pages_with_issues = 0
    pages_with_skipped_tests = 0
    pages_with_actual_tests = 0
    
    for page in pages_with_responsive_testing:
        responsive_testing = page.get('results', {}).get('accessibility', {}).get('responsive_testing', {})
        
        # Check if responsive testing was skipped due to no breakpoints
        if responsive_testing.get('status') == 'skipped':
            pages_with_skipped_tests += 1
            continue
            
        pages_with_actual_tests += 1
        
        # Collect breakpoints
        breakpoints = responsive_testing.get('breakpoints', [])
        all_breakpoints.update(breakpoints)
        
        # Collect issues
        consolidated = responsive_testing.get('consolidated', {})
        summary = consolidated.get('summary', {})
        
        if summary.get('totalIssues', 0) > 0:
            total_issues += summary.get('totalIssues', 0)
            affected_breakpoints = max(affected_breakpoints, summary.get('affectedBreakpoints', 0))
            pages_with_issues += 1
    
    # If all pages were skipped due to no breakpoints found
    if pages_with_skipped_tests > 0 and pages_with_actual_tests == 0:
        add_paragraph(
            document,
            f"Responsive testing was skipped on {pages_with_skipped_tests} pages because no CSS media query breakpoints were found. "
            "The site may not be using responsive design techniques with CSS media queries, or the media queries do not contain width-based breakpoints."
        )
        return
    
    # Convert to list and sort breakpoints for display
    sorted_breakpoints = sorted(list(all_breakpoints))
    tested_breakpoints = len(sorted_breakpoints)
    
    if total_issues > 0:
        add_paragraph(
            document, 
            f"Found {total_issues} responsive issues across {affected_breakpoints}/{tested_breakpoints} breakpoints."
        )
    else:
        add_paragraph(
            document, 
            f"No responsive issues found across {tested_breakpoints} breakpoints."
        )
    
    # Add breakpoints summary in a compact table
    add_subheading_h3(document, "Tested Breakpoints")
    
    if sorted_breakpoints:
        # Create a simple, compact table of breakpoints
        headers = ["Breakpoint", "Device Category"]
        rows = []
        
        for bp in sorted_breakpoints:
            rows.append([
                f"{bp}px",
                get_breakpoint_category(bp)
            ])
            
        if rows:
            add_table(document, headers, rows)
    else:
        add_paragraph(document, "No breakpoints were tested.")
    
    # Add issues summary by test type
    add_subheading_h3(document, "Responsive Issues Summary")
    
    # Collect test summaries across all pages
    all_test_summaries = {
        'overflow': {'issueCount': 0, 'affectedBreakpoints': set()},
        'touchTargets': {'issueCount': 0, 'affectedBreakpoints': set()},
        'fontScaling': {'issueCount': 0, 'affectedBreakpoints': set()},
        'fixedPosition': {'issueCount': 0, 'affectedBreakpoints': set()},
        'contentStacking': {'issueCount': 0, 'affectedBreakpoints': set()}
    }
    
    # Aggregate test summaries across all pages
    for page in pages_with_responsive_testing:
        responsive_testing = page.get('results', {}).get('accessibility', {}).get('responsive_testing', {})
        consolidated = responsive_testing.get('consolidated', {})
        tests_summary = consolidated.get('testsSummary', {})
        
        if not tests_summary:
            continue
            
        for test_key in all_test_summaries.keys():
            if test_key in tests_summary:
                test_data = tests_summary[test_key]
                # Make sure test_data is a dictionary
                if isinstance(test_data, dict):
                    all_test_summaries[test_key]['issueCount'] += test_data.get('issueCount', 0)
                    affected_bps = test_data.get('affectedBreakpoints', [])
                    # Make sure affected_bps is an iterable
                    if hasattr(affected_bps, '__iter__') and not isinstance(affected_bps, str):
                        all_test_summaries[test_key]['affectedBreakpoints'].update(affected_bps)
                    elif affected_bps:  # Handle single value
                        all_test_summaries[test_key]['affectedBreakpoints'].add(affected_bps)
    
    # Convert sets to lists for the template
    for test_key in all_test_summaries:
        if 'affectedBreakpoints' in all_test_summaries[test_key]:
            all_test_summaries[test_key]['affectedBreakpoints'] = sorted(
                list(all_test_summaries[test_key]['affectedBreakpoints'])
            )
    
    # Check if we have any issues at all
    if not any(test_data['issueCount'] > 0 for test_data in all_test_summaries.values()):
        add_paragraph(document, "No responsive testing issues were identified.")
        return
    
    # Create a table for issue types
    headers = ["Test Type", "Issues", "Affected Breakpoints", "Severity"]
    rows = []
    
    # Define test categories in a specific order with better names
    test_categories = {
        'overflow': {
            'name': 'Content Overflow',
            'description': 'Elements that overflow the viewport at specific breakpoints'
        },
        'touchTargets': {
            'name': 'Touch Target Size',
            'description': 'Interactive elements that are too small for touch interaction'
        },
        'fontScaling': {
            'name': 'Font Scaling',
            'description': 'Text that becomes too small at certain viewport sizes'
        },
        'fixedPosition': {
            'name': 'Fixed Position Elements',
            'description': 'Fixed elements that obscure content at certain viewport sizes'
        },
        'contentStacking': {
            'name': 'Content Stacking Order',
            'description': 'Issues with content reflow and reading order at different breakpoints'
        }
    }
    
    for test_key, test_info in test_categories.items():
        if test_key not in all_test_summaries:
            continue
            
        test_data = all_test_summaries[test_key]
        issue_count = test_data.get('issueCount', 0)
        
        if issue_count == 0:
            continue
            
        affected_bps = test_data.get('affectedBreakpoints', [])
        bp_ranges = []
        
        # Group breakpoints into ranges for cleaner display
        if affected_bps:
            affected_bps = sorted([int(bp) for bp in affected_bps])
            bp_text = []
            for bp in affected_bps:
                category = get_breakpoint_category(bp)
                bp_text.append(f"{bp}px ({category})")
            bp_ranges = ", ".join(bp_text)  # Show all breakpoints
        
        # Determine severity based on issue count and types
        severity = "Low"
        if test_key in ['overflow', 'touchTargets'] and issue_count > 3:
            severity = "High"
        elif issue_count > 5:
            severity = "Medium"
        
        rows.append([
            test_info['name'],
            str(issue_count),
            bp_ranges if bp_ranges else "None",
            severity
        ])
    
    if rows:
        table = add_table(document, headers, rows)
    else:
        add_paragraph(document, "No significant responsive accessibility issues were identified.")
    
    # Add recommendations section
    add_subheading_h3(document, "Recommendations")
    
    # Collect top issues for a concise recommendation
    top_issues = []
    for test_key, test_info in test_categories.items():
        if test_key in all_test_summaries and all_test_summaries[test_key].get('issueCount', 0) > 0:
            top_issues.append((test_key, all_test_summaries[test_key].get('issueCount', 0)))
    
    # Sort by issue count (descending)
    top_issues.sort(key=lambda x: x[1], reverse=True)
    
    # Get top 3 issues for recommendations
    top_3_issues = [issue[0] for issue in top_issues[:3]]
    
    # Standard recommendations based on test types with issues
    recommendations = []
    
    if 'touchTargets' in top_3_issues:
        recommendations.append(
            "Increase touch target sizes to at least 44x44 pixels on mobile breakpoints."
        )
    
    if 'overflow' in top_3_issues:
        recommendations.append(
            "Fix content that overflows the viewport by using responsive units (%, em, rem)."
        )
    
    if 'fontScaling' in top_3_issues:
        recommendations.append(
            "Ensure text remains readable with minimum size of 12px at all viewport sizes."
        )
    
    if 'fixedPosition' in top_3_issues:
        recommendations.append(
            "Review fixed position elements that may cause issues on small viewports."
        )
    
    if 'contentStacking' in top_3_issues:
        recommendations.append(
            "Ensure content maintains a logical reading order when it reflows at different widths."
        )
    
    # Add generic recommendation if needed to get to 3 recommendations
    if len(recommendations) < 3:
        recommendations.append(
            "Test content at all key breakpoints to ensure consistent experience across devices."
        )
    
    # Only show top 3 recommendations
    for recommendation in recommendations[:3]:
        add_list_item(document, recommendation)