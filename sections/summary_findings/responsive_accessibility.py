"""
Summary findings for responsive accessibility tests across breakpoints
"""
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import logging
from typing import Dict, Any, List

from ...report_styling import (
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
    h2 = document.add_heading('Responsive Accessibility Analysis', level=2)
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
    
    for page in pages_with_responsive_testing:
        responsive_testing = page.get('results', {}).get('accessibility', {}).get('responsive_testing', {})
        
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
    
    # Convert to list and sort breakpoints for display
    sorted_breakpoints = sorted(list(all_breakpoints))
    tested_breakpoints = len(sorted_breakpoints)
    
    if total_issues > 0:
        add_paragraph(
            document, 
            f"Responsive accessibility testing identified {total_issues} issues across "
            f"{affected_breakpoints} of {tested_breakpoints} tested breakpoints. "
            f"These issues may impact users on specific device sizes."
        )
    else:
        add_paragraph(
            document, 
            f"Responsive accessibility testing across {tested_breakpoints} breakpoints found no significant issues. "
            f"The site appears to maintain accessibility at different viewport sizes."
        )
    
    # Add breakpoints summary
    add_subheading(document, "Tested Breakpoints")
    
    bp_text = []
    for bp in sorted_breakpoints:
        bp_text.append(f"{bp}px ({get_breakpoint_category(bp)})")
    
    if bp_text:
        add_paragraph(document, "The following responsive breakpoints were tested:")
        for text in bp_text:
            add_list_item(document, text)
    else:
        add_paragraph(document, "No breakpoints were tested.")
    
    # Add issues summary by test type
    add_subheading(document, "Responsive Issues Summary")
    
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
                all_test_summaries[test_key]['issueCount'] += test_data.get('issueCount', 0)
                all_test_summaries[test_key]['affectedBreakpoints'].update(
                    test_data.get('affectedBreakpoints', [])
                )
    
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
            bp_ranges = ", ".join(bp_text)
        
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
        
        # Add test descriptions after the table
        add_paragraph(document)
        add_subheading_h3(document, "Test Descriptions")
        
        # Find examples from actual page results for each test type
        examples_by_test = {}
        
        for page in pages_with_responsive_testing[:5]:  # Limit to first 5 pages for efficiency
            responsive_testing = page.get('results', {}).get('accessibility', {}).get('responsive_testing', {})
            breakpoint_results = responsive_testing.get('breakpoint_results', {})
            
            for bp, bp_data in breakpoint_results.items():
                tests = bp_data.get('tests', {})
                
                for test_key in test_categories.keys():
                    if test_key in tests and test_key not in examples_by_test:
                        test_result = tests[test_key]
                        issues = test_result.get('issues', [])
                        
                        if issues and len(issues) > 0:
                            examples_by_test[test_key] = {
                                'url': page.get('url', 'Unknown URL'),
                                'breakpoint': bp,
                                'example': issues[0]
                            }
        
        # Display test descriptions with examples where available
        for test_key, test_info in test_categories.items():
            if test_key in all_test_summaries and all_test_summaries[test_key].get('issueCount', 0) > 0:
                add_subheading_h4(document, test_info['name'])
                add_paragraph(document, test_info['description'])
                
                # Add specific examples if available
                if test_key in examples_by_test:
                    example_data = examples_by_test[test_key]
                    example = example_data['example']
                    add_list_item(
                        document, 
                        f"Example at {example_data['breakpoint']}px: {example.get('element', 'Unknown element')} "
                        f"({example.get('details', 'No details available')})"
                    )
                    add_list_item(
                        document,
                        f"Found on: {example_data['url']}"
                    )
    else:
        add_paragraph(document, "No significant responsive accessibility issues were identified.")
    
    # Add recommendations section
    add_subheading(document, "Recommendations")
    
    # Determine recommendations based on the aggregated test results
    if not any(test_data['issueCount'] > 0 for test_data in all_test_summaries.values()):
        add_paragraph(
            document, 
            "The site demonstrates good responsive accessibility practices. Continue to test across "
            "different viewport sizes when making significant layout changes."
        )
        return
    
    # Standard recommendations based on test types with issues
    recommendations = []
    
    if all_test_summaries.get('touchTargets', {}).get('issueCount', 0) > 0:
        recommendations.append(
            "Increase touch target sizes to at least 44x44 pixels on mobile breakpoints. "
            "Ensure sufficient spacing between interactive elements."
        )
    
    if all_test_summaries.get('overflow', {}).get('issueCount', 0) > 0:
        recommendations.append(
            "Fix content that overflows the viewport at specific breakpoints. "
            "Use relative units (%, em, rem) and ensure content properly wraps."
        )
    
    if all_test_summaries.get('fontScaling', {}).get('issueCount', 0) > 0:
        recommendations.append(
            "Ensure text remains readable at all viewport sizes. "
            "Minimum text size should be 12px, with 16px recommended for body text."
        )
    
    if all_test_summaries.get('fixedPosition', {}).get('issueCount', 0) > 0:
        recommendations.append(
            "Review fixed position elements that may cause issues on small viewports. "
            "Consider adjusting or removing fixed elements at mobile breakpoints."
        )
    
    if all_test_summaries.get('contentStacking', {}).get('issueCount', 0) > 0:
        recommendations.append(
            "Ensure content maintains a logical reading order when it reflows at different widths. "
            "Avoid using CSS that changes the visual order from the DOM order."
        )
    
    # Add generic recommendation if none of the above applied
    if not recommendations:
        recommendations.append(
            "Review and test content at all key breakpoints to ensure a consistent user experience "
            "across different device sizes."
        )
    
    for recommendation in recommendations:
        add_list_item(document, recommendation)
        
    # Add note about testing approach
    add_paragraph(document)
    add_paragraph(
        document,
        "Responsive testing was performed at multiple breakpoints using automated viewport resizing. "
        f"Tests were run on {len(pages_with_responsive_testing)} pages across {tested_breakpoints} different viewport widths."
    )