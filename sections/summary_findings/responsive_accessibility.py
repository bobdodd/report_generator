"""
Summary findings for responsive accessibility tests across breakpoints
"""
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import logging
from typing import Dict, Any, List

from ..sections_header import create_section_heading
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

def add_responsive_accessibility_summary(document, results: Dict[str, Any], include_details: bool = True) -> None:
    """
    Add responsive accessibility summary to the report
    
    Args:
        document: Document object to add content to
        results: Test results dictionary
        include_details: Whether to include detailed results
    """
    # Create section heading
    create_section_heading(
        document, 
        "Responsive Accessibility Analysis", 
        "Evaluation of accessibility at different viewport sizes"
    )
    
    responsive_testing = results.get('responsive_testing', {})
    if not responsive_testing:
        add_paragraph(document, "No responsive testing data available.")
        return
    
    # Extract key information
    breakpoints = responsive_testing.get('breakpoints', [])
    consolidated_results = responsive_testing.get('consolidated', {})
    summary = consolidated_results.get('summary', {})
    
    # Add top-level summary paragraph
    total_issues = summary.get('totalIssues', 0)
    affected_breakpoints = summary.get('affectedBreakpoints', 0)
    tested_breakpoints = summary.get('totalBreakpoints', len(breakpoints))
    
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
    for bp in sorted(breakpoints):
        bp_text.append(f"{bp}px ({get_breakpoint_category(bp)})")
    
    if bp_text:
        add_paragraph(document, "The following responsive breakpoints were tested:")
        for text in bp_text:
            add_list_item(document, text)
    else:
        add_paragraph(document, "No breakpoints were tested.")
    
    # Add issues summary by test type
    add_subheading(document, "Responsive Issues Summary")
    
    tests_summary = consolidated_results.get('testsSummary', {})
    if not tests_summary:
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
        if test_key not in tests_summary:
            continue
            
        test_data = tests_summary[test_key]
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
        
        for test_key, test_info in test_categories.items():
            if test_key in tests_summary and tests_summary[test_key].get('issueCount', 0) > 0:
                add_subheading_h4(document, test_info['name'])
                add_paragraph(document, test_info['description'])
                
                # Add specific examples if available
                test_data = tests_summary[test_key]
                for bp, elements in test_data.get('elementsByBreakpoint', {}).items():
                    if elements and len(elements) > 0:
                        example = elements[0]
                        add_list_item(
                            document, 
                            f"Example at {bp}px: {example.get('element', 'Unknown element')} "
                            f"({example.get('details', 'No details available')})"
                        )
                        break
    else:
        add_paragraph(document, "No significant responsive accessibility issues were identified.")
    
    # Add recommendations section
    add_subheading(document, "Recommendations")
    
    issues_by_type = consolidated_results.get('issuesByType', {})
    if not issues_by_type:
        add_paragraph(
            document, 
            "The site demonstrates good responsive accessibility practices. Continue to test across "
            "different viewport sizes when making significant layout changes."
        )
        return
    
    # Standard recommendations based on issue types
    recommendations = []
    
    if 'smallTouchTarget' in issues_by_type or 'adjacentTouchTargets' in issues_by_type:
        recommendations.append(
            "Increase touch target sizes to at least 44x44 pixels on mobile breakpoints. "
            "Ensure sufficient spacing between interactive elements."
        )
    
    if 'overflow' in issues_by_type or any(k.startswith('overflow') for k in issues_by_type):
        recommendations.append(
            "Fix content that overflows the viewport at specific breakpoints. "
            "Use relative units (%, em, rem) and ensure content properly wraps."
        )
    
    if 'smallText' in issues_by_type:
        recommendations.append(
            "Ensure text remains readable at all viewport sizes. "
            "Minimum text size should be 12px, with 16px recommended for body text."
        )
    
    if any(k.startswith('fixedPosition') for k in issues_by_type):
        recommendations.append(
            "Review fixed position elements that may cause issues on small viewports. "
            "Consider adjusting or removing fixed elements at mobile breakpoints."
        )
    
    if 'visualDomMismatch' in issues_by_type or 'cssOrderProperty' in issues_by_type:
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