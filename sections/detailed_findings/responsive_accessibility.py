"""
Detailed findings for responsive accessibility tests across breakpoints
"""
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import logging
from typing import Dict, Any, List

from ..sections_header import create_section_heading
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

def add_responsive_accessibility_detailed(document, results: Dict[str, Any], screenshots_dir: str = None) -> None:
    """
    Add responsive accessibility detailed findings to the report
    
    Args:
        document: Document object to add content to
        results: Test results dictionary
        screenshots_dir: Directory containing screenshots
    """
    # Create section heading
    create_section_heading(
        document, 
        "Responsive Accessibility Analysis", 
        "Detailed evaluation of accessibility at different viewport sizes"
    )
    
    responsive_testing = results.get('responsive_testing', {})
    if not responsive_testing:
        add_paragraph(document, "No responsive testing data available.")
        return
    
    # Extract key information
    breakpoints = responsive_testing.get('breakpoints', [])
    consolidated_results = responsive_testing.get('consolidated', {})
    breakpoint_results = responsive_testing.get('breakpoint_results', {})
    
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
    
    # Format breakpoints into a table with categories
    headers = ["Breakpoint Width", "Device Category", "Issues Found"]
    rows = []
    
    for bp in sorted([int(bp) for bp in breakpoints]):
        bp_str = str(bp)
        category = get_breakpoint_category(bp)
        
        # Count issues at this breakpoint
        issues_count = 0
        if bp_str in breakpoint_results:
            bp_data = breakpoint_results[bp_str]
            for test_name, test_data in bp_data.get('tests', {}).items():
                issues_count += len(test_data.get('issues', []))
        
        rows.append([
            f"{bp}px",
            category,
            str(issues_count)
        ])
    
    if rows:
        add_table(document, headers, rows)
    else:
        add_paragraph(document, "No breakpoints were analyzed.")
    
    # Process detailed results for each test type
    tests_summary = consolidated_results.get('testsSummary', {})
    
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
        if test_key not in tests_summary or tests_summary[test_key].get('issueCount', 0) == 0:
            continue
        
        test_data = tests_summary[test_key]
        issue_count = test_data.get('issueCount', 0)
        affected_bps = test_data.get('affectedBreakpoints', [])
        
        # Create section for this test type
        add_subheading(document, test_info['name'])
        
        # Add description and WCAG references
        add_paragraph(
            document,
            f"{test_info['description']}. "
            f"This test evaluates compliance with WCAG {test_info['wcag']}."
        )
        
        # Summary of findings for this test type
        add_paragraph(
            document,
            f"Found {issue_count} issues across {len(affected_bps)} breakpoints."
        )
        
        # Detail affected breakpoints
        if affected_bps:
            add_subheading_h3(document, "Affected Breakpoints")
            for bp in sorted([int(bp) for bp in affected_bps]):
                category = get_breakpoint_category(bp)
                
                # Get issues for this breakpoint and test
                bp_issues = []
                if str(bp) in breakpoint_results:
                    bp_data = breakpoint_results[str(bp)]
                    if 'tests' in bp_data and test_key in bp_data['tests']:
                        test_result = bp_data['tests'][test_key]
                        bp_issues = test_result.get('issues', [])
                
                # Add breakpoint heading with issue count
                add_subheading_h4(document, f"{bp}px ({category}) - {len(bp_issues)} issues")
                
                # Add details for up to 3 issues per breakpoint to avoid overwhelming the report
                if bp_issues:
                    for i, issue in enumerate(bp_issues[:3]):
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
                    
                    # If there are more issues, add a note
                    if len(bp_issues) > 3:
                        add_paragraph(
                            document,
                            f"... and {len(bp_issues) - 3} more issues at this breakpoint."
                        )
                else:
                    add_paragraph(document, "No specific issues documented at this breakpoint.")
        
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
    
    # Add consolidated findings summary
    elements = consolidated_results.get('elements', {})
    if elements:
        add_subheading(document, "Elements with Issues Across Multiple Breakpoints")
        
        # Find elements with issues at multiple breakpoints
        multi_breakpoint_elements = {
            k: v for k, v in elements.items() 
            if len(v.get('breakpoints', [])) > 1
        }
        
        if multi_breakpoint_elements:
            headers = ["Element", "Issue Type", "Breakpoints Affected", "Details"]
            rows = []
            
            for elem_key, elem_data in multi_breakpoint_elements.items():
                element_name = elem_data.get('element', 'Unknown')
                if elem_data.get('id'):
                    element_name += f" (id: {elem_data['id']})"
                
                issue_type = elem_data.get('issueType', 'Unknown issue')
                breakpoints = [str(bp) for bp in elem_data.get('breakpoints', [])]
                details = elem_data.get('details', 'No details provided')
                
                rows.append([
                    element_name,
                    issue_type,
                    ", ".join(breakpoints),
                    details
                ])
            
            if rows:
                table = add_table(document, headers, rows)
            
            add_paragraph(
                document,
                "These elements have issues across multiple breakpoints and should be prioritized for remediation."
            )
        else:
            add_paragraph(
                document,
                "No elements have issues across multiple breakpoints."
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