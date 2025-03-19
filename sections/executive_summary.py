def add_executive_summary(doc, db_connection, total_domains):
    """Add the executive summary section to the report"""
    h1 = doc.add_heading('Executive Summary', level=1)
    h1.style = doc.styles['Heading 1']
    
    h2 = doc.add_heading('Disclaimer', level=2)
    h2.style = doc.styles['Heading 2']
    
    # Get the page and domain counts directly from the database
    all_test_runs = db_connection.get_all_test_runs()
    page_count = 0
    domain_count = 0
    
    if all_test_runs:
        test_run_ids = [str(run['_id']) for run in all_test_runs]
        unique_urls = db_connection.page_results.distinct('url', {'test_run_id': {'$in': test_run_ids}})
        page_count = len(unique_urls)
        
        # Extract unique domains
        domains = set()
        for url in unique_urls:
            domain = url.replace('http://', '').replace('https://', '').split('/')[0]
            domains.add(domain)
        domain_count = len(domains)    
    
    overview_disclaimer = doc.add_paragraph()
    overview_disclaimer.add_run(f"""
Disclaimer: This accessibility review represents a digital accessibility inspection of {page_count} pages from {domain_count} websites, looking at page properties that help indicate the accessibility health of the page and site. It is not designed to be a comprehensive report on every accessibility issue on each of the pages inspected, merely indicative of potential issues. A formal manual inspection as part of a digital accessibility audit that includes an element of lived experience user testing is required to make any claim on conformance to accessibility standards.
""".strip())

    h2 = doc.add_heading('Subject areas', level=2)
    h2.style = doc.styles['Heading 2']

    overview_gen_1 = doc.add_paragraph()
    overview_gen_1.add_run("""
The report sets out to identify features of websites that impact on web accessibility across a random selection of pages from each site, and to consider what potential issues those pages and sites may have, and to compare this across sites to see which sites are at most risk of failing accessibility.
""".strip())

    overview_gen_2 = doc.add_paragraph()
    overview_gen_2.add_run("""
A number of subjects that impact on web accessibility were considered:
""".strip())
    
    doc.add_paragraph("Basic HTML structure (page language, title etc.)", style='List Number')
    doc.add_paragraph("Content styling (fonts, lists, animation, title attribute)", style='List Number')
    doc.add_paragraph("Multimedia content (images, videos, maps)", style='List Number')
    doc.add_paragraph("Non-text content (adjacent blocks)", style='List Number')
    doc.add_paragraph("Underlying semantic structure (headings, landmarks, tables, timers)", style='List Number')  
    doc.add_paragraph("Navigation aids (menus, forms, dialogs, links, buttons, 'more' controls, floating dialogs), ", style='List Number')
    doc.add_paragraph("Navigation order (tab order, tabindex), ", style='List Number')
    doc.add_paragraph("Accessibility supported (event handling, focus management, floating content, accessible names) ", style='List Number')
    doc.add_paragraph("Use of Colour (text/non-text contrast,user as indicator) ", style='List Number')
    doc.add_paragraph("Local electronic documents (Linked PDFs)", style='List Number')
    