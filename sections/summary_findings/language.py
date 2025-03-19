def add_language_section(doc, db_connection, total_domains):
    """Add the Language of Page section to the summary findings"""
    h2 = doc.add_heading('Language of Page', level=2)
    h2.style = doc.styles['Heading 2']

    # If there are pages without lang attribute, list them
    pages_without_lang = list(db_connection.page_results.find(
        {"results.accessibility.tests.html_structure.html_structure.tests.hasValidLang": False},
        {"url": 1, "_id": 0}
    ).sort("url", 1))

    # Count affected domains
    affected_domains = set()
    for page in pages_without_lang:
        domain = page['url'].replace('http://', '').replace('https://', '').split('/')[0]
        affected_domains.add(domain)

    # Calculate percentage
    percentage = (len(affected_domains) / len(total_domains)) * 100 if total_domains else 0

    if pages_without_lang:
        doc.add_paragraph(f"Found {len(pages_without_lang)} pages ({percentage:.1f}% of sites) without valid language attribute.")
        doc.add_paragraph("A properly defined language attribute is crucial for screen readers to use the correct pronunciation rules.")
    else:
        doc.add_paragraph("All pages have a valid lang attribute.")
        