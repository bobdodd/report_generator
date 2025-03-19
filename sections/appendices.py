from report_styling import format_table_text

def add_test_coverage_appendix(doc, db_connection):
    """Add the Test Coverage appendix section"""
    h2 = doc.add_heading('Test Coverage', level=2)
    h2.style = doc.styles['Heading 2']

    # Add explanation
    doc.add_paragraph("""
    This section provides an overview of the sites and pages included in the accessibility analysis. Understanding the scope
    of testing is important for evaluating the comprehensiveness of the assessment and identifying areas that may need
    additional coverage.
    """.strip())

    # Add notes about coverage
    doc.add_paragraph()
    doc.add_paragraph("Notes about Coverage:", style='Normal')
    doc.add_paragraph("The pages tested represent a sample of each site's content", style='List Bullet')
    doc.add_paragraph("Testing includes various page types (home pages, content pages, forms, etc.)", style='List Bullet')
    doc.add_paragraph("Coverage may vary by site based on site structure and complexity", style='List Bullet')

    doc.add_paragraph()

    # Query for all tested pages
    tested_pages = list(db_connection.page_results.find(
        {},
        {
            "url": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Process the pages
    sites_data = {}
    for page in tested_pages:
        url = page['url']
        domain = url.replace('http://', '').replace('https://', '').split('/')[0]
        
        if domain not in sites_data:
            sites_data[domain] = {
                'pages': set(),
                'count': 0
            }
        
        sites_data[domain]['pages'].add(url)
        sites_data[domain]['count'] += 1

    # Create summary statistics
    total_sites = len(sites_data)
    total_pages = sum(site['count'] for site in sites_data.values())
    avg_pages_per_site = total_pages / total_sites if total_sites > 0 else 0

    # Add summary statistics
    doc.add_paragraph("Summary Statistics:", style='Normal')
    summary_table = doc.add_table(rows=3, cols=2)
    summary_table.style = 'Table Grid'

    # Add summary data
    rows = summary_table.rows
    rows[0].cells[0].text = "Total Sites Tested"
    rows[0].cells[1].text = str(total_sites)
    rows[1].cells[0].text = "Total Pages Tested"
    rows[1].cells[1].text = str(total_pages)
    rows[2].cells[0].text = "Average Pages per Site"
    rows[2].cells[1].text = f"{avg_pages_per_site:.1f}"

    format_table_text(summary_table)

    # Add site-by-site breakdown
    doc.add_paragraph()
    doc.add_paragraph("Coverage by Site:", style='Normal')

    # Create sites overview table
    sites_table = doc.add_table(rows=len(sites_data) + 1, cols=2)
    sites_table.style = 'Table Grid'

    # Add headers
    headers = sites_table.rows[0].cells
    headers[0].text = "Site"
    headers[1].text = "Pages Tested"

    # Add site data
    for i, (domain, data) in enumerate(sorted(sites_data.items()), 1):
        row = sites_table.rows[i].cells
        row[0].text = domain
        row[1].text = str(data['count'])

    format_table_text(sites_table)

def add_documents_appendix(doc, db_connection):
    """Add the Electronic Documents appendix section"""
    doc.add_page_break()
    h2 = doc.add_heading('Electronic documents found', level=2)
    h2.style = doc.styles['Heading 2']

    # Add explanation
    doc.add_paragraph("""
    This section lists the electronic documents found across all tested pages.
    """.strip())

    doc.add_paragraph()

    # Query for pages with document information
    pages_with_documents = list(db_connection.page_results.find(
        {
            "results.accessibility.tests.documents.document_links": {"$exists": True}
        },
        {
            "url": 1,
            "results.accessibility.tests.documents.document_links": 1,
            "_id": 0
        }
    ).sort("url", 1))

    # Create a list of all documents
    all_documents = []
    for page in pages_with_documents:
        try:
            doc_links = page.get('results', {}).get('accessibility', {}).get('tests', {}).get('documents', {}).get('document_links', {})
            if 'documents' in doc_links:
                for document in doc_links['documents']:
                    all_documents.append({
                        'page_url': page['url'],
                        'doc_url': document.get('url'),
                        'type': document.get('type')
                    })
        except Exception as e:
            print(f"Error processing page {page.get('url')}: {str(e)}")
            continue

    # Count documents by type
    type_counts = {}
    for document in all_documents:
        doc_type = document.get('type', 'unknown').upper()
        type_counts[doc_type] = type_counts.get(doc_type, 0) + 1

    # Create summary table
    summary_table = doc.add_table(rows=len(type_counts) + 1, cols=2)
    summary_table.style = 'Table Grid'

    # Add summary headers
    headers = summary_table.rows[0].cells
    headers[0].text = "Document Type"
    headers[1].text = "Count"

    # Add summary data
    for i, (doc_type, count) in enumerate(sorted(type_counts.items()), 1):
        row = summary_table.rows[i].cells
        row[0].text = doc_type
        row[1].text = str(count)

    format_table_text(summary_table)

    doc.add_paragraph()
    doc.add_paragraph("Document Listing:", style='Normal')

    # Create document listing table
    table = doc.add_table(rows=len(all_documents) + 1, cols=3)
    table.style = 'Table Grid'

    # Add headers
    headers = table.rows[0].cells
    headers[0].text = "Type"
    headers[1].text = "Document URL"
    headers[2].text = "Found On Page"

    # Add documents
    for i, document in enumerate(sorted(all_documents, key=lambda x: x['type']), 1):
        row = table.rows[i].cells
        row[0].text = document.get('type', 'unknown').upper()
        row[1].text = document.get('doc_url', 'No URL')
        row[2].text = document.get('page_url', 'Unknown page')

    format_table_text(table)

    # Add total count
    doc.add_paragraph()
    doc.add_paragraph(f"Total Documents Found: {len(all_documents)}")

def add_appendices(doc, db_connection):
    """Add all appendix sections to the report"""
    h1 = doc.add_heading('APPENDICES', level=1)
    h1.style = doc.styles['Heading 1']

    doc.add_paragraph()
    add_test_coverage_appendix(doc, db_connection)
    add_documents_appendix(doc, db_connection)
    