import click
from datetime import datetime
import sys
import os

# Add the test_with_mongo directory to the path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../test_with_mongo')))
from database import AccessibilityDB
from report_generator import generate_report

@click.command()
@click.option('--title', '-t', 
              default='Accessibility Test Report', 
              help='Title of the report')
@click.option('--author', '-a', 
              required=True, 
              help='Author of the report')
@click.option('--output_folder', '-o', 
              default='reports', 
              help='Folder where the report will be filed')
@click.option('--date', '-d', 
              default=datetime.now().strftime("%Y-%m-%d"),
              help='Date of the report (YYYY-MM-DD)')
@click.option('--database', '-db',
              default=None,
              help='MongoDB database name to use (default: accessibility_tests)')
def main(title, author, date, output_folder, database):
    """Generate an accessibility test report with specified parameters."""
    try:
        datetime.strptime(date, "%Y-%m-%d")
    except ValueError:
        click.echo("Error: Date must be in YYYY-MM-DD format")
        return

    click.echo(f"Generating report with the following parameters:")
    click.echo(f"Title: {title}")
    click.echo(f"Author: {author}")
    click.echo(f"Date: {date}")
    click.echo(f"Output folder: {output_folder}")
    if database:
        click.echo(f"Database: {database}")
    
    db = AccessibilityDB(db_name=database)
    report_file = generate_report(db, title, author, date, output_folder)
    
    if report_file:
        click.echo(f"\nReport generated successfully: {report_file}")
        click.echo("\nIMPORTANT: To complete the report formatting:")
        click.echo("1. Open the document in Microsoft Word")
        click.echo("2. Right-click anywhere in the table of contents")
        click.echo("3. Select 'Update Field'")
        click.echo("4. Choose 'Update entire table'")
    else:
        click.echo("Failed to generate report", err=True)

if __name__ == "__main__":
    main()