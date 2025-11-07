"""
Main report generation module
"""

import os
from .excel_generator import MarketStudyExcelReport
from .powerpoint_generator import MarketStudyPresentation

def ensure_directories():
    """Ensure required directories exist."""
    directories = ['data', 'reports']
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"Created directory: {directory}")

def generate_all_reports():
    """Generate all market study reports."""
    print("Generating market study reports...")
    
    # Ensure directories exist
    ensure_directories()
    
    # Define file paths
    competitor_data_file = 'data/competitor_research.xlsx'
    market_data_file = 'data/market_overview.xlsx'
    
    # Check if input data files exist, if not, try to create them
    if not os.path.exists(competitor_data_file):
        print(f"Warning: {competitor_data_file} not found. Attempting to create template.")
        from ..data.competitor_data import CompetitorDataCollector
        collector = CompetitorDataCollector()
        collector.create_competitor_template()
        if not os.path.exists(competitor_data_file):
            print(f"Error: Failed to create {competitor_data_file}. Cannot generate reports.")
            return None, None

    if not os.path.exists(market_data_file):
        print(f"Warning: {market_data_file} not found. Attempting to create market data excel.")
        from ..data.market_data import MarketDataHandler
        handler = MarketDataHandler()
        handler.create_market_summary_excel()
        if not os.path.exists(market_data_file):
            print(f"Error: Failed to create {market_data_file}. Cannot generate reports.")
            return None, None

    # Generate Excel report
    print("  - Generating Excel report...")
    excel_report = MarketStudyExcelReport(
        competitor_data_file=competitor_data_file,
        market_data_file=market_data_file
    )
    excel_filename = excel_report.save_workbook()
    if excel_filename:
        print(f"  ✓ Excel report generated: {excel_filename}")
    else:
        print("  ✗ Excel report generation failed.")
    
    # Generate PowerPoint presentation
    print("  - Generating PowerPoint presentation...")
    ppt_report = MarketStudyPresentation(
        competitor_data_file=competitor_data_file,
        market_data_file=market_data_file
    )
    ppt_filename = ppt_report.save_presentation()
    if ppt_filename:
        print(f"  ✓ PowerPoint presentation generated: {ppt_filename}")
    else:
        print("  ✗ PowerPoint presentation generation failed.")
    
    return excel_filename, ppt_filename

if __name__ == "__main__":
    # Ensure data files are available before generating reports
    # This part is crucial for the main execution flow of generate_reports itself
    if not os.path.exists('data'):
        os.makedirs('data')
        
    if not os.path.exists('data/competitor_research.xlsx'):
        from ..data.competitor_data import CompetitorDataCollector
        collector = CompetitorDataCollector()
        collector.create_competitor_template()
        print("Created dummy competitor_research.xlsx for testing.")
        
    if not os.path.exists('data/market_overview.xlsx'):
        from ..data.market_data import MarketDataHandler
        handler = MarketDataHandler()
        handler.create_market_summary_excel()
        print("Created dummy market_overview.xlsx for testing.")

    excel_file, ppt_file = generate_all_reports()
    
    if excel_file and ppt_file:
        print("\nAll reports generated successfully!")
        print(f"  - Excel: {excel_file}")
        print(f"  - PowerPoint: {ppt_file}")
    else:
        print("\nReport generation process encountered errors.")
