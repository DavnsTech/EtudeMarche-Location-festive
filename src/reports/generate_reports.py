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
    excel_filename = excel_report.generate_report()
    
    # Generate PowerPoint presentation
    print("  - Generating PowerPoint presentation...")
    ppt_report = MarketStudyPresentation(
        competitor_data_file=competitor_data_file,
        market_data_file=market_data_file
    )
    ppt_filename = ppt_report.create_presentation()
    
    return excel_filename, ppt_filename
