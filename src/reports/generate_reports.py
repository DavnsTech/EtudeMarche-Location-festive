"""
Main report generation module
"""

import os
from .excel_generator import MarketStudyExcelReport
from .powerpoint_generator import MarketStudyPresentation
from typing import Tuple, Optional

def ensure_directories() -> None:
    """Ensures required directories for data and reports exist."""
    directories = ['data', 'reports']
    for directory in directories:
        if not os.path.exists(directory):
            try:
                os.makedirs(directory)
                print(f"Created directory: {directory}")
            except OSError as e:
                print(f"Error creating directory {directory}: {e}")

def generate_all_reports() -> Tuple[Optional[str], Optional[str]]:
    """
    Generates all market study reports (Excel and PowerPoint).

    Returns:
        Tuple[Optional[str], Optional[str]]: Paths to the generated Excel and PowerPoint reports, or None if generation failed.
    """
    print("Generating market study reports...")
    
    # Ensure directories exist
    ensure_directories()
    
    # Define file paths
    competitor_data_file = 'data/competitor_research.xlsx'
    market_data_file = 'data/market_overview.xlsx'
    
    # Check if input data files exist, if not, try to create them
    if not os.path.exists(competitor_data_file):
        print(f"Warning: {competitor_data_file} not found. Attempting to create template.")
        try:
            from ..data.competitor_data import CompetitorDataCollector
            collector = CompetitorDataCollector()
            collector.create_competitor_template()
            if not os.path.exists(competitor_data_file):
                print(f"Error: Failed to create {competitor_data_file}. Cannot generate reports.")
                return None, None
        except ImportError:
            print("Error: Could not import CompetitorDataCollector. Ensure the data module is correctly structured.")
            return None, None
        except Exception as e:
            print(f"An unexpected error occurred while creating competitor template: {e}")
            return None, None

    if not os.path.exists(market_data_file):
        print(f"Warning: {market_data_file} not found. Attempting to create market overview.")
        try:
            from ..data.market_data import MarketDataHandler
            handler = MarketDataHandler()
            handler.create_market_summary_excel()
            if not os.path.exists(market_data_file):
                print(f"Error: Failed to create {market_data_file}. Cannot generate reports.")
                return None, None
        except ImportError:
            print("Error: Could not import MarketDataHandler. Ensure the data module is correctly structured.")
            return None, None
        except Exception as e:
            print(f"An unexpected error occurred while creating market overview: {e}")
            return None, None

    # Generate Excel report
    print("  - Generating Excel report...")
    excel_report = MarketStudyExcelReport(competitor_data_file, market_data_file)
    excel_path = excel_report.generate_report()
    
    # Generate PowerPoint presentation
    print("  - Generating PowerPoint presentation...")
    ppt_report = MarketStudyPresentation(competitor_data_file, market_data_file)
    ppt_path = ppt_report.generate_presentation()
    
    if excel_path and ppt_path:
        print("✓ All reports generated successfully.")
        return excel_path, ppt_path
    else:
        print("✗ Failed to generate one or more reports.")
        return None, None
