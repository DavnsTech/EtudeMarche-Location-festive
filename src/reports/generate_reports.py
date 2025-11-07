"""
Main report generation module
"""

import os
from .excel_generator import MarketStudyExcelReport
from .powerpoint_generator import MarketStudyPresentation

def ensure_directories():
    """Ensure required directories exist"""
    directories = ['data', 'reports', 'src']
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)

def generate_all_reports():
    """Generate all market study reports"""
    print("Generating market study reports...")
    
    # Ensure directories exist
    ensure_directories()
    
    # Generate Excel report
    excel_report = MarketStudyExcelReport()
    excel_filename = excel_report.save_workbook()
    print(f"Excel report generated: {excel_filename}")
    
    # Generate PowerPoint presentation
    ppt_report = MarketStudyPresentation()
    ppt_filename = ppt_report.save_presentation()
    print(f"PowerPoint presentation generated: {ppt_filename}")
    
    return excel_filename, ppt_filename

if __name__ == "__main__":
    generate_all_reports()
    print("All reports generated successfully!")
