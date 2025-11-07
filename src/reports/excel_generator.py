"""
Excel report generator for market study data
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import os
from typing import Optional, Dict, Any

class MarketStudyExcelReport:
    def __init__(self, 
                 competitor_data_file: str = 'data/competitor_research.xlsx',
                 market_data_file: str = 'data/market_overview.xlsx'):
        """
        Initializes the Excel report generator.

        Args:
            competitor_data_file (str): Path to the competitor research Excel file.
            market_data_file (str): Path to the market overview Excel file.
        """
        self.wb = Workbook()
        self.competitor_data_file = competitor_data_file
        self.market_data_file = market_data_file
        self.reports_dir: str = 'reports'
        self.output_filename: str = os.path.join(self.reports_dir, 'Location_Festive_Niort_Market_Study_Report.xlsx')

        # Ensure reports directory exists
        if not os.path.exists(self.reports_dir):
            try:
                os.makedirs(self.reports_dir)
                print(f"Created directory: {self.reports_dir}")
            except OSError as e:
                print(f"Error creating directory {self.reports_dir}: {e}")

    def _style_header(self, ws, title: str, start_row: int = 1, start_col: int = 1) -> None:
        """Helper to style a title cell."""
        cell = ws.cell(row=start_row, column=start_col, value=title)
        cell.font = Font(size=18, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        # Merge cells for header (assuming 5 columns width)
        ws.merge_cells(start_row=start_row, start_column=start_col, 
                       end_row=start_row, end_column=start_col+4)

    def _add_dataframe_to_sheet(self, ws, df: pd.DataFrame, start_row: int = 2) -> None:
        """Adds a DataFrame to a worksheet with styling."""
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start_row):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == start_row:  # Header row
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                
                # Add border to all cells
                thin_border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                cell.border = thin_border

        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)  # Cap at 50 characters
            ws.column_dimensions[column_letter].width = adjusted_width

    def generate_report(self) -> Optional[str]:
        """
        Generates the complete Excel market study report.

        Returns:
            Optional[str]: Path to the generated report or None if failed.
        """
        try:
            # Remove default sheet
            default_sheet = self.wb.active
            self.wb.remove(default_sheet)

            # Create Competitor Analysis sheet
            competitor_sheet = self.wb.create_sheet("Competitor Analysis")
            self._style_header(competitor_sheet, "Competitor Analysis", 1, 1)
            
            # Load competitor data
            if os.path.exists(self.competitor_data_file):
                competitor_df = pd.read_excel(self.competitor_data_file)
                self._add_dataframe_to_sheet(competitor_sheet, competitor_df, 3)
            else:
                competitor_sheet.cell(row=3, column=1, value="Competitor data not available")

            # Create Market Overview sheet
            market_sheet = self.wb.create_sheet("Market Overview")
            self._style_header(market_sheet, "Market Overview", 1, 1)
            
            # Load market data
            if os.path.exists(self.market_data_file):
                market_df = pd.read_excel(self.market_data_file)
                self._add_dataframe_to_sheet(market_sheet, market_df, 3)
            else:
                market_sheet.cell(row=3, column=1, value="Market data not available")

            # Save workbook
            self.wb.save(self.output_filename)
            print(f"Excel report saved to: {self.output_filename}")
            return self.output_filename

        except Exception as e:
            print(f"Error generating Excel report: {e}")
            return None
