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
        # Adjust merge based on expected content width, assuming 5 columns for this title
        # A more dynamic approach might be needed if content varies greatly
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + 4)

    def _add_title_to_sheet(self, ws, title: str, row: int = 1):
        """Adds a styled title to a worksheet."""
        cell = ws.cell(row=row, column=1, value=title)
        cell.font = Font(size=16, bold=True, color="000000")
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = Border(bottom=Side(style='thin', color='000000'))

    def _load_excel_data(self, filepath: str) -> Optional[pd.DataFrame]:
        """Loads data from an Excel file, handling potential errors."""
        if not os.path.exists(filepath):
            print(f"Error: Data file not found at '{filepath}'.")
            return None
        try:
            df = pd.read_excel(filepath)
            return df
        except FileNotFoundError:
            print(f"Error: File not found: {filepath}")
            return None
        except Exception as e:
            print(f"Error reading Excel file {filepath}: {e}")
            return None

    def generate_competitor_report(self, ws_name: str = "Competitor Analysis"):
        """Generates a report sheet for competitor analysis."""
        ws = self.wb.create_sheet(ws_name)
        self._add_title_to_sheet(ws, "Analyse de la Concurrence")

        df_competitors = self._load_excel_data(self.competitor_data_file)

        if df_competitors is None or df_competitors.empty:
            ws.cell(row=3, column=1, value="Données de concurrents non trouvées ou vides.")
            return

        # Add data to worksheet
        row_idx = 3 # Start data from row 3 after title and header
        for r_idx, row in enumerate(dataframe_to_rows(df_competitors, index=False)):
            for c_idx, value in enumerate(row):
                cell = ws.cell(row=row_idx + r_idx, column=1 + c_idx, value=value)
                # Basic styling for header row
                if r_idx == 0:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid") # Light gray
                # Adjust column widths for readability
                ws.column_dimensions[get_column_letter(1 + c_idx)].width = max(len(str(value)) + 2, 15) # Minimum width of 15

        # Add a basic chart if possible (e.g., number of strengths)
        if 'Strengths' in df_competitors.columns and 'Competitor' in df_competitors.columns:
            try:
                df_competitors['Num_Strengths'] = df_competitors['Strengths'].str.split(',').str.len().fillna(0)
                
                chart = BarChart()
                chart.title = "Nombre de Points Forts Listés par Concurrent"
                chart.style = 10 # Choose a chart style
                chart.y_axis.title = 'Concurrent'
                chart.x_axis.title = 'Nombre de Points Forts'

                # Data for the chart
                values = Reference(ws, min_col=ws.max_column, min_row=4, max_row=ws.max_row) # Assuming Num_Strengths is the last column added
                categories = Reference(ws, min_col=1, min_row=4, max_row=ws.max_row) # Competitor names

                chart.add_data(values, titles_from_data=False)
                chart.set_categories(categories)
                chart.legend = None # No legend needed for single series

                # Position the chart
                chart_pos_row = ws.max_row + 2
                ws.add_chart(chart, f"A{chart_pos_row}")
                print("Competitor strengths bar chart added to the report.")
            except Exception as e:
                print(f"Could not generate competitor strengths chart: {e}")

    def generate_market_report(self, ws_name: str = "Market Overview"):
        """Generates a report sheet for market overview."""
        ws = self.wb.create_sheet(ws_name)
        self._add_title_to_sheet(ws, "Aperçu du Marché")

        # Option 1: Load from existing market overview Excel
        df_market = self._load_excel_data(self.market_data_file)

        if df_market is None or df_market.empty:
            # Option 2: Use the internal market_info if Excel is not available
            try:
                from ..data.market_data import MarketDataHandler
                handler = MarketDataHandler()
                market_data_dict = handler.get_market_info()
                if market_data_dict:
                    data_for_excel = []
                    for key, value in market_data_dict.items():
                        data_for_excel.append({'Category': key.replace('_', ' ').title(), 'Details': value})
                    df_market = pd.DataFrame(data_for_excel)
                    print("Using internal market data as Excel source not found.")
                else:
                    ws.cell(row=3, column=1, value="Données du marché non trouvées.")
                    return
            except ImportError:
                ws.cell(row=3, column=1, value="Données du marché non trouvées et MarketDataHandler non importable.")
                return
            except Exception as e:
                ws.cell(row=3, column=1, value=f"Erreur lors du chargement des données de marché internes : {e}")
                return

        # Add data to worksheet
        row_idx = 3 # Start data from row 3 after title and header
        for r_idx, row in enumerate(dataframe_to_rows(df_market, index=False)):
            for c_idx, value in enumerate(row):
                cell = ws.cell(row=row_idx + r_idx, column=1 + c_idx, value=value)
                # Basic styling for header row
                if r_idx == 0:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid") # Light gray
                # Adjust column widths for readability
                ws.column_dimensions[get_column_letter(1 + c_idx)].width = max(len(str(value)) + 2, 20) # Increased width for details

        # Add a pie chart for target market distribution if applicable
        if 'Category' in df_market.columns and 'Details' in df_market.columns:
            target_market_row = df_market[df_market['Category'] == 'Target Market']
            if not target_market_row.empty and isinstance(target_market_row.iloc[0]['Details'], list):
                try:
                    target_market_list = target_market_row.iloc[0]['Details']
                    # Create a DataFrame for chart data if needed, or use the list directly
                    chart_data_dict = {
                        'Market Segment': [item.split(':')[0].strip() if ':' in item else item for item in target_market_list],
                        'Value': [1] * len(target_market_list) # Assign arbitrary value for count
                    }
                    df_chart = pd.DataFrame(chart_data_dict)
                    
                    chart = PieChart()
                    chart.title = "Répartition du Marché Cible"
                    
                    # Data for the chart
                    labels = Reference(ws, min_col=ws.max_column - 1, min_row=4, max_row=3 + len(df_chart)) # Target Market Segment column
                    data = Reference(ws, min_col=ws.max_column, min_row=4, max_row=3 + len(df_chart)) # Value column

                    chart.add_data(data, titles_from_data=False)
                    chart.set_categories(labels)
                    chart.legend.position = 'right'
                    chart.legend.font.size = Pt(8)

                    # Position the chart
                    chart_pos_row = ws.max_row + 5
                    ws.add_chart(chart, f"A{chart_pos_row}")
                    print("Target market distribution pie chart added to the report.")
                except Exception as e:
                    print(f"Could not generate target market pie chart: {e}")


    def save_report(self) -> Optional[str]:
        """Saves the entire workbook to the specified output file."""
        if not os.path.exists(self.reports_dir):
            try:
                os.makedirs(self.reports_dir)
                print(f"Created directory: {self.reports_dir}")
            except OSError as e:
                print(f"Error creating directory {self.reports_dir}: {e}")
                return None

        try:
            self.wb.save(self.output_filename)
            print(f"Market study report saved successfully to '{self.output_filename}'.")
            return self.output_filename
        except Exception as e:
            print(f"Error saving report to {self.output_filename}: {e}")
            return None

    def generate_full_report(self) -> Optional[str]:
        """Generates all sections of the Excel report and saves it."""
        self.generate_competitor_report()
        self.generate_market_report()
        return self.save_report()

