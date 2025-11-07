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
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + 4)

    def _add_dataframe_to_sheet(self, ws, df: pd.DataFrame, title: str, start_row: int = 1, start_col: int = 1) -> None:
        """Adds a DataFrame to a worksheet with a title and basic styling."""
        ws.cell(row=start_row, column=start_col, value=title).font = Font(size=14, bold=True)
        
        # Write DataFrame rows
        for r_idx, row_data in enumerate(dataframe_to_rows(df, index=False, header=True), start=start_row + 1):
            for c_idx, value in enumerate(row_data, start=start_col):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                # Apply header style
                if r_idx == start_row + 1:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid") # Light gray
                    cell.border = Border(bottom=Side(style='thin'))

        # Auto-adjust column widths
        for col in ws.columns:
            max_length = 0
            column_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width

    def generate_report(self) -> Optional[str]:
        """
        Generates the comprehensive Excel market study report.

        Returns:
            Optional[str]: The path to the generated Excel report file, or None if generation failed.
        """
        try:
            # --- Sheet 1: Market Overview ---
            ws_market = self.wb.create_sheet("Market Overview")
            self._style_header(ws_market, "Vue d'Ensemble du Marché")
            
            market_data_handler = MarketDataHandler()
            market_info = market_data_handler.load_market_data()

            if market_info:
                # Convert market_info dictionary to DataFrame for easier handling
                market_df_dict = {}
                for key, value in market_info.items():
                    if isinstance(value, list):
                        market_df_dict[key.replace('_', ' ').title()] = ["; ".join(value)]
                    else:
                        market_df_dict[key.replace('_', ' ').title()] = [value]
                market_df = pd.DataFrame([market_df_dict])
                self._add_dataframe_to_sheet(ws_market, market_df, "Informations Clés du Marché", start_row=2, start_col=1)
            else:
                ws_market.cell(row=3, column=1, value="Données du marché non disponibles.").font = Font(italic=True, color="FF0000")

            # --- Sheet 2: Competitor Analysis ---
            ws_competitor = self.wb.create_sheet("Competitor Analysis")
            self._style_header(ws_competitor, "Analyse de la Concurrence")

            competitor_df = None
            if os.path.exists(self.competitor_data_file):
                try:
                    competitor_df = pd.read_excel(self.competitor_data_file)
                    competitor_df.columns = competitor_df.columns.str.strip() # Clean column names
                    self._add_dataframe_to_sheet(ws_competitor, competitor_df, "Données des Concurrents", start_row=2, start_col=1)
                except Exception as e:
                    print(f"Error reading competitor data file {self.competitor_data_file}: {e}")
                    ws_competitor.cell(row=3, column=1, value=f"Erreur lors de la lecture de {self.competitor_data_file}: {e}").font = Font(italic=True, color="FF0000")
            else:
                ws_competitor.cell(row=3, column=1, value="Fichier de données concurrentielles non trouvé. Veuillez exécuter la collecte de données.").font = Font(italic=True, color="FF0000")

            # --- Add Charts (if data is available) ---
            # Competitor Strength Comparison Chart (if generated by analysis)
            competitor_chart_path = 'reports/competitor_strength_comparison.png'
            if os.path.exists(competitor_chart_path):
                try:
                    img = openpyxl.drawing.image.Image(competitor_chart_path)
                    # Adjust image size and position as needed
                    img.height = 300 
                    img.width = 500
                    ws_competitor.add_image(img, 'F2') # Position the image
                except Exception as e:
                    print(f"Could not add competitor chart image: {e}")
                    ws_competitor.cell(row=len(competitor_df) + 5 if competitor_df is not None else 5, column=1, value="Erreur lors de l'ajout du graphique de comparaison des concurrents.").font = Font(italic=True, color="FF0000")
            
            # --- Save the workbook ---
            self.wb.save(self.output_filename)
            print(f"Market study report generated successfully at: {self.output_filename}")
            return self.output_filename

        except Exception as e:
            print(f"An error occurred during report generation: {e}")
            return None

