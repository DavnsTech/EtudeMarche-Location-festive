"""
Excel report generator for market study data
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows
import os

class MarketStudyExcelReport:
    def __init__(self, 
                 competitor_data_file='data/competitor_research.xlsx',
                 market_data_file='data/market_overview.xlsx'):
        self.wb = Workbook()
        self.competitor_data_file = competitor_data_file
        self.market_data_file = market_data_file
        self.reports_dir = 'reports'
        self.output_filename = os.path.join(self.reports_dir, 'Location_Festive_Niort_Market_Study_Report.xlsx')

        # Ensure reports directory exists
        if not os.path.exists(self.reports_dir):
            os.makedirs(self.reports_dir)

    def _style_header(self, ws, title, start_row=1, start_col=1):
        """Helper to style a title cell."""
        cell = ws.cell(row=start_row, column=start_col, value=title)
        cell.font = Font(size=16, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + 3) # Assuming title spans 4 columns

    def _style_section_title(self, ws, title, row, col=1):
        """Helper to style section titles."""
        cell = ws.cell(row=row, column=col, value=title)
        cell.font = Font(size=14, bold=True)
        cell.alignment = Alignment(horizontal="left")
        return row + 1 # Return next row

    def create_executive_summary_sheet(self):
        """Create executive summary sheet."""
        ws = self.wb.active
        ws.title = "Résumé Exécutif"
        
        # Add title
        self._style_header(ws, "ÉTUDE DE MARCHE - LOCATION FESTIVE NIORT", 1, 1)
        
        # Add summary content
        content = [
            "Cette étude de marché analyse le secteur de la location de matériel festif à Niort et ses environs.",
            "L'analyse couvre les principaux concurrents, les tendances du marché et les opportunités pour Location Festive Niort.",
            "Les principales conclusions de cette étude sont présentées ci-dessous :",
            "",
            "1. Marché concurrentiel bien établi avec plusieurs acteurs locaux",
            "2. Opportunités de différenciation par l'approvisionnement en Chine",
            "3. Collaboration potentielle avec les établissements scolaires (APE)",
            "4. Importance d'une présence numérique forte (réseaux sociaux, site web)"
        ]
        
        row = 3
        for line in content:
            ws.cell(row=row, column=1, value=line)
            row += 1

    def create_competitor_analysis_sheet(self):
        """Create competitor analysis sheet."""
        ws = self.wb.create_sheet("Analyse Concurrentielle")
        
        # Add title
        self._style_header(ws, "ANALYSE DES CONCURRENTS", 1, 1)
        
        # Load competitor data if available
        if os.path.exists(self.competitor_data_file):
            try:
                df = pd.read_excel(self.competitor_data_file)
                # Add data to worksheet
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
                
                # Style header row
                for cell in ws[2]:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            except Exception as e:
                print(f"Error loading competitor data: {e}")
                ws.cell(row=3, column=1, value="Données concurrentielles non disponibles")
        else:
            ws.cell(row=3, column=1, value="Fichier de données concurrentielles non trouvé")

    def create_market_overview_sheet(self):
        """Create market overview sheet."""
        ws = self.wb.create_sheet("Aperçu du Marché")
        
        # Add title
        self._style_header(ws, "APERÇU DU MARCHÉ", 1, 1)
        
        # Load market data if available
        if os.path.exists(self.market_data_file):
            try:
                df = pd.read_excel(self.market_data_file)
                # Add data to worksheet
                for r in dataframe_to_rows(df, index=False, header=True):
                    ws.append(r)
                
                # Style header row
                for cell in ws[2]:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            except Exception as e:
                print(f"Error loading market data: {e}")
                ws.cell(row=3, column=1, value="Données de marché non disponibles")
        else:
            ws.cell(row=3, column=1, value="Fichier de données de marché non trouvé")

    def create_charts_sheet(self):
        """Create charts sheet with market visualization."""
        ws = self.wb.create_sheet("Visualisations")
        
        # Add title
        self._style_header(ws, "VISUALISATIONS DU MARCHÉ", 1, 1)
        
        # Add placeholder for chart
        ws.cell(row=3, column=1, value="Graphiques d'analyse du marché")
        ws.cell(row=4, column=1, value="Les graphiques seront générés après la collecte des données")

    def generate_report(self):
        """Generate the complete Excel report."""
        # Create all sheets
        self.create_executive_summary_sheet()
        self.create_competitor_analysis_sheet()
        self.create_market_overview_sheet()
        self.create_charts_sheet()
        
        # Save workbook
        try:
            self.wb.save(self.output_filename)
            print(f"Excel report generated successfully: {self.output_filename}")
            return self.output_filename
        except Exception as e:
            print(f"Error generating Excel report: {e}")
            return None
