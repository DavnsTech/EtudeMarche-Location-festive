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
        ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col + 1) # Assuming title spans 2 columns

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
        self._style_header(ws, "ÉTUDE DE MARCHE - LOCATION FESTIVE NIORT", start_row=1, start_col=1)
        
        current_row = 3 # Start after title and some spacing

        # Add key metrics placeholder
        ws.cell(row=current_row, column=1, value="Principaux Indicateurs").font = Font(size=12, bold=True)
        current_row += 1
        
        data_metrics = [
            ["Taille du marché estimée", "À déterminer"],
            ["Nombre de concurrents identifiés", "12"], # This should ideally be dynamic
            ["Positionnement concurrentiel envisagé", "Fournisseur unique avec approvisionnement optimisé"],
            ["Croissance annuelle prévue", "À déterminer"],
            ["Investissement initial requis", "À déterminer"]
        ]
        
        for row_data in data_metrics:
            ws.cell(row=current_row, column=1, value=row_data[0]).font = Font(bold=True)
            ws.cell(row=current_row, column=2, value=row_data[1])
            current_row += 1
            
        # Add some spacing
        current_row += 1
        
        # Add key insights
        ws.cell(row=current_row, column=1, value="Points Clés et Opportunités").font = Font(size=12, bold=True)
        current_row += 1
        
        insights = [
            "Opportunité distincte via l'approvisionnement optimisé.",
            "Relations existantes avec les APE scolaires à développer.",
            "Marché local sous-exploité pour la location de matériel festif.",
            "Potentiel de croissance dans les événements privés et d'entreprise."
        ]
        
        for insight in insights:
            ws.cell(row=current_row, column=1, value=f"• {insight}")
            current_row += 1
        
        # Adjust column widths
        for col in ['A', 'B']:
            ws.column_dimensions[col].width = 40

    def create_market_overview_sheet(self):
        """Create market overview sheet from market data."""
        ws = self.wb.create_sheet("Aperçu Marché")
        
        self._style_header(ws, "APERÇU DU MARCHÉ", start_row=1, start_col=1)
        current_row = 3

        try:
            market_df = pd.read_excel(self.market_data_file)
            
            # Use the DataFrame directly to populate the sheet
            # Add DataFrame rows to the worksheet
            for r_idx, row in enumerate(dataframe_to_rows(market_df, index=False, header=True), current_row):
                for c_idx, value in enumerate(row, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=value)
                    if c_idx == 1: # Style first column (Category)
                        cell.font = Font(bold=True)
                    if r_idx == current_row: # Style header row
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="8FAADC", end_color="8FAADC", fill_type="solid")
                        cell.alignment = Alignment(horizontal="center", vertical="center")

            current_row = ws.max_row + 2 # Move past the data

        except FileNotFoundError:
            ws.cell(row=current_row, column=1, value=f"Error: Market data file not found at {self.market_data_file}").font = Font(color="FF0000")
            current_row += 1
        except Exception as e:
            ws.cell(row=current_row, column=1, value=f"Error loading market data: {e}").font = Font(color="FF0000")
            current_row += 1

        # Add a placeholder for market size estimation if not already covered
        ws.cell(row=current_row, column=1, value="Estimation de la Taille du Marché").font = Font(size=12, bold=True)
        current_row += 1
        ws.cell(row=current_row, column=1, value="À déterminer via des recherches supplémentaires (statistiques locales, démographie, etc.)")
        current_row += 1

        # Adjust column widths
        for col in ['A', 'B']:
            ws.column_dimensions[col].width = 45


    def create_competitor_analysis_sheet(self):
        """Create competitor analysis sheet from competitor data."""
        ws = self.wb.create_sheet("Analyse Concurrentielle")
        
        self._style_header(ws, "ANALYSE CONCURRENTIELLE", start_row=1, start_col=1)
        current_row = 3

        try:
            competitor_df = pd.read_excel(self.competitor_data_file)
            
            # Basic analysis metrics
            if not competitor_df.empty:
                total_competitors = len(competitor_df)
                avg_strengths = competitor_df['Strengths'].dropna().apply(lambda x: len(str(x).split(',')) if pd.notnull(x) and str(x).strip() else 0).mean()
                avg_weaknesses = competitor_df['Weaknesses'].dropna().apply(lambda x: len(str(x).split(',')) if pd.notnull(x) and str(x).strip() else 0).mean()
                
                ws.cell(row=current_row, column=1, value="Résumé des Concurrents").font = Font(size=12, bold=True)
                current_row += 1
                ws.cell(row=current_row, column=1, value="Nombre Total de Concurrents Identifiés:").font = Font(bold=True)
                ws.cell(row=current_row, column=2, value=total_competitors)
                current_row += 1
                ws.cell(row=current_row, column=1, value="Nombre Moyen de Points Forts par Concurrent:").font = Font(bold=True)
                ws.cell(row=current_row, column=2, value=f"{avg_strengths:.2f}" if pd.notna(avg_strengths) else "N/A")
                current_row += 1
                ws.cell(row=current_row, column=1, value="Nombre Moyen de Points Faibles par Concurrent:").font = Font(bold=True)
                ws.cell(row=current_row, column=2, value=f"{avg_weaknesses:.2f}" if pd.notna(avg_weaknesses) else "N/A")
                current_row += 2

                # Market Position Analysis
                market_pos_counts = competitor_df['Market Position'].value_counts()
                if not market_pos_counts.empty:
                    ws.cell(row=current_row, column=1, value="Répartition par Position sur le Marché").font = Font(size=12, bold=True)
                    current_row += 1
                    
                    market_pos_df = market_pos_counts.reset_index()
                    market_pos_df.columns = ['Position', 'Count']
                    
                    # Add DataFrame rows
                    for r_idx, row in enumerate(dataframe_to_rows(market_pos_df, index=False, header=True), current_row):
                        for c_idx, value in enumerate(row, 1):
                            cell = ws.cell(row=r_idx, column=c_idx, value=value)
                            if c_idx == 1: # Style first column (Position)
                                cell.font = Font(bold=True)
                            if r_idx == current_row: # Style header row
                                cell.font = Font(bold=True, color="FFFFFF")
                                cell.fill = PatternFill(start_color="8FAADC", end_color="8FAADC", fill_type="solid")
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                    current_row = ws.max_row + 2

                # Detailed Competitor Table
                ws.cell(row=current_row, column=1, value="Détails des Concurrents").font = Font(size=12, bold=True)
                current_row += 1
                
                # Add DataFrame rows for detailed competitor info
                for r_idx, row in enumerate(dataframe_to_rows(competitor_df, index=False, header=True), current_row):
                    for c_idx, value in enumerate(row, 1):
                        cell = ws.cell(row=r_idx, column=c_idx, value=value)
                        if r_idx == current_row: # Style header row
                            cell.font = Font(bold=True, color="FFFFFF")
                            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                current_row = ws.max_row + 1

            else:
                ws.cell(row=current_row, column=1, value="No competitor data found in the specified file.").font = Font(color="FF0000")
                current_row += 1

        except FileNotFoundError:
            ws.cell(row=current_row, column=1, value=f"Error: Competitor data file not found at {self.competitor_data_file}").font = Font(color="FF0000")
            current_row += 1
        except Exception as e:
            ws.cell(row=current_row, column=1, value=f"Error loading competitor data: {e}").font = Font(color="FF0000")
            current_row += 1
        
        # Adjust column widths
        for i, col_name in enumerate(competitor_df.columns if not competitor_df.empty else ['A', 'B', 'C']): # Default if df empty
            try:
                ws.column_dimensions[chr(65 + i)].width = 25 # Basic width, adjust as needed
            except IndexError:
                pass # Handle cases where there are many columns

    def save_workbook(self):
        """Save the workbook to a file."""
        try:
            self.wb.save(self.output_filename)
            print(f"Excel report saved to: {self.output_filename}")
            return self.output_filename
        except Exception as e:
            print(f"Error saving Excel workbook: {e}")
            return None

if __name__ == "__main__":
    # Ensure dummy data files exist for testing
    if not os.path.exists('data/competitor_research.xlsx'):
        from ..data.competitor_data import CompetitorDataCollector
        collector = CompetitorDataCollector()
        collector.create_competitor_template()
        # Add dummy data for testing
        try:
            df_test = pd.read_excel('data/competitor_research.xlsx')
            df_test.loc[0, ['Services', 'Strengths', 'Weaknesses', 'Market Position']] = "Rental,Delivery,Setup,Teardown,Decoration,Lighting,Sound,DJ Equipment", "Wide range of services,Good pricing", "Limited online presence", "Market Leader"
            df_test.loc[1, ['Services', 'Strengths', 'Weaknesses', 'Market Position']] = "Rental,Delivery,Setup,Sound", "Excellent customer service,Fast delivery", "High prices", "Challenger"
            df_test.loc[2, ['Services', 'Strengths', 'Weaknesses', 'Market Position']] = "Rental,DJ Services,Sound", "Specializes in music events", "High prices", "Niche Player"
            df_test.loc[3, ['Services', 'Strengths', 'Weaknesses', 'Market Position']] = "Rental,Delivery", "Affordable", "Poor quality", "Challenger"
            df_test.to_excel('data/competitor_research.xlsx', index=False)
        except Exception as e:
            print(f"Error adding dummy data to competitor_research.xlsx: {e}")

    if not os.path.exists('data/market_overview.xlsx'):
        from ..data.market_data import MarketDataHandler
        handler = MarketDataHandler()
        handler.create_market_summary_excel()
        
    excel_reporter = MarketStudyExcelReport()
    excel_reporter.create_executive_summary_sheet()
    excel_reporter.create_market_overview_sheet()
    excel_reporter.create_competitor_analysis_sheet()
    excel_reporter.save_workbook()
