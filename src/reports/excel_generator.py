"""
Excel report generator for market study data
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows

class MarketStudyExcelReport:
    def __init__(self):
        self.wb = Workbook()
        
    def create_executive_summary_sheet(self):
        """Create executive summary sheet"""
        ws = self.wb.active
        ws.title = "Résumé Exécutif"
        
        # Add title
        ws['A1'] = "ÉTUDE DE MARCHE - LOCATION FESTIVE NIORT"
        ws['A1'].font = Font(size=16, bold=True)
        
        # Add key metrics
        data = [
            ["Indicateur", "Valeur"],
            ["Taille du marché estimée", "À déterminer"],
            ["Nombre de concurrents identifiés", "12"],
            ["Positionnement concurrentiel", "Avantage sur approvisionnement"],
            ["Croissance annuelle prévue", "À déterminer"],
            ["Investissement initial requis", "À déterminer"]
        ]
        
        for row in data:
            ws.append(row)
            
        # Style the table
        for cell in ws["A2:B7"]:
            for c in cell:
                c.font = Font(bold=True if c.row == 2 else False)
                
        # Add some spacing
        ws.append([])
        
        # Add key insights
        ws.append(["Points Clés", ""])
        ws['A9'].font = Font(bold=True)
        insights = [
            "Opportunité distincte via l'approvisionnement chinois",
            "Relations existantes avec les APE scolaires",
            "Marché local sous-exploité selon nos premières analyses",
            "Potentiel de croissance dans les événements privés"
        ]
        
        for insight in insights:
            ws.append([insight, ""])
            
    def create_competitor_analysis_sheet(self):
        """Create competitor analysis sheet"""
        ws = self.wb.create_sheet("Analyse Concurrentielle")
        
        # Headers
        headers = ["Concurrent", "Spécialisation", "Forces", "Faiblesses", "Position sur le marché"]
        ws.append(headers)
        
        # Apply header styling
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
        # Sample data (would be populated from research)
        competitors = [
            ["LS Réception", "Événements corporatifs", "Expérience", "Prix élevés", "Établi"],
            ["Autrement Location", "Mariages", "Service personnalisé", "Gamme limitée", "Croissant"],
            ["Organi-Sons", "Sonorisation", "Technicité", "Hors spécialité principale", "Établi"],
            ["SR Événements", "Événements privés", "Créativité", "Capacité limitée", "Croissant"],
            ["Au Comptoir Des Vaisselles", "Location vaisselle", "Logistique", "Hors spécialité principale", "Établi"]
        ]
        
        for competitor in competitors:
            ws.append(competitor)
            
        # Auto-adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
            
    def create_financial_projections_sheet(self):
        """Create financial projections sheet"""
        ws = self.wb.create_sheet("Projections Financières")
        
        ws['A1'] = "PROJECTIONS FINANCIÈRES - ANNÉE 1"
        ws['A1'].font = Font(size=14, bold=True)
        
        # Sample projection data
        projection_data = [
            ["Mois", "Chiffre d'affaires (€)", "Coût des ventes (€)", "Marge brute (€)", "Dépenses opérationnelles (€)", "Résultat net (€)"],
            ["Janvier", 5000, 2000, 3000, 2500, 500],
            ["Février", 4500, 1800, 2700, 2500, 200],
            ["Mars", 6000, 2400, 3600, 2500, 1100],
            ["Avril", 7500, 3000, 4500, 2500, 2000],
            ["Mai", 9000, 3600, 5400, 2500, 2900],
            ["Juin", 12000, 4800, 7200, 2500, 4700],
            ["Juillet", 15000, 6000, 9000, 2500, 6500],
            ["Août", 14000, 5600, 8400, 2500, 5900],
            ["Septembre", 10000, 4000, 6000, 2500, 3500],
            ["Octobre", 8000, 3200, 4800, 2500, 2300],
            ["Novembre", 11000, 4400, 6600, 2500, 4100],
            ["Décembre", 13000, 5200, 7800, 2500, 5300]
        ]
        
        for row in projection_data:
            ws.append(row)
            
        # Style headers
        for cell in ws[2]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
        # Create a simple chart
        chart = BarChart()
        chart.title = "Projection du Chiffre d'Affaires Mensuel"
        chart.x_axis.title = "Mois"
        chart.y_axis.title = "Chiffre d'Affaires (€)"
        
        # Data for chart
        data = Reference(ws, min_col=2, min_row=2, max_row=13, max_col=2)
        categories = Reference(ws, min_col=1, min_row=3, max_row=13)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        
        ws.add_chart(chart, "H2")
        
    def save_workbook(self, filename='reports/etude_marche_location_festive.xlsx'):
        """Save workbook to file"""
        self.create_executive_summary_sheet()
        self.create_competitor_analysis_sheet()
        self.create_financial_projections_sheet()
        self.wb.save(filename)
        return filename

if __name__ == "__main__":
    report = MarketStudyExcelReport()
    filename = report.save_workbook()
    print(f"Market study Excel report saved to {filename}")
