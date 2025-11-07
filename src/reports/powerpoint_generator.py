"""
PowerPoint presentation generator for market study findings
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
import pandas as pd
import os

class MarketStudyPresentation:
    def __init__(self, 
                 competitor_data_file='data/competitor_research.xlsx',
                 market_data_file='data/market_overview.xlsx'):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.33)
        self.prs.slide_height = Inches(7.5)
        self.competitor_data_file = competitor_data_file
        self.market_data_file = market_data_file
        self.reports_dir = 'reports'
        self.output_filename = os.path.join(self.reports_dir, 'Location_Festive_Niort_Market_Study_Presentation.pptx')

        # Ensure reports directory exists
        if not os.path.exists(self.reports_dir):
            os.makedirs(self.reports_dir)

    def _add_title_slide(self, title_text, subtitle_text):
        """Adds a title slide."""
        slide_layout = self.prs.slide_layouts[0]  # Title Slide
        slide = self.prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = title_text
        subtitle.text = subtitle_text
        
        # Style the title
        title_format = title.text_frame.paragraphs[0]
        title_format.font.size = Pt(36)
        title_format.font.bold = True
        title_format.alignment = PP_ALIGN.CENTER
        
        subtitle.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

    def _add_content_slide(self, title_text, content_items):
        """Adds a title and content slide."""
        slide_layout = self.prs.slide_layouts[1]  # Title and Content
        slide = self.prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        content = slide.placeholders[1]
        
        title.text = title_text
        
        # Add content
        tf = content.text_frame
        tf.clear()  # Clear any existing content
        
        for item in content_items:
            p = tf.paragraphs[0] if len(tf.paragraphs) == 1 and not tf.paragraphs[0].text else tf.add_paragraph()
            p.text = item
            p.level = 0

    def _add_data_slide_from_excel(self, title_text, excel_file):
        """Adds a slide with data from Excel file."""
        slide_layout = self.prs.slide_layouts[5]  # Title and Blank
        slide = self.prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        title.text = title_text
        
        # Try to load and display data
        if os.path.exists(excel_file):
            try:
                df = pd.read_excel(excel_file)
                # Add a textbox for data
                left = Inches(1)
                top = Inches(1.5)
                width = Inches(11.33)
                height = Inches(5)
                txBox = slide.shapes.add_textbox(left, top, width, height)
                tf = txBox.text_frame
                tf.text = df.head(10).to_string(index=False)  # Show first 10 rows
            except Exception as e:
                print(f"Error loading data from {excel_file}: {e}")
                # Add error message
                left = Inches(1)
                top = Inches(1.5)
                width = Inches(11.33)
                height = Inches(5)
                txBox = slide.shapes.add_textbox(left, top, width, height)
                tf = txBox.text_frame
                tf.text = "Données non disponibles"
        else:
            # Add message about missing file
            left = Inches(1)
            top = Inches(1.5)
            width = Inches(11.33)
            height = Inches(5)
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.text = f"Fichier non trouvé: {excel_file}"

    def create_presentation(self):
        """Create the complete PowerPoint presentation."""
        # Title slide
        self._add_title_slide(
            "Étude de Marché - Location Festive Niort",
            "Analyse du secteur de la location de matériel festif à Niort"
        )
        
        # Executive Summary
        self._add_content_slide(
            "Résumé Exécutif",
            [
                "Analyse du marché de la location de matériel festif à Niort",
                "Identification des principaux concurrents locaux",
                "Évaluation des tendances du marché et opportunités",
                "Recommandations stratégiques pour Location Festive Niort"
            ]
        )
        
        # Market Overview
        market_points = [
            "Industrie: Location de matériel festif",
            "Zone géographique: Niort et ses environs",
            "Segments cibles:",
            "  - Organisateurs de mariages",
            "  - Planificateurs d'événements d'entreprise",
            "  - Établissements scolaires",
            "  - Municipalités",
            "  - Organisateurs de fêtes privées"
        ]
        self._add_content_slide("Aperçu du Marché", market_points)
        
        # Competitor Analysis slide
        self._add_content_slide(
            "Analyse Concurrentielle",
            [
                "12 principaux concurrents identifiés à Niort",
                "Opportunités de différenciation:",
                "  - Approvisionnement direct en Chine",
                "  - Collaboration avec les établissements scolaires (APE)",
                "  - Offres personnalisées pour les événements uniques"
            ]
        )
        
        # Data slides
        self._add_data_slide_from_excel("Données Concurrentielles", self.competitor_data_file)
        self._add_data_slide_from_excel("Aperçu du Marché", self.market_data_file)
        
        # Recommendations
        recommendations = [
            "Développer une stratégie de marketing numérique forte",
            "Mettre en avant l'approvisionnement en Chine comme avantage concurrentiel",
            "Établir des partenariats avec les écoles locales",
            "Créer des offres spéciales pour les clients récurrents",
            "Investir dans des équipements uniques et tendance"
        ]
        self._add_content_slide("Recommandations", recommendations)
        
        # Save presentation
        try:
            self.prs.save(self.output_filename)
            print(f"PowerPoint presentation generated successfully: {self.output_filename}")
            return self.output_filename
        except Exception as e:
            print(f"Error generating PowerPoint presentation: {e}")
            return None
