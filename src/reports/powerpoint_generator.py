"""
PowerPoint presentation generator for market study findings
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
import pandas as pd
import os
from typing import Optional, List, Dict, Any

class MarketStudyPresentation:
    def __init__(self, 
                 competitor_data_file: str = 'data/competitor_research.xlsx',
                 market_data_file: str = 'data/market_overview.xlsx'):
        """
        Initializes the PowerPoint presentation generator.

        Args:
            competitor_data_file (str): Path to the competitor research Excel file.
            market_data_file (str): Path to the market overview Excel file.
        """
        self.prs = Presentation()
        # Standard widescreen aspect ratio (16:9)
        self.prs.slide_width = Inches(13.33) 
        self.prs.slide_height = Inches(7.5)
        self.competitor_data_file = competitor_data_file
        self.market_data_file = market_data_file
        self.reports_dir: str = 'reports'
        self.output_filename: str = os.path.join(self.reports_dir, 'Location_Festive_Niort_Market_Study_Presentation.pptx')

        # Define color scheme
        self.primary_color = RGBColor(0x4F, 0x81, 0xBD) # Blue
        self.secondary_color = RGBColor(0xC0, 0x50, 0x4E) # Red
        self.accent_color = RGBColor(0x9B, 0x9B, 0x9B) # Gray

        # Ensure reports directory exists
        if not os.path.exists(self.reports_dir):
            try:
                os.makedirs(self.reports_dir)
                print(f"Created directory: {self.reports_dir}")
            except OSError as e:
                print(f"Error creating directory {self.reports_dir}: {e}")

    def _add_title_slide(self, title_text: str, subtitle: Optional[str] = None) -> None:
        """Adds a title slide to the presentation."""
        title_slide_layout = self.prs.slide_layouts[0] # Title slide layout
        slide = self.prs.slides.add_slide(title_slide_layout)
        title = slide.shapes.title
        subtitle_shape = slide.placeholders[1] if len(slide.placeholders) > 1 else None

        title.text = title_text
        if subtitle and subtitle_shape:
            subtitle_shape.text = subtitle

    def _add_bullet_slide(self, title_text: str, bullet_points: List[str]) -> None:
        """Adds a slide with a title and bullet points."""
        bullet_slide_layout = self.prs.slide_layouts[1] # Title and Content layout
        slide = self.prs.slides.add_slide(bullet_slide_layout)
        title = slide.shapes.title
        body_shape = slide.shapes.placeholders[1]
        tf = body_shape.text_frame

        title.text = title_text
        tf.clear() # Clear existing text
        p = tf.add_paragraph()
        p.text = bullet_points[0]
        p.font.size = Pt(14)
        p.level = 0

        for i in range(1, len(bullet_points)):
            p = tf.add_paragraph()
            p.text = bullet_points[i]
            p.font.size = Pt(12)
            p.level = 1 # Indent for sub-bullets

    def _load_excel_data(self, filepath: str) -> Optional[pd.DataFrame]:
        """Loads data from an Excel file, handling potential errors."""
        if not os.path.exists(filepath):
            print(f"Error: Data file not found at '{filepath}'.")
            return None
        try:
            df = pd.read_excel(filepath)
            # Clean column names in case they have leading/trailing spaces
            df.columns = df.columns.str.strip()
            return df
        except FileNotFoundError:
            print(f"Error: File not found: {filepath}")
            return None
        except Exception as e:
            print(f"Error reading Excel file {filepath}: {e}")
            return None

    def _add_competitor_analysis_slide(self):
        """Adds a slide summarizing competitor analysis."""
        slide_title = "Analyse de la Concurrence"
        bullet_points = []

        df_competitors = self._load_excel_data(self.competitor_data_file)

        if df_competitors is None or df_competitors.empty:
            bullet_points.append("Données de concurrents non trouvées ou vides.")
        else:
            total_competitors = len(df_competitors)
            bullet_points.append(f"Nombre total de concurrents analysés : {total_competitors}")

            # Example: Most common specializations
            if 'Specialization' in df_competitors.columns:
                specializations = df_competitors['Specialization'].dropna().astype(str).str.lower().str.split(',')
                all_specializations = [item.strip() for sublist in specializations for item in sublist if item.strip()]
                if all_specializations:
                    from collections import Counter
                    spec_counts = Counter(all_specializations)
                    top_specs = spec_counts.most_common(3)
                    specs_str = ", ".join([f"{spec} ({count})" for spec, count in top_specs])
                    bullet_points.append(f"Spécialisations courantes : {specs_str}")
                else:
                    bullet_points.append("Aucune spécialisation trouvée.")
            else:
                bullet_points.append("Colonne 'Specialization' manquante.")

            # Example: Market position distribution
            if 'Market Position' in df_competitors.columns:
                market_position_counts = df_competitors['Market Position'].dropna().astype(str).str.lower().value_counts()
                if not market_position_counts.empty:
                    pos_str = ", ".join([f"{pos}: {count}" for pos, count in market_position_counts.items()])
                    bullet_points.append(f"Position sur le marché : {pos_str}")
                else:
                    bullet_points.append("Aucune position sur le marché enregistrée.")
            else:
                bullet_points.append("Colonne 'Market Position' manquante.")

        self._add_bullet_slide(slide_title, bullet_points)

    def _add_market_overview_slide(self):
        """Adds a slide summarizing market overview."""
        slide_title = "Aperçu du Marché"
        bullet_points = []

        # Prefer loading from the market_data.json for direct access to structured data
        market_data_dict = None
        try:
            from ..data.market_data import MarketDataHandler
            handler = MarketDataHandler()
            market_data_dict = handler.get_market_info()
        except ImportError:
            print("Warning: Could not import MarketDataHandler. Trying to load from Excel.")
            df_market = self._load_excel_data(self.market_data_file)
            if df_market is not None:
                try:
                    # Convert DataFrame back to dictionary, assuming 'Category' and 'Details' columns
                    market_data_dict = {}
                    for index, row in df_market.iterrows():
                        market_data_dict[row['Category'].lower().replace(' ', '_')] = row['Details']
                except KeyError:
                    print("Error: Market overview Excel missing 'Category' or 'Details' columns.")
                    market_data_dict = None
        except Exception as e:
            print(f"Error loading market data: {e}")

        if market_data_dict is None:
            bullet_points.append("Données du marché non trouvées ou chargement échoué.")
        else:
            bullet_points.append(f"Industrie : {market_data_dict.get('industry', 'N/A')}")
            bullet_points.append(f"Localisation : {market_data_dict.get('location', 'N/A')}")

            target_market = market_data_dict.get('target_market', [])
            if target_market:
                bullet_points.append("Marché Cible :")
                for item in target_market:
                    bullet_points.append(f"  - {item}")
            
            trends = market_data_dict.get('market_trends', [])
            if trends:
                bullet_points.append("Tendances du Marché :")
                for item in trends:
                    bullet_points.append(f"  - {item}")
            
            opportunities = market_data_dict.get('potential_opportunities', [])
            if opportunities:
                bullet_points.append("Opportunités Potentielles :")
                for item in opportunities:
                    bullet_points.append(f"  - {item}")

        self._add_bullet_slide(slide_title, bullet_points)
        
    def _add_key_takeaways_slide(self):
        """Adds a slide for key takeaways and recommendations."""
        slide_title = "Points Clés et Recommandations"
        bullet_points = [
            "Nécessité d'une veille concurrentielle continue.",
            "Exploiter les contacts APE pour les événements scolaires.",
            "Envisager l'achat direct en Chine pour les machines afin de réduire les coûts.",
            "Développer des offres packagées pour diversifier les revenus.",
            "Renforcer la présence en ligne et sur les réseaux sociaux.",
            "Identifier les niches de marché non couvertes par la concurrence."
        ]
        self._add_bullet_slide(slide_title, bullet_points)


    def generate_full_presentation(self) -> Optional[str]:
        """
        Generates the full PowerPoint presentation and saves it.

        Returns:
            Optional[str]: The path to the saved presentation file, or None if generation failed.
        """
        try:
            self._add_title_slide("Étude de Marché - Location Festive Niort", "Analyse Complète du Secteur")
            self._add_competitor_analysis_slide()
            self._add_market_overview_slide()
            self._add_key_takeaways_slide()

            # Save the presentation
            if not os.path.exists(self.reports_dir):
                try:
                    os.makedirs(self.reports_dir)
                    print(f"Created directory: {self.reports_dir}")
                except OSError as e:
                    print(f"Error creating directory {self.reports_dir}: {e}")
                    return None

            self.prs.save(self.output_filename)
            print(f"Presentation saved successfully to '{self.output_filename}'")
            return self.output_filename

        except ImportError:
            print("Error: python-pptx library is not installed. Please install it: pip install python-pptx")
            return None
        except Exception as e:
            print(f"An error occurred during presentation generation: {e}")
            return None

