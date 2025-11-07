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

    def _add_title_slide(self, title_text: str, subtitle_text: str = "") -> None:
        """Adds a title slide to the presentation."""
        title_slide_layout = self.prs.slide_layouts[0]
        slide = self.prs.slides.add_slide(title_slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        title.text = title_text
        subtitle.text = subtitle_text

    def _add_bullet_slide(self, title_text: str, bullet_points: List[str]) -> None:
        """Adds a slide with a title and bullet points."""
        bullet_slide_layout = self.prs.slide_layouts[1]
        slide = self.prs.slides.add_slide(bullet_slide_layout)
        shapes = slide.shapes

        title_shape = shapes.title
        body_shape = shapes.placeholders[1]

        title_shape.text = title_text

        tf = body_shape.text_frame
        tf.text = bullet_points[0] # Set the first bullet point
        for point in bullet_points[1:]:
            p = tf.add_paragraph()
            p.text = point
            p.level = 0 # Adjust level for sub-bullets if needed

    def _add_chart_slide(self, title_text: str, chart_data: CategoryChartData, chart_type=XL_CHART_TYPE.BAR) -> None:
        """Adds a slide with a chart."""
        slide_layout = self.prs.slide_layouts[5] # Blank slide layout
        slide = self.prs.slides.add_slide(slide_layout)
        shapes = slide.shapes

        # Add chart title
        left = top = width = height = Inches(1) # Placeholder values, will be adjusted
        txBox = shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.text = title_text
        tf.paragraphs[0].font.size = Pt(18)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER

        # Position and size the chart
        chart_left = Inches(0.5)
        chart_top = Inches(1.5)
        chart_width = Inches(12.33)
        chart_height = Inches(5.5)

        graphic_frame = shapes.add_chart(
            chart_type, chart_left, chart_top, chart_width, chart_height, chart_data
        )
        chart = graphic_frame.chart

        # Customize chart appearance
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.value_axis.has_major_gridlines = True

        # Add data labels if applicable (e.g., for bar charts)
        if chart_type == XL_CHART_TYPE.BAR:
            plot = chart.plots[0]
            plot.has_data_labels = True
            data_labels = plot.data_labels
            data_labels.position = XL_LABEL_POSITION.OUTSIDE_END


    def _load_competitor_data_for_chart(self) -> Optional[pd.DataFrame]:
        """Loads and prepares competitor data for charting."""
        if os.path.exists(self.competitor_data_file):
            try:
                df = pd.read_excel(self.competitor_data_file)
                df.columns = df.columns.str.strip()
                # Calculate strength count for the chart
                df['Strength_Count'] = df['Strengths'].astype(str).apply(lambda x: len(x.split(',')) if x and x.strip() else 0)
                # Sort by strength count for better visualization
                return df.sort_values('Strength_Count', ascending=False)
            except Exception as e:
                print(f"Error loading or processing competitor data for chart: {e}")
                return None
        else:
            print(f"Competitor data file not found: {self.competitor_data_file}")
            return None

    def _load_market_data_for_presentation(self) -> Optional[Dict[str, Any]]:
        """Loads market data from JSON."""
        if os.path.exists('data/market_data.json'):
            try:
                with open('data/market_data.json', 'r', encoding='utf-8') as f:
                    return pd.DataFrame(json.load(f)).to_dict('records')[0]
            except Exception as e:
                print(f"Error loading market data from JSON: {e}")
                return None
        else:
            print("Market data JSON file not found.")
            return None

    def generate_presentation(self) -> Optional[str]:
        """
        Generates the market study PowerPoint presentation.

        Returns:
            Optional[str]: The path to the generated PPTX file, or None if generation failed.
        """
        try:
            # --- Slide 1: Title Slide ---
            self._add_title_slide("Étude de Marché - Location Festive Niort", "Analyse et Recommandations Stratégiques")

            # --- Slide 2: Introduction / Business Context ---
            self._add_bullet_slide("Contexte de l'Entreprise", [
                "Entreprise : Location Festive Niort",
                "Secteur : Location de matériel festif (machines à popcorn, barbe à papa, etc.)",
                "Zone Géographique : Niort et ses environs",
                "Objectif : Analyser le marché et identifier les opportunités de croissance."
            ])

            # --- Slide 3: Market Overview ---
            market_info = self._load_market_data_for_presentation()
            if market_info:
                bullet_points = [f"{key.replace('_', ' ').title()}: {', '.join(value) if isinstance(value, list) else value}" for key, value in market_info.items()]
                self._add_bullet_slide("Aperçu du Marché", bullet_points)
            else:
                self._add_bullet_slide("Aperçu du Marché", ["Données du marché non disponibles."])

            # --- Slide 4: Competitor Analysis Summary ---
            competitor_df = self._load_competitor_data_for_chart()
            if competitor_df is not None and not competitor_df.empty:
                competitor_analysis_results = CompetitorAnalysis().analyze_competitor_strengths()
                
                summary_points = [
                    f"Nombre total de concurrents analysés : {competitor_analysis_results.get('total_competitors', 'N/A')}",
                    f"Positionnement moyen sur le marché : {competitor_analysis_results.get('market_position_distribution', {})}",
                    f"Répartition par spécialisation : {competitor_analysis_results.get('competitors_by_specialization', {})}"
                ]
                self._add_bullet_slide("Analyse de la Concurrence (Synthèse)", summary_points)

                # --- Slide 5: Competitor Strengths Chart ---
                chart_data = CategoryChartData()
                chart_data.categories = competitor_df['Competitor'].tolist()
                chart_data.add_series('Nombre de points forts', competitor_df['Strength_Count'].tolist())
                self._add_chart_slide("Nombre de Points Forts par Concurrent", chart_data, XL_CHART_TYPE.BAR)
            else:
                 self._add_bullet_slide("Analyse de la Concurrence", ["Données des concurrents non disponibles pour l'analyse."])


            # --- Slide 6: Differentiation & Opportunities ---
            self._add_bullet_slide("Facteurs Clés de Différenciation et Opportunités", [
                "Possibilité d'acheter des machines directement en Chine (avantage coût).",
                "Réseau de contacts avec des Associations de Parents d'Élèves (APE) pour événements scolaires.",
                "Offrir des services personnalisés et des forfaits attractifs.",
                "Mettre en avant la présence locale et la réactivité.",
                "Développer une stratégie de marketing digital ciblée (SEO local, réseaux sociaux)."
            ])

            # --- Slide 7: Conclusion & Recommendations ---
            self._add_bullet_slide("Conclusion et Recommandations", [
                "Le marché de la location de matériel festif à Niort présente des opportunités.",
                "Exploiter les avantages de sourcing (Chine) pour proposer des prix compétitifs.",
                "Renforcer les partenariats avec les établissements scolaires et organisateurs d'événements.",
                "Investir dans la présence en ligne et la communication ciblée.",
                "Se différencier par la qualité du service et la personnalisation des offres."
            ])

            # --- Save the presentation ---
            self.prs.save(self.output_filename)
            print(f"Market study presentation generated successfully at: {self.output_filename}")
            return self.output_filename

        except Exception as e:
            print(f"An error occurred during presentation generation: {e}")
            return None

