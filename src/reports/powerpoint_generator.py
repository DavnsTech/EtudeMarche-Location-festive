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
        title.text = title_text
        
        content_placeholder = slide.placeholders[1]
        text_frame = content_placeholder.text_frame
        text_frame.clear()
        
        for item in content_items:
            p = text_frame.add_paragraph()
            p.text = f"• {item}"
            p.font.size = Pt(18)
            p.level = 0 # Main bullet point

    def _add_chart_slide(self, title_text, chart_data_df, x_col, y_col, chart_type=XL_CHART_TYPE.BAR):
        """Adds a slide with a chart."""
        slide_layout = self.prs.slide_layouts[5] # Blank slide layout for more control
        slide = self.prs.slides.add_slide(slide_layout)
        
        # Add title
        title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.33), Inches(1))
        title_frame = title_shape.text_frame
        p = title_frame.add_paragraph()
        p.text = title_text
        p.font.size = Pt(24)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER

        # Add chart
        chart_data = CategoryChartData()
        chart_data.categories = chart_data_df[x_col].tolist()
        chart_data.add_series(y_col, chart_data_df[y_col].tolist())
        
        x, y, cx, cy = Inches(1), Inches(1.5), Inches(10), Inches(5)
        
        try:
            chart = slide.shapes.add_chart(
                chart_type, x, y, cx, cy, chart_data
            ).chart
            
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False

            # Optional: Add data labels to bars
            plot = chart.plots[0]
            plot.has_data_labels = True
            data_labels = plot.data_labels
            data_labels.position = XL_LABEL_POSITION.OUTSIDE_END # Or INSIDE_END, CENTER etc.
            data_labels.font.size = Pt(8)

            # Style the chart axes
            value_axis = chart.value_axis
            value_axis.has_title = True
            value_axis.axis_title.text_frame.text = y_col
            value_axis.axis_title.font.size = Pt(10)

            category_axis = chart.category_axis
            category_axis.has_title = True
            category_axis.axis_title.text_frame.text = x_col
            category_axis.axis_title.font.size = Pt(10)
            
        except Exception as e:
            print(f"Error adding chart to slide: {e}")
            # Add a text box with error message if chart fails
            error_shape = slide.shapes.add_textbox(x, y, cx, cy)
            error_frame = error_shape.text_frame
            error_frame.text = f"Error generating chart: {e}"
            error_frame.paragraphs[0].font.color.rgb = RGBColor(255, 0, 0)


    def add_title_slide(self):
        """Add title slide to presentation."""
        self._add_title_slide(
            title_text="Étude de Marché - Location Festive Niort",
            subtitle_text="Analyse des opportunités dans la location de matériel festif\nà Niort et ses environs"
        )

    def add_market_overview_slide(self):
        """Add market overview slide."""
        content_items = [
            "Secteur : Location de matériel festif",
            "Localisation : Niort et alentours",
            "Segments cibles : Organisateurs de mariages, événements d'entreprise, écoles, particuliers",
            "Tendances clés : Demande croissante pour expériences uniques, importance du local, présence digitale"
        ]
        self._add_content_slide("Aperçu du Marché", content_items)

    def add_competitive_landscape_slide(self):
        """Add competitive landscape slide."""
        try:
            competitor_df = pd.read_excel(self.competitor_data_file)
            
            if competitor_df.empty:
                content_items = ["Aucune donnée concurrentielle trouvée."]
                self._add_content_slide("Paysage Concurrentiel", content_items)
                return

            # Prepare data for a chart (e.g., number of services per competitor)
            # For simplicity, let's focus on Market Position distribution
            market_pos_counts = competitor_df['Market Position'].value_counts().reset_index()
            market_pos_counts.columns = ['Market Position', 'Number of Competitors']
            
            # Filter out empty positions if any
            market_pos_counts = market_pos_counts[market_pos_counts['Market Position'].notna()]

            if not market_pos_counts.empty:
                self._add_chart_slide(
                    title_text="Répartition des Concurrents par Position sur le Marché",
                    chart_data_df=market_pos_counts,
                    x_col='Market Position',
                    y_col='Number of Competitors',
                    chart_type=XL_CHART_TYPE.BAR # Or PIE chart
                )
            else:
                content_items = ["Aucune donnée de position sur le marché disponible pour la visualisation."]
                self._add_content_slide("Paysage Concurrentiel", content_items)

        except FileNotFoundError:
            content_items = [f"Fichier de données concurrentielles non trouvé: {self.competitor_data_file}"]
            self._add_content_slide("Paysage Concurrentiel", content_items)
        except Exception as e:
            content_items = [f"Erreur lors du chargement des données concurrentielles : {e}"]
            self._add_content_slide("Paysage Concurrentiel", content_items)

    def add_key_findings_slide(self):
        """Add key findings and recommendations slide."""
        content_items = [
            "Opportunité claire pour un fournisseur offrant une large gamme de services et un approvisionnement optimisé.",
            "Potentiel de partenariat avec les établissements scolaires et les organisateurs d'événements privés.",
            "Nécessité d'une forte présence en ligne et sur les réseaux sociaux pour le marketing.",
            "Recommandation : Développer une stratégie de prix compétitive et un service client exceptionnel."
        ]
        self._add_content_slide("Principales Conclusions et Recommandations", content_items)

    def save_presentation(self):
        """Save the presentation to a file."""
        try:
            self.prs.save(self.output_filename)
            print(f"PowerPoint presentation saved to: {self.output_filename}")
            return self.output_filename
        except Exception as e:
            print(f"Error saving PowerPoint presentation: {e}")
            return None

if __name__ == "__main__":
    # Ensure dummy data files exist for testing
    if not os.path.exists('data'):
        os.makedirs('data')
        
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
            print("Dummy data added to competitor_research.xlsx for PPT testing.")
        except Exception as e:
            print(f"Error adding dummy data to competitor_research.xlsx: {e}")

    if not os.path.exists('data/market_overview.xlsx'):
        from ..data.market_data import MarketDataHandler
        handler = MarketDataHandler()
        handler.create_market_summary_excel()
        print("Created dummy market_overview.xlsx for PPT testing.")
        
    ppt_reporter = MarketStudyPresentation()
    ppt_reporter.add_title_slide()
    ppt_reporter.add_market_overview_slide()
    ppt_reporter.add_competitive_landscape_slide()
    ppt_reporter.add_key_findings_slide()
    ppt_reporter.save_presentation()
