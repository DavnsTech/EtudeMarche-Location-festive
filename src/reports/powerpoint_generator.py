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
        slide_layout = self.prs.slide_layouts[0]  # Title Slide layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = title_text
        title.text_frame.paragraphs[0].font.size = Pt(44)
        title.text_frame.paragraphs[0].font.color.rgb = self.primary_color
        
        if subtitle_text:
            subtitle.text = subtitle_text
            subtitle.text_frame.paragraphs[0].font.size = Pt(28)

    def _add_content_slide(self, title: str, content: str) -> None:
        """Adds a content slide with title and text."""
        slide_layout = self.prs.slide_layouts[1]  # Title and Content layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        slide.shapes.title.text = title
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = self.primary_color
        
        text_frame = slide.placeholders[1].text_frame
        text_frame.clear()  # Remove any default text
        
        p = text_frame.paragraphs[0]
        p.text = content
        p.font.size = Pt(24)

    def _add_table_slide(self, title: str, df: pd.DataFrame) -> None:
        """Adds a slide with a table."""
        slide_layout = self.prs.slide_layouts[5]  # Title Only layout
        slide = self.prs.slides.add_slide(slide_layout)
        
        slide.shapes.title.text = title
        slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = self.primary_color
        
        # Add table
        rows, cols = df.shape
        left = Inches(1)
        top = Inches(2)
        width = Inches(11.33)
        height = Inches(0.5 * (rows + 1))  # +1 for header row
        
        table = slide.shapes.add_table(rows + 1, cols, left, top, width, height).table
        
        # Set column headers
        for i, col in enumerate(df.columns):
            cell = table.cell(0, i)
            cell.text = str(col)
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.primary_color
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.color.rgb = RGBColor(255, 255, 255)  # White text
                paragraph.font.size = Pt(16)
                paragraph.alignment = PP_ALIGN.CENTER
        
        # Fill in data
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                cell = table.cell(i + 1, j)
                cell.text = str(value) if pd.notna(value) else ""
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.size = Pt(14)

    def generate_presentation(self) -> Optional[str]:
        """
        Generates the PowerPoint presentation with market study findings.

        Returns:
            Optional[str]: Path to the generated presentation or None if failed.
        """
        try:
            # Title slide
            self._add_title_slide(
                "Étude de Marché - Location Festive Niort",
                "Analyse du marché local et de la concurrence"
            )

            # Market Overview slide
            if os.path.exists(self.market_data_file):
                market_df = pd.read_excel(self.market_data_file)
                self._add_table_slide("Aperçu du Marché", market_df)
            else:
                self._add_content_slide(
                    "Aperçu du Marché",
                    "Données de marché non disponibles"
                )

            # Competitor Analysis slide
            if os.path.exists(self.competitor_data_file):
                competitor_df = pd.read_excel(self.competitor_data_file)
                # Show only key columns for presentation
                key_columns = ['Competitor', 'Services', 'Pricing Range', 'Market Position']
                competitor_summary = competitor_df[key_columns].head(8)  # Limit to first 8 for readability
                self._add_table_slide("Analyse Concurrentielle", competitor_summary)
            else:
                self._add_content_slide(
                    "Analyse Concurrentielle",
                    "Données concurrentielles non disponibles"
                )

            # Save presentation
            self.prs.save(self.output_filename)
            print(f"PowerPoint presentation saved to: {self.output_filename}")
            return self.output_filename

        except Exception as e:
            print(f"Error generating PowerPoint presentation: {e}")
            return None
