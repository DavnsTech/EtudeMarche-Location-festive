"""
PowerPoint presentation generator for market study findings
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

class MarketStudyPresentation:
    def __init__(self):
        self.prs = Presentation()
        self.prs.slide_width = Inches(13.33)
        self.prs.slide_height = Inches(7.5)
        
    def add_title_slide(self):
        """Add title slide to presentation"""
        slide_layout = self.prs.slide_layouts[0]  # Title Slide
        slide = self.prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        
        title.text = "Étude de Marché - Location Festive Niort"
        subtitle.text = "Analyse des opportunités dans la location de matériel festif\nà Niort et ses environs"
        
        # Style the title
        title_format = title.text_frame.paragraphs[0]
        title_format.font.size = Pt(36)
        title_format.font.bold = True
        
    def add_market_overview_slide(self):
        """Add market overview slide"""
        slide_layout = self.prs.slide_layouts[1]  # Title and Content
        slide = self.prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        title.text = "Aperçu du Marché"
        
        content = slide.placeholders[1]
        text_frame = content.text_frame
        text_frame.clear()
        
        p1 = text_frame.paragraphs[0]
        p1.text = "• Secteur : Location de matériel festif"
        p2 = text_frame.add_paragraph()
        p2.text = "• Localisation : Niort et alentours"
        p3 = text_frame.add_paragraph()
        p3.text = "• Segments cibles : Organisateurs de mariages, événements d'entreprise, écoles"
        p4 = text_frame.add_paragraph()
        p4.text = "• Tendance : Demande croissante pour des expériences événementielles uniques"
        
    def add_competitive_landscape_slide(self):
        """Add competitive landscape slide"""
        slide_layout = self.prs.slide_layouts[1]  # Title and Content
        slide = self.prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        title.text = "Paysage Concurrentiel"
        
        content = slide.placeholders[1]
        text_frame = content.text_frame
        text_frame.clear()
        
        p1 = text_frame.paragraphs[0]
        p1.text = "Principaux concurrents identifiés :"
        p2 = text_frame.add_paragraph()
        p2.text = "• LS Réception"
        p3 = text_frame.add_paragraph()
        p3.text = "• Autrement Location"
        p4 = text_frame.add_paragraph()
        p4.text = "• Organi-Sons"
        p5 = text_frame.add_paragraph()
        p5.text = "• SR Événements"
        p6 = text_frame.add_paragraph()
        p6.text = "+ 8 autres acteurs locaux"
        
    def add_differentiation_slide(self):
        """Add company differentiation slide"""
        slide_layout = self.prs.slide_layouts[1]  # Title and Content
        slide = self.prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        title.text = "Facteurs de Différenciation"
        
        content = slide.placeholders[1]
        text_frame = content.text_frame
        text_frame.clear()
        
        p1 = text_frame.paragraphs[0]
        p1.text = "Avantages concurrentiels de Location Festive Niort :"
        p2 = text_frame.add_paragraph()
        p2.text = "✓ Possibilité d'achat direct en Chine"
        p3 = text_frame.add_paragraph()
        p3.text = "✓ Coûts réduits par rapport aux fournisseurs traditionnels"
        p4 = text_frame.add_paragraph()
        p4.text = "✓ Relations établies avec les APE d'écoles locales"
        p5 = text_frame.add_paragraph()
        p5.text = "✓ Flexibilité dans l'approvisionnement"
        
    def add_recommendations_slide(self):
        """Add recommendations slide"""
        slide_layout = self.prs.slide_layouts[1]  # Title and Content
        slide = self.prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        title.text = "Recommandations Stratégiques"
        
        content = slide.placeholders[1]
        text_frame = content.text_frame
        text_frame.clear()
        
        p1 = text_frame.paragraphs[0]
        p1.text = "Opportunités identifiées :"
        p2 = text_frame.add_paragraph()
        p2.text = "1. Développer une offre spéciale pour les APE scolaires"
        p3 = text_frame.add_paragraph()
        p3.text = "2. Créer des packages événementiels complets"
        p4 = text_frame.add_paragraph()
        p4.text = "3. Mettre en avant l'aspect économique et local"
        p5 = text_frame.add_paragraph()
        p5.text = "4. Investir dans une présence digitale renforcée"
        
    def save_presentation(self, filename='reports/etude_marche_location_festive.pptx'):
        """Save the presentation to file"""
        self.add_title_slide()
        self.add_market_overview_slide()
        self.add_competitive_landscape_slide()
        self.add_differentiation_slide()
        self.add_recommendations_slide()
        self.prs.save(filename)
        return filename

if __name__ == "__main__":
    presentation = MarketStudyPresentation()
    filename = presentation.save_presentation()
    print(f"Market study presentation saved to {filename}")
