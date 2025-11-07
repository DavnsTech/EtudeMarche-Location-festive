"""
Competitor data collection and management module
"""

import pandas as pd
import requests
from bs4 import BeautifulSoup
import json
import os

class CompetitorDataCollector:
    def __init__(self):
        self.competitors = [
            "LS Réception",
            "Autrement Location",
            "Organi-Sons",
            "SR Événements",
            "Au Comptoir Des Vaisselles",
            "SIEG Event",
            "Geste Scénique",
            "AMB EVENT 79",
            "Ouest Sono Live",
            "Sonovolante",
            "MAX MUSIQUE SA",
            "Carrément Prod"
        ]
        self.data_dir = 'data'
        self.template_filepath = os.path.join(self.data_dir, 'competitor_research.xlsx')
        self.json_filepath = os.path.join(self.data_dir, 'competitor_data.json')

    def create_competitor_template(self):
        """Create a template Excel file for competitor data."""
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)

        data = {
            'Competitor': self.competitors,
            'Website': [''] * len(self.competitors),
            'Services': [''] * len(self.competitors),
            'Pricing Range': [''] * len(self.competitors),
            'Specialization': [''] * len(self.competitors),
            'Strengths': [''] * len(self.competitors),
            'Weaknesses': [''] * len(self.competitors),
            'Market Position': [''] * len(self.competitors)
        }

        df = pd.DataFrame(data)
        try:
            df.to_excel(self.template_filepath, index=False)
            print(f"Competitor research template created at {self.template_filepath}")
            return self.template_filepath
        except Exception as e:
            print(f"Error creating competitor template: {e}")
            return None

    def save_competitor_data(self, data_dict):
        """Save competitor data to JSON."""
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
        
        try:
            with open(self.json_filepath, 'w', encoding='utf-8') as f:
                json.dump(data_dict, f, ensure_ascii=False, indent=4)
            print(f"Competitor data saved to {self.json_filepath}")
            return f"Data saved successfully to {self.json_filepath}"
        except Exception as e:
            print(f"Error saving competitor data: {e}")
            return None

    def load_competitor_data(self):
        """Load competitor data from JSON."""
        try:
            if os.path.exists(self.json_filepath):
                with open(self.json_filepath, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                print(f"Competitor data file {self.json_filepath} not found.")
                return {}
        except Exception as e:
            print(f"Error loading competitor data: {e}")
            return {}
