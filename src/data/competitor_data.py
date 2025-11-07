"""
Competitor data collection and management module
"""

import pandas as pd
import json
import os
from typing import List, Dict, Any, Optional

class CompetitorDataCollector:
    def __init__(self):
        """
        Initializes the CompetitorDataCollector with a predefined list of competitors.
        """
        self.competitors: List[str] = [
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
        self.data_dir: str = 'data'
        self.template_filepath: str = os.path.join(self.data_dir, 'competitor_research.xlsx')
        self.json_filepath: str = os.path.join(self.data_dir, 'competitor_data.json')

    def create_competitor_template(self) -> Optional[str]:
        """
        Creates a template Excel file for competitor data collection.

        Returns:
            Optional[str]: The path to the created template file if successful, None otherwise.
        """
        if not os.path.exists(self.data_dir):
            try:
                os.makedirs(self.data_dir)
                print(f"Created directory: {self.data_dir}")
            except OSError as e:
                print(f"Error creating directory {self.data_dir}: {e}")
                return None

        # Check if template already exists to avoid overwriting without explicit intent
        if os.path.exists(self.template_filepath):
            print(f"Template file '{self.template_filepath}' already exists. Skipping creation.")
            return self.template_filepath

        data = {
            'Competitor': self.competitors,
            'Website': [''] * len(self.competitors),
            'Services': [''] * len(self.competitors),
            'Pricing Range': [''] * len(self.competitors),
            'Strengths': [''] * len(self.competitors),
            'Weaknesses': [''] * len(self.competitors),
            'Differentiation': [''] * len(self.competitors),
            'Market Position': [''] * len(self.competitors),
            'Social Media Presence': [''] * len(self.competitors)
        }

        try:
            df = pd.DataFrame(data)
            df.to_excel(self.template_filepath, index=False)
            print(f"Competitor research template created at: {self.template_filepath}")
            return self.template_filepath
        except Exception as e:
            print(f"Error creating competitor research template: {e}")
            return None

    def load_competitor_data(self) -> Optional[pd.DataFrame]:
        """
        Loads competitor data from the Excel file.

        Returns:
            Optional[pd.DataFrame]: DataFrame with competitor data or None if failed.
        """
        try:
            if os.path.exists(self.template_filepath):
                df = pd.read_excel(self.template_filepath)
                return df
            else:
                print(f"Competitor data file not found: {self.template_filepath}")
                return None
        except Exception as e:
            print(f"Error loading competitor data: {e}")
            return None
