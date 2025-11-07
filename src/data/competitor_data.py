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
            print(f"Competitor research template created at: {self.template_filepath}")
            return self.template_filepath
        except Exception as e:
            print(f"Error saving competitor template to {self.template_filepath}: {e}")
            return None

    def load_competitor_data(self) -> Optional[pd.DataFrame]:
        """
        Loads competitor data from the JSON file if it exists.

        Returns:
            Optional[pd.DataFrame]: DataFrame with competitor data, or None if file not found or error.
        """
        if os.path.exists(self.json_filepath):
            try:
                with open(self.json_filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                df = pd.DataFrame(data)
                print(f"Competitor data loaded from {self.json_filepath}")
                return df
            except json.JSONDecodeError:
                print(f"Error: Could not decode JSON from {self.json_filepath}. File might be corrupted.")
                return None
            except Exception as e:
                print(f"Error loading competitor data from {self.json_filepath}: {e}")
                return None
        else:
            print(f"Competitor data file not found at {self.json_filepath}. Please run data collection.")
            return None

    def save_competitor_data(self, df: pd.DataFrame) -> str:
        """
        Saves competitor data to a JSON file.

        Args:
            df (pd.DataFrame): The DataFrame containing competitor data.

        Returns:
            str: A status message indicating success or failure.
        """
        if not os.path.exists(self.data_dir):
            try:
                os.makedirs(self.data_dir)
                print(f"Created directory: {self.data_dir}")
            except OSError as e:
                return f"Error creating directory {self.data_dir}: {e}"

        try:
            df.to_json(self.json_filepath, orient='records', indent=4, force_ascii=False)
            return f"Competitor data saved successfully to {self.json_filepath}"
        except Exception as e:
            return f"Error saving competitor data to {self.json_filepath}: {e}"

