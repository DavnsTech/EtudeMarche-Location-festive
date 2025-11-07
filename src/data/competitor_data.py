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
            'Specialization': [''] * len(self.competitors),
            'Strengths': [''] * len(self.competitors),
            'Weaknesses': [''] * len(self.competitors),
            'Market Position': [''] * len(self.competitors)
        }

        df = pd.DataFrame(data)

        try:
            df.to_excel(self.template_filepath, index=False)
            print(f"Competitor template created at: {self.template_filepath}")
            return self.template_filepath
        except Exception as e:
            print(f"Error saving competitor template to {self.template_filepath}: {e}")
            return None

    def load_competitor_data(self) -> Optional[pd.DataFrame]:
        """
        Loads competitor data from the JSON file if it exists.

        Returns:
            Optional[pd.DataFrame]: A DataFrame with competitor data, or None if the file is not found or an error occurs.
        """
        if not os.path.exists(self.json_filepath):
            print(f"Competitor data file '{self.json_filepath}' not found. Please run data collection first.")
            return None
        try:
            with open(self.json_filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            df = pd.DataFrame(data)
            print(f"Competitor data loaded from '{self.json_filepath}'.")
            return df
        except FileNotFoundError:
            print(f"Error: Competitor data file '{self.json_filepath}' not found.")
            return None
        except json.JSONDecodeError:
            print(f"Error: Could not decode JSON from '{self.json_filepath}'. File might be corrupted.")
            return None
        except Exception as e:
            print(f"Error loading competitor data from {self.json_filepath}: {e}")
            return None

    def save_competitor_data(self, df: pd.DataFrame) -> str:
        """
        Saves the provided DataFrame to a JSON file.

        Args:
            df (pd.DataFrame): The DataFrame to save.

        Returns:
            str: A status message indicating whether the save operation was successful.
        """
        if not os.path.exists(self.data_dir):
            try:
                os.makedirs(self.data_dir)
                print(f"Created directory: {self.data_dir}")
            except OSError as e:
                print(f"Error creating directory {self.data_dir}: {e}")
                return f"Failed to save: Could not create data directory."

        try:
            # Convert NaN to None for JSON compatibility
            df_json = df.where(pd.notnull(df), None)
            df_json.to_json(self.json_filepath, indent=4, orient='records', force_ascii=False)
            return f"Competitor data saved successfully to '{self.json_filepath}'."
        except Exception as e:
            return f"Error saving competitor data to {self.json_filepath}: {e}"

