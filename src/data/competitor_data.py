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
            print(f"Competitor research template created at {self.template_filepath}")
            return self.template_filepath
        except Exception as e:
            print(f"Error creating competitor template: {e}")
            return None

    def save_competitor_data(self, data_dict: Dict[str, Any]) -> str:
        """
        Saves competitor data to a JSON file.

        Args:
            data_dict (Dict[str, Any]): A dictionary containing competitor data.

        Returns:
            str: A status message indicating success or failure.
        """
        if not os.path.exists(self.data_dir):
            try:
                os.makedirs(self.data_dir)
                print(f"Created directory: {self.data_dir}")
            except OSError as e:
                return f"Error: Could not create data directory: {e}"

        try:
            with open(self.json_filepath, 'w', encoding='utf-8') as f:
                json.dump(data_dict, f, indent=4, ensure_ascii=False)
            return f"Competitor data saved successfully to {self.json_filepath}"
        except Exception as e:
            return f"Error saving competitor data to JSON: {e}"

    def load_competitor_data(self) -> Optional[Dict[str, Any]]:
        """
        Loads competitor data from the JSON file.

        Returns:
            Optional[Dict[str, Any]]: The loaded data dictionary, or None if the file doesn't exist or an error occurs.
        """
        if not os.path.exists(self.json_filepath):
            print(f"Competitor data file '{self.json_filepath}' not found.")
            return None
        try:
            with open(self.json_filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading competitor data from JSON: {e}")
            return None

    def update_competitor_entry(self, competitor_name: str, data: Dict[str, Any]) -> str:
        """
        Updates a specific competitor's entry in the JSON data.

        Args:
            competitor_name (str): The name of the competitor to update.
            data (Dict[str, Any]): The new data for the competitor.

        Returns:
            str: A status message.
        """
        all_data = self.load_competitor_data()
        if all_data is None:
            all_data = {"competitors": []}

        found = False
        for entry in all_data["competitors"]:
            if entry.get("Competitor") == competitor_name:
                entry.update(data)
                found = True
                break
        
        if not found:
            new_entry = {"Competitor": competitor_name}
            new_entry.update(data)
            all_data["competitors"].append(new_entry)

        return self.save_competitor_data(all_data)

