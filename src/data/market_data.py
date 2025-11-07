"""
Market data collection and management module
"""

import pandas as pd
import json
import os
from typing import Dict, Any, List

class MarketDataHandler:
    def __init__(self):
        """
        Initializes the MarketDataHandler with predefined market information.
        """
        self.market_info: Dict[str, Any] = {
            "industry": "Festive Equipment Rental",
            "location": "Niort, France",
            "target_market": [
                "Wedding organizers",
                "Corporate event planners",
                "Schools and educational institutions",
                "Municipalities for public events",
                "Private party organizers"
            ],
            "seasonality_factors": [
                "Spring/Summer: Weddings, outdoor events",
                "Fall/Winter: Corporate events, holiday parties",
                "Back-to-school season: School events"
            ],
            "market_trends": [
                "Increasing demand for unique event experiences",
                "Growing preference for locally-owned vs. chain providers",
                "Importance of social media presence for marketing",
                "Sustainability in event planning is gaining traction"
            ]
        }
        self.data_dir: str = 'data'
        self.market_overview_excel_path: str = os.path.join(self.data_dir, 'market_overview.xlsx')
        self.market_data_json_path: str = os.path.join(self.data_dir, 'market_data.json')

    def create_market_summary_excel(self) -> Optional[str]:
        """
        Creates an Excel file summarizing key market information.

        Returns:
            Optional[str]: The path to the created Excel file if successful, None otherwise.
        """
        if not os.path.exists(self.data_dir):
            try:
                os.makedirs(self.data_dir)
                print(f"Created directory: {self.data_dir}")
            except OSError as e:
                print(f"Error creating directory {self.data_dir}: {e}")
                return None

        # Market overview data for Excel
        overview_data = {
            "Category": [
                "Industry",
                "Primary Location",
                "Target Segments",
                "Seasonal Peaks",
                "Key Trends"
            ],
            "Details": [
                self.market_info["industry"],
                self.market_info["location"],
                ", ".join(self.market_info["target_market"]),
                ", ".join(self.market_info["seasonality_factors"]),
                ", ".join(self.market_info["market_trends"])
            ]
        }

        df = pd.DataFrame(overview_data)
        try:
            df.to_excel(self.market_overview_excel_path, index=False)
            print(f"Market overview summary Excel created at {self.market_overview_excel_path}")
            return self.market_overview_excel_path
        except Exception as e:
            print(f"Error creating market summary Excel: {e}")
            return None

    def save_market_data(self) -> str:
        """
        Saves the market information to a JSON file.

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
            with open(self.market_data_json_path, 'w', encoding='utf-8') as f:
                json.dump(self.market_info, f, indent=4, ensure_ascii=False)
            return f"Market data saved successfully to {self.market_data_json_path}"
        except Exception as e:
            return f"Error saving market data to JSON: {e}"

    def load_market_data(self) -> Optional[Dict[str, Any]]:
        """
        Loads market data from the JSON file.

        Returns:
            Optional[Dict[str, Any]]: The loaded market data dictionary, or None if the file doesn't exist or an error occurs.
        """
        if not os.path.exists(self.market_data_json_path):
            print(f"Market data file '{self.market_data_json_path}' not found.")
            return None
        try:
            with open(self.market_data_json_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading market data from JSON: {e}")
            return None

    def get_market_info(self) -> Dict[str, Any]:
        """
        Returns the internal market information dictionary.

        Returns:
            Dict[str, Any]: The market information.
        """
        return self.market_info

