"""
Market data collection and management module
"""

import pandas as pd
import json
import os
from typing import Dict, Any, List, Optional

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
            ],
            "potential_opportunities": [
                "Partnerships with local schools for fundraising events (leveraging APE contacts)",
                "Sourcing unique machines directly from China for competitive pricing",
                "Offering package deals for specific event types (e.g., birthdays, corporate picnics)"
            ],
            "challenges": [
                "High initial investment for equipment",
                "Seasonal demand fluctuations",
                "Competition from established players",
                "Logistics and maintenance of equipment"
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

        # Check if Excel file already exists
        if os.path.exists(self.market_overview_excel_path):
            print(f"Market overview Excel file '{self.market_overview_excel_path}' already exists. Skipping creation.")
            return self.market_overview_excel_path
            
        # Convert dictionary to a list of dictionaries for easier DataFrame creation
        # Each key-value pair will become a row in the Excel sheet
        data_for_excel = []
        for key, value in self.market_info.items():
            data_for_excel.append({'Category': key.replace('_', ' ').title(), 'Details': value})

        df = pd.DataFrame(data_for_excel)

        try:
            df.to_excel(self.market_overview_excel_path, index=False)
            print(f"Market overview Excel file created at: {self.market_overview_excel_path}")
            return self.market_overview_excel_path
        except Exception as e:
            print(f"Error saving market overview Excel to {self.market_overview_excel_path}: {e}")
            return None

    def save_market_data(self) -> str:
        """
        Saves the market information dictionary to a JSON file.

        Returns:
            str: A status message indicating the success or failure of the save operation.
        """
        if not os.path.exists(self.data_dir):
            try:
                os.makedirs(self.data_dir)
                print(f"Created directory: {self.data_dir}")
            except OSError as e:
                print(f"Error creating directory {self.data_dir}: {e}")
                return "Failed to save: Could not create data directory."

        try:
            with open(self.market_data_json_path, 'w', encoding='utf-8') as f:
                json.dump(self.market_info, f, indent=4, ensure_ascii=False)
            print(f"Market data saved to '{self.market_data_json_path}'.")
            return "Market data saved successfully."
        except Exception as e:
            return f"Error saving market data to {self.market_data_json_path}: {e}"

    def load_market_data(self) -> Optional[Dict[str, Any]]:
        """
        Loads market data from the JSON file.

        Returns:
            Optional[Dict[str, Any]]: The loaded market data, or None if the file is not found or an error occurs.
        """
        if not os.path.exists(self.market_data_json_path):
            print(f"Market data file '{self.market_data_json_path}' not found.")
            return None
        try:
            with open(self.market_data_json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            print(f"Market data loaded from '{self.market_data_json_path}'.")
            return data
        except FileNotFoundError:
            print(f"Error: Market data file '{self.market_data_json_path}' not found.")
            return None
        except json.JSONDecodeError:
            print(f"Error: Could not decode JSON from '{self.market_data_json_path}'. File might be corrupted.")
            return None
        except Exception as e:
            print(f"Error loading market data from {self.market_data_json_path}: {e}")
            return None

    def get_market_info(self) -> Dict[str, Any]:
        """
        Returns the current market information dictionary.
        """
        return self.market_info

