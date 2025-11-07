"""
Market data collection and analysis module
"""

import pandas as pd
import json
import os

class MarketDataHandler:
    def __init__(self):
        self.market_info = {
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
                "Importance of social media presence for marketing"
            ]
        }
        self.data_dir = 'data'
        self.market_overview_excel_path = os.path.join(self.data_dir, 'market_overview.xlsx')
        self.market_data_json_path = os.path.join(self.data_dir, 'market_data.json')

    def create_market_summary_excel(self):
        """Create an Excel summary of market data."""
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)

        # Market overview data
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
            print(f"Market overview Excel created at {self.market_overview_excel_path}")
            return self.market_overview_excel_path
        except Exception as e:
            print(f"Error creating market overview Excel: {e}")
            return None

    def save_market_data(self):
        """Save market data to JSON."""
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
        
        try:
            with open(self.market_data_json_path, 'w', encoding='utf-8') as f:
                json.dump(self.market_info, f, ensure_ascii=False, indent=4)
            print(f"Market data saved to {self.market_data_json_path}")
            return f"Data saved successfully to {self.market_data_json_path}"
        except Exception as e:
            print(f"Error saving market data: {e}")
            return None

    def load_market_data(self):
        """Load market data from JSON."""
        try:
            if os.path.exists(self.market_data_json_path):
                with open(self.market_data_json_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                print(f"Market data file {self.market_data_json_path} not found.")
                return {}
        except Exception as e:
            print(f"Error loading market data: {e}")
            return {}
