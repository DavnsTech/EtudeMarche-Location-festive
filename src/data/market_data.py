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
        Creates a summary Excel file with market information.

        Returns:
            Optional[str]: Path to the created Excel file or None if failed.
        """
        if not os.path.exists(self.data_dir):
            try:
                os.makedirs(self.data_dir)
                print(f"Created directory: {self.data_dir}")
            except OSError as e:
                print(f"Error creating directory {self.data_dir}: {e}")
                return None

        try:
            # Create a DataFrame for market overview
            overview_data = {
                "Category": [
                    "Industry",
                    "Location",
                    "Target Market",
                    "Seasonality Factors",
                    "Market Trends",
                    "Opportunities",
                    "Challenges"
                ],
                "Details": [
                    self.market_info["industry"],
                    self.market_info["location"],
                    "\n".join(self.market_info["target_market"]),
                    "\n".join(self.market_info["seasonality_factors"]),
                    "\n".join(self.market_info["market_trends"]),
                    "\n".join(self.market_info["potential_opportunities"]),
                    "\n".join(self.market_info["challenges"])
                ]
            }

            df = pd.DataFrame(overview_data)
            df.to_excel(self.market_overview_excel_path, index=False)
            print(f"Market overview Excel created at: {self.market_overview_excel_path}")
            return self.market_overview_excel_path
        except Exception as e:
            print(f"Error creating market overview Excel: {e}")
            return None

    def save_market_data(self) -> str:
        """
        Saves market data to a JSON file.

        Returns:
            str: Status message indicating success or failure.
        """
        try:
            if not os.path.exists(self.data_dir):
                os.makedirs(self.data_dir)

            with open(self.market_data_json_path, 'w', encoding='utf-8') as f:
                json.dump(self.market_info, f, ensure_ascii=False, indent=4)
            return f"Market data saved successfully to {self.market_data_json_path}"
        except Exception as e:
            return f"Error saving market data: {e}"
