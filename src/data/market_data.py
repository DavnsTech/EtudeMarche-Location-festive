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
        df_overview = pd.DataFrame(overview_data)

        # Target market breakdown (simplified for this example)
        target_data = {
            "Segment": self.market_info["target_market"],
            "Estimated Market Share": [""] * len(self.market_info["target_market"]), # Placeholder
            "Notes": [""] * len(self.market_info["target_market"]) # Placeholder
        }
        df_target = pd.DataFrame(target_data)
        
        try:
            with pd.ExcelWriter(self.market_overview_excel_path) as writer:
                df_overview.to_excel(writer, sheet_name='Market Overview', index=False)
                df_target.to_excel(writer, sheet_name='Target Segments', index=False)
            print(f"Market overview Excel created at: {self.market_overview_excel_path}")
            return self.market_overview_excel_path
        except Exception as e:
            print(f"Error creating market overview Excel: {e}")
            return None

    def save_market_data(self):
        """Save market data dictionary to a JSON file."""
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)
            
        try:
            with open(self.market_data_json_path, 'w') as f:
                json.dump(self.market_info, f, indent=2)
            print(f"Market data saved to {self.market_data_json_path}")
            return "Market data saved successfully."
        except Exception as e:
            print(f"Error saving market data to JSON: {e}")
            return f"Failed to save market data: {e}"

if __name__ == "__main__":
    handler = MarketDataHandler()
    excel_path = handler.create_market_summary_excel()
    json_status = handler.save_market_data()
    
    print(f"Excel path: {excel_path}")
    print(f"JSON status: {json_status}")
