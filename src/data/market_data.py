"""
Market data collection and analysis module
"""

import pandas as pd
import json

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
    
    def create_market_summary_excel(self):
        """Create an Excel summary of market data"""
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
        
        # Target market breakdown
        target_data = {
            "Segment": self.market_info["target_market"],
            "Estimated Market Share": [""] * len(self.market_info["target_market"]),
            "Growth Potential": [""] * len(self.market_info["target_market"]),
            "Marketing Approach": [""] * len(self.market_info["target_market"])
        }
        
        df_target = pd.DataFrame(target_data)
        
        # Write to Excel with multiple sheets
        with pd.ExcelWriter('data/market_overview.xlsx') as writer:
            df_overview.to_excel(writer, sheet_name='Market Overview', index=False)
            df_target.to_excel(writer, sheet_name='Target Segments', index=False)
            
        return df_overview, df_target
    
    def save_market_data(self):
        """Save market data to JSON"""
        with open('data/market_data.json', 'w') as f:
            json.dump(self.market_info, f, indent=2)

if __name__ == "__main__":
    handler = MarketDataHandler()
    overview, target = handler.create_market_summary_excel()
    handler.save_market_data()
    print("Market data summary created at data/market_overview.xlsx")
