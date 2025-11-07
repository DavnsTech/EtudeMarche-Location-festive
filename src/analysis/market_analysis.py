"""
Main market analysis module
"""

import pandas as pd
from .competitor_analysis import CompetitorAnalysis
from ..data.market_data import MarketDataHandler

class MarketAnalyzer:
    def __init__(self):
        self.competitor_analysis = CompetitorAnalysis()
        self.market_handler = MarketDataHandler()
    
    def run_full_analysis(self):
        """Run complete market analysis"""
        print("Running Location Festive Niort Market Analysis...")
        
        # Competitor analysis
        competitor_results = self.competitor_analysis.analyze_competitor_strengths()
        
        # Market data preparation
        self.market_handler.create_market_summary_excel()
        self.market_handler.save_market_data()
        
        # Save results
        analysis_results = {
            "competitor_analysis": competitor_results,
            "market_data_status": "Created market overview files"
        }
        
        # Save to Excel report
        results_df = pd.DataFrame([analysis_results])
        results_df.to_excel('reports/full_analysis_summary.xlsx', index=False)
        
        return analysis_results

if __name__ == "__main__":
    analyzer = MarketAnalyzer()
    results = analyzer.run_full_analysis()
    print("Full analysis completed. Summary saved to reports/full_analysis_summary.xlsx")
