"""
Main market analysis module
"""

import pandas as pd
from .competitor_analysis import CompetitorAnalysis
from ..data.market_data import MarketDataHandler
import os
from typing import Dict, Any, Optional

class MarketAnalyzer:
    def __init__(self):
        """
        Initializes the MarketAnalyzer with necessary data handlers and analyzers.
        """
        self.competitor_analyzer = CompetitorAnalysis()
        self.market_handler = MarketDataHandler()
        self.analysis_results: Dict[str, Any] = {}

    def run_full_analysis(self) -> Dict[str, Any]:
        """
        Runs a comprehensive market analysis, including competitor and market data.

        Returns:
            Dict[str, Any]: A dictionary containing all analysis results and paths to generated assets.
        """
        print("Running Location Festive Niort Market Analysis...")

        # Step 1: Competitor Analysis
        print("  - Performing competitor analysis...")
        competitor_results = self.competitor_analyzer.analyze_competitor_strengths()
        self.analysis_results["competitor_analysis"] = competitor_results

        # Generate competitor comparison chart
        competitor_chart_path = self.competitor_analyzer.generate_comparison_chart()
        if competitor_chart_path:
            self.analysis_results["competitor_comparison_chart"] = competitor_chart_path
            print(f"   ✓ Competitor comparison chart saved: {competitor_chart_path}")
        else:
            self.analysis_results["competitor_comparison_chart"] = "Chart generation failed."
            print("   ✗ Competitor comparison chart generation failed.")

        # Step 2: Market Data Analysis and Preparation
        print("  - Preparing market data...")
        # Ensure data directory exists
        if not os.path.exists('data'):
            try:
                os.makedirs('data')
                print("   Created 'data' directory.")
            except OSError as e:
                print(f"   Error creating 'data' directory: {e}")
                # Depending on severity, you might want to exit or handle this more robustly
                print("   ✗ Failed to prepare market data due to directory error.")
                return {} # Return empty if critical directory creation fails

        market_overview_excel_path = self.market_handler.create_market_summary_excel()
        market_data_saved_status = self.market_handler.save_market_data()

        market_data_info = {
            "market_overview_excel": market_overview_excel_path,
            "market_data_json_status": market_data_saved_status
        }
        self.analysis_results["market_data_preparation"] = market_data_info

        if market_overview_excel_path and "saved successfully" in market_data_saved_status:
            print("   ✓ Market overview Excel and JSON data prepared.")
        else:
            print("   ✗ Market overview Excel or JSON data preparation incomplete.")

        # Step 3: Integrate findings (example: combining competitor insights with market opportunities)
        print("  - Integrating analysis findings...")
        combined_insights = {
            "competitor_summary": competitor_results,
            "market_overview": self.market_handler.get_market_info(), # Assuming MarketDataHandler has this method
            "potential_opportunities": self.market_handler.get_market_info().get("potential_opportunities", [])
        }
        self.analysis_results["combined_insights"] = combined_insights
        print("   ✓ Analysis findings integrated.")

        print("Location Festive Niort Market Analysis completed.")
        return self.analysis_results

    # Placeholder for potential future methods
    # def analyze_market_demand(self): ...
    # def identify_niche_opportunities(self): ...

