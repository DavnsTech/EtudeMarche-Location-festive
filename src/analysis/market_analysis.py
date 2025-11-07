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
                print(f"   Error creating directory 'data': {e}")
                self.analysis_results["market_data_status"] = "Failed to create data directory"
                return self.analysis_results

        # Create market summary Excel
        market_excel_path = self.market_handler.create_market_summary_excel()
        market_json_status = self.market_handler.save_market_data()

        self.analysis_results["market_data_excel"] = market_excel_path
        self.analysis_results["market_data_json"] = market_json_status

        print("   ✓ Market data prepared successfully.")
        print("✓ Full market analysis completed.")

        return self.analysis_results
