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
        else:
            self.analysis_results["competitor_comparison_chart"] = "Chart generation failed."
            print("   ✗ Competitor comparison chart generation failed.")

        # Step 2: Market Data Analysis and Preparation
        print("  - Preparing market data...")
        # Ensure data directory exists
        if not os.path.exists('data'):
            os.makedirs('data')
        
        market_overview_excel_path = self.market_handler.create_market_summary_excel()
        market_data_saved_status = self.market_handler.save_market_data()

        self.analysis_results["market_data_status"] = market_data_saved_status
        if market_overview_excel_path:
             self.analysis_results["market_overview_excel"] = market_overview_excel_path
        else:
             self.analysis_results["market_overview_excel"] = "Excel file generation failed."
             print("   ✗ Market overview Excel file generation failed.")
        
        # Step 3: Integrate and summarize findings (can be expanded)
        print("  - Summarizing findings...")
        # This is a placeholder, actual summary generation would involve more complex logic
        summary_message = "Market and competitor analysis complete. Review generated reports for details."
        self.analysis_results["full_analysis_summary_report"] = summary_message
        print(f"   ✓ Summary: {summary_message}")

        print("Location Festive Niort Market Analysis finished.")
        return self.analysis_results

    def get_analysis_results(self) -> Dict[str, Any]:
        """
        Returns the stored analysis results.
        """
        return self.analysis_results

