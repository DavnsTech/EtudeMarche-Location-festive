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
        self.analysis_results["market_overview_excel"] = market_overview_excel_path

        # Step 3: Consolidate and Save Full Analysis Summary
        print("  - Consolidating analysis results...")
        
        # Create a DataFrame for the summary report
        summary_data = {
            "Metric": [],
            "Value": []
        }

        # Add competitor analysis metrics
        comp_analysis = self.analysis_results.get("competitor_analysis", {})
        summary_data["Metric"].append("Total Competitors")
        summary_data["Value"].append(comp_analysis.get('total_competitors', 'N/A'))
        summary_data["Metric"].append("Avg Strengths per Competitor")
        summary_data["Value"].append(f"{comp_analysis.get('avg_strengths_per_competitor', 0.0):.2f}")
        summary_data["Metric"].append("Avg Weaknesses per Competitor")
        summary_data["Value"].append(f"{comp_analysis.get('avg_weaknesses_per_competitor', 0.0):.2f}")
        
        # Add market data metrics
        market_data = self.market_handler.get_market_info() # Assuming get_market_info() exists
        summary_data["Metric"].append("Industry")
        summary_data["Value"].append(market_data.get('industry', 'N/A'))
        summary_data["Metric"].append("Primary Location")
        summary_data["Value"].append(market_data.get('location', 'N/A'))
        summary_data["Metric"].append("Number of Target Segments")
        summary_data["Value"].append(len(market_data.get('target_market', [])))
        summary_data["Metric"].append("Number of Seasonality Factors")
        summary_data["Value"].append(len(market_data.get('seasonality_factors', [])))
        summary_data["Metric"].append("Number of Market Trends")
        summary_data["Value"].append(len(market_data.get('market_trends', [])))

        # Add competitor market position counts
        for position, count in comp_analysis.get('competitor_market_position_counts', {}).items():
            summary_data["Metric"].append(f"Competitors in '{position}' Position")
            summary_data["Value"].append(count)

        # Convert to DataFrame and save
        summary_df = pd.DataFrame(summary_data)
        summary_report_path = os.path.join('reports', 'full_analysis_summary.md') # Save as markdown for readability
        
        try:
            summary_df.to_markdown(summary_report_path, index=False)
            self.analysis_results["full_analysis_summary_report"] = summary_report_path
            print(f"  - Full analysis summary saved to: {summary_report_path}")
        except Exception as e:
            print(f"Error saving full analysis summary to markdown: {e}")
            self.analysis_results["full_analysis_summary_report"] = "Markdown save failed."

        print("Market analysis complete.")
        return self.analysis_results

