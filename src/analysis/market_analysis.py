"""
Main market analysis module
"""

import pandas as pd
from .competitor_analysis import CompetitorAnalysis
from ..data.market_data import MarketDataHandler
import os

class MarketAnalyzer:
    def __init__(self):
        self.competitor_analyzer = CompetitorAnalysis()
        self.market_handler = MarketDataHandler()
        self.analysis_results = {}

    def run_full_analysis(self):
        """Run complete market analysis."""
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
        if competitor_results:
            summary_data["Metric"].extend([
                "Total Competitors Analyzed",
                "Average Strengths per Competitor",
                "Average Weaknesses per Competitor"
            ])
            summary_data["Value"].extend([
                competitor_results['total_competitors'],
                round(competitor_results['avg_strengths_per_competitor'], 2),
                round(competitor_results['avg_weaknesses_per_competitor'], 2)
            ])
            
            # Add market position counts
            for position, count in competitor_results['competitor_market_position_counts'].items():
                summary_data["Metric"].append(f"Competitors with '{position}' Position")
                summary_data["Value"].append(count)

        # Create summary DataFrame
        summary_df = pd.DataFrame(summary_data)
        
        # Save summary to Excel
        summary_excel_path = 'reports/full_analysis_summary.xlsx'
        if not os.path.exists('reports'):
            os.makedirs('reports')
            
        try:
            summary_df.to_excel(summary_excel_path, index=False)
            print(f"Full analysis summary saved to {summary_excel_path}")
            self.analysis_results["full_analysis_summary_report"] = summary_excel_path
        except Exception as e:
            print(f"Error saving analysis summary: {e}")
            self.analysis_results["full_analysis_summary_report"] = None

        return self.analysis_results
