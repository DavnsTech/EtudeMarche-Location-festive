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
                "Total Competitors Identified",
                "Avg Strengths per Competitor",
                "Avg Weaknesses per Competitor"
            ])
            summary_data["Value"].extend([
                competitor_results.get('total_competitors', 'N/A'),
                f"{competitor_results.get('avg_strengths_per_competitor', 0):.2f}" if isinstance(competitor_results.get('avg_strengths_per_competitor'), (int, float)) else 'N/A',
                f"{competitor_results.get('avg_weaknesses_per_competitor', 0):.2f}" if isinstance(competitor_results.get('avg_weaknesses_per_competitor'), (int, float)) else 'N/A'
            ])
            
            # Add market position counts
            for position, count in competitor_results.get('competitor_market_position_counts', {}).items():
                summary_data["Metric"].append(f"Competitors in '{position}' Position")
                summary_data["Value"].append(count)
        else:
            summary_data["Metric"].extend(["Competitor Analysis Status"])
            summary_data["Value"].extend(["No data found"])

        # Add market data status
        summary_data["Metric"].append("Market Data Status")
        summary_data["Value"].append(self.analysis_results.get("market_data_status", "Not processed"))
        
        # Add file paths
        summary_data["Metric"].append("Market Overview Excel")
        summary_data["Value"].append(self.analysis_results.get("market_overview_excel", "Not generated"))
        summary_data["Metric"].append("Competitor Comparison Chart")
        summary_data["Value"].append(self.analysis_results.get("competitor_comparison_chart", "Not generated"))


        results_df = pd.DataFrame(summary_data)
        
        # Save to Excel report
        output_dir = 'reports'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        summary_filepath = os.path.join(output_dir, 'full_analysis_summary.xlsx')
        
        try:
            results_df.to_excel(summary_filepath, index=False)
            print(f"  ✓ Full analysis summary saved to: {summary_filepath}")
            self.analysis_results["full_analysis_summary_report"] = summary_filepath
        except Exception as e:
            print(f"Error saving full analysis summary to {summary_filepath}: {e}")
            self.analysis_results["full_analysis_summary_report"] = f"Failed to save: {e}"

        print("Location Festive Niort Market Analysis complete.")
        return self.analysis_results

if __name__ == "__main__":
    analyzer = MarketAnalyzer()
    results = analyzer.run_full_analysis()
    print("\nAnalysis Results Dictionary:")
    print(results)
