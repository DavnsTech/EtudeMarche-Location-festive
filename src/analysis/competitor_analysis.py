"""
Competitor analysis module
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

class CompetitorAnalysis:
    def __init__(self, data_file='data/competitor_research.xlsx'):
        self.data_file = data_file
        self.df = None
        try:
            if os.path.exists(self.data_file):
                self.df = pd.read_excel(data_file)
            else:
                print(f"Warning: Competitor data file '{data_file}' not found. Analysis will be limited.")
        except Exception as e:
            print(f"Error loading competitor data from {data_file}: {e}")
            self.df = pd.DataFrame() # Ensure df is at least an empty DataFrame

    def analyze_competitor_strengths(self):
        """Analyze competitor strengths and weaknesses based on counts."""
        if self.df is None or self.df.empty:
            print("No competitor data available for strength analysis.")
            return {
                'total_competitors': 0,
                'avg_strengths_per_competitor': 0,
                'avg_weaknesses_per_competitor': 0,
                'competitor_market_position_counts': {}
            }

        # Count non-empty entries for strengths/weaknesses
        # Assuming strengths/weaknesses can be comma-separated lists
        strength_counts = self.df['Strengths'].dropna().apply(lambda x: len(str(x).split(',')) if pd.notnull(x) and str(x).strip() else 0)
        weakness_counts = self.df['Weaknesses'].dropna().apply(lambda x: len(str(x).split(',')) if pd.notnull(x) and str(x).strip() else 0)

        # Analyze Market Position
        market_position_counts = self.df['Market Position'].value_counts().to_dict()

        analysis = {
            'total_competitors': len(self.df),
            'avg_strengths_per_competitor': strength_counts.mean() if not strength_counts.empty else 0,
            'avg_weaknesses_per_competitor': weakness_counts.mean() if not weakness_counts.empty else 0,
            'competitor_market_position_counts': market_position_counts
        }

        return analysis

    def generate_comparison_chart(self, output_dir='reports'):
        """Generate a comparison chart of competitors based on a simple metric (e.g., number of services)."""
        if self.df is None or self.df.empty:
            print("No competitor data available for chart generation.")
            return None

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        chart_filename = os.path.join(output_dir, 'competitor_services_chart.png')

        # For demonstration, let's use 'Services' as a proxy for complexity or offering size.
        # A more robust analysis would involve defining specific quantifiable metrics.
        self.df['Service_Count'] = self.df['Services'].dropna().apply(lambda x: len(str(x).split(',')) if pd.notnull(x) and str(x).strip() else 0)

        # Sort by number of services for better visualization
        df_sorted = self.df.sort_values('Service_Count', ascending=False).head(10) # Limit to top 10 for clarity

        if df_sorted.empty:
            print("No data to plot for competitor comparison chart.")
            return None

        plt.figure(figsize=(12, 8))
        sns.barplot(x='Service_Count', y='Competitor', data=df_sorted, palette='viridis')
        plt.title("Top 10 Competitors by Number of Services Offered")
        plt.xlabel("Number of Services")
        plt.ylabel("Competitor")
        plt.tight_layout()

        try:
            plt.savefig(chart_filename)
            print(f"Competitor comparison chart saved to: {chart_filename}")
            return chart_filename
        except Exception as e:
            print(f"Error saving competitor comparison chart: {e}")
            return None
        finally:
            plt.close() # Close the plot to free memory

if __name__ == "__main__":
    # This main block is for testing the module directly
    # Ensure data/competitor_research.xlsx exists or is created by CompetitorDataCollector
    from ..data.competitor_data import CompetitorDataCollector
    
    # Create template if it doesn't exist for testing
    if not os.path.exists('data/competitor_research.xlsx'):
        print("Creating competitor template for testing...")
        collector = CompetitorDataCollector()
        collector.create_competitor_template()
        # Manually adding some dummy data for testing the analysis
        try:
            df_test = pd.read_excel('data/competitor_research.xlsx')
            df_test.loc[0, ['Services', 'Strengths', 'Weaknesses', 'Market Position']] = "Rental,Delivery,Setup,Teardown,Decoration,Lighting,Sound,DJ Equipment,Stage Rental,Dance Floor,Photo Booth,Special Effects", "Wide range of services,Good pricing,Local presence", "Limited online presence,Old website,Slow response time", "Market Leader"
            df_test.loc[1, ['Services', 'Strengths', 'Weaknesses', 'Market Position']] = "Rental,Delivery,Setup,Teardown,Sound", "Excellent customer service,Fast delivery", "High prices,Limited inventory", "Challenger"
            df_test.loc[2, ['Services', 'Strengths', 'Weaknesses', 'Market Position']] = "Rental,DJ Services,Sound", "Specializes in music events,Good sound quality", "High prices,Limited equipment variety", "Niche Player"
            df_test.to_excel('data/competitor_research.xlsx', index=False)
            print("Dummy data added to competitor_research.xlsx")
        except Exception as e:
            print(f"Error adding dummy data: {e}")

    competitor_analyzer = CompetitorAnalysis()
    analysis = competitor_analyzer.analyze_competitor_strengths()
    print("\nCompetitor Analysis Results:")
    print(analysis)

    chart_path = competitor_analyzer.generate_comparison_chart()
    if chart_path:
        print(f"Chart generated at: {chart_path}")
