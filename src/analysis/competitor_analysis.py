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

    def generate_comparison_chart(self):
        """Generate a comparison chart of competitors."""
        if self.df is None or self.df.empty:
            print("No competitor data available for chart generation.")
            return None

        # Create a simple bar chart of market positions
        plt.figure(figsize=(10, 6))
        market_positions = self.df['Market Position'].value_counts()
        
        # Create bar chart
        bars = plt.bar(market_positions.index, market_positions.values, color=['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd'])
        
        # Add value labels on bars
        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, yval + 0.1, int(yval), ha='center', va='bottom')
        
        plt.title('Competitor Market Position Distribution')
        plt.xlabel('Market Position')
        plt.ylabel('Number of Competitors')
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        
        # Save chart
        chart_path = 'reports/competitor_comparison_chart.png'
        if not os.path.exists('reports'):
            os.makedirs('reports')
            
        plt.savefig(chart_path)
        plt.close()
        
        print(f"Competitor comparison chart saved to {chart_path}")
        return chart_path
