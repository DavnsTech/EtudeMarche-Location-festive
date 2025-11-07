"""
Competitor analysis module
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from typing import Dict, Any, Optional

class CompetitorAnalysis:
    def __init__(self, data_file: str = 'data/competitor_research.xlsx'):
        """
        Initializes the CompetitorAnalysis class.

        Args:
            data_file (str): Path to the competitor research Excel file.
        """
        self.data_file = data_file
        self.df: Optional[pd.DataFrame] = None
        self._load_data()

    def _load_data(self) -> None:
        """Loads competitor data from the specified Excel file."""
        try:
            if os.path.exists(self.data_file):
                self.df = pd.read_excel(self.data_file)
                # Basic data cleaning: strip whitespace from column names
                self.df.columns = self.df.columns.str.strip()
                print(f"Competitor data loaded successfully from '{self.data_file}'.")
            else:
                print(f"Warning: Competitor data file '{self.data_file}' not found. Analysis will be limited.")
                self.df = pd.DataFrame() # Ensure df is at least an empty DataFrame
        except FileNotFoundError:
            print(f"Error: Competitor data file '{self.data_file}' not found.")
            self.df = pd.DataFrame()
        except Exception as e:
            print(f"Error loading competitor data from {self.data_file}: {e}")
            self.df = pd.DataFrame() # Ensure df is at least an empty DataFrame

    def analyze_competitor_strengths(self) -> Dict[str, Any]:
        """
        Analyzes competitor strengths, weaknesses, and market position.

        Returns:
            Dict[str, Any]: A dictionary containing analysis results.
        """
        if self.df is None or self.df.empty:
            print("No competitor data available for strength analysis.")
            return {
                'total_competitors': 0,
                'avg_strengths_per_competitor': 0,
                'top_strengths': [],
                'market_leaders': []
            }

        # Count non-empty entries for strengths
        strength_counts = self.df['Strengths'].str.split(',').str.len().fillna(0)
        avg_strengths = strength_counts.mean() if not strength_counts.empty else 0
        
        # Get top 5 strengths mentioned across competitors
        all_strengths = []
        for strengths in self.df['Strengths'].dropna():
            all_strengths.extend([s.strip() for s in strengths.split(',')])
        
        top_strengths = pd.Series(all_strengths).value_counts().head(5).to_dict()
        
        # Identify market leaders based on market position
        market_leaders = self.df['Market Position'].value_counts().head(3).to_dict()
        
        return {
            'total_competitors': len(self.df),
            'avg_strengths_per_competitor': round(avg_strengths, 2),
            'top_strengths': top_strengths,
            'market_leaders': market_leaders
        }

    def generate_comparison_chart(self) -> Optional[str]:
        """
        Generates a comparison chart of competitors by pricing range.

        Returns:
            Optional[str]: Path to the generated chart or None if failed.
        """
        if self.df is None or self.df.empty:
            print("No competitor data available for chart generation.")
            return None

        # Filter out rows without pricing data
        valid_df = self.df[self.df['Pricing Range'].notna() & (self.df['Pricing Range'] != '')]
        
        if valid_df.empty:
            print("No valid pricing data available for chart generation.")
            return None

        try:
            # Create bar chart of competitors by pricing range
            plt.figure(figsize=(12, 6))
            sns.barplot(x='Competitor', y='Pricing Range', data=valid_df)
            plt.title('Competitor Pricing Range Comparison')
            plt.xlabel('Competitor')
            plt.ylabel('Pricing Range')
            plt.xticks(rotation=45, ha='right')
            
            # Save chart
            chart_path = 'reports/competitor_pricing_comparison.png'
            os.makedirs('reports', exist_ok=True)
            plt.tight_layout()
            plt.savefig(chart_path)
            plt.close()
            
            return chart_path
        except Exception as e:
            print(f"Error generating competitor comparison chart: {e}")
            return None
