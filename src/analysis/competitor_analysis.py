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
                'avg_strengths_per_competitor': 0.0,
                'avg_weaknesses_per_competitor': 0.0,
                'competitor_market_position_counts': {}
            }

        # Count non-empty entries for strengths/weaknesses
        # Assuming strengths/weaknesses can be comma-separated lists or single entries
        strength_counts = self.df['Strengths'].dropna().apply(
            lambda x: len(str(x).split(',')) if pd.notnull(x) and str(x).strip() else 0
        )
        weakness_counts = self.df['Weaknesses'].dropna().apply(
            lambda x: len(str(x).split(',')) if pd.notnull(x) and str(x).strip() else 0
        )

        # Analyze Market Position
        market_position_counts = self.df['Market Position'].value_counts().to_dict()

        analysis = {
            'total_competitors': len(self.df),
            'avg_strengths_per_competitor': float(strength_counts.mean()) if not strength_counts.empty else 0.0,
            'avg_weaknesses_per_competitor': float(weakness_counts.mean()) if not weakness_counts.empty else 0.0,
            'competitor_market_position_counts': market_position_counts
        }
        
        print("  - Competitor strengths and market position analyzed.")
        return analysis

    def generate_comparison_chart(self, output_dir: str = 'reports') -> Optional[str]:
        """
        Generates a bar chart comparing the number of strengths and weaknesses per competitor.

        Args:
            output_dir (str): Directory to save the chart.

        Returns:
            Optional[str]: Path to the saved chart file, or None if failed.
        """
        if self.df is None or self.df.empty:
            print("No competitor data available to generate comparison chart.")
            return None

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Prepare data for chart
        chart_data = self.df.copy()
        chart_data['Num_Strengths'] = chart_data['Strengths'].dropna().apply(
            lambda x: len(str(x).split(',')) if pd.notnull(x) and str(x).strip() else 0
        )
        chart_data['Num_Weaknesses'] = chart_data['Weaknesses'].dropna().apply(
            lambda x: len(str(x).split(',')) if pd.notnull(x) and str(x).strip() else 0
        )

        plt.figure(figsize=(12, 7))
        sns.set_theme(style="whitegrid")

        # Melt the DataFrame for easier plotting with seaborn
        melted_df = chart_data.melt(id_vars='Competitor', value_vars=['Num_Strengths', 'Num_Weaknesses'],
                                    var_name='Attribute', value_name='Count')

        ax = sns.barplot(x='Competitor', y='Count', hue='Attribute', data=melted_df, palette="viridis")

        plt.title('Comparison of Strengths and Weaknesses per Competitor', fontsize=16)
        plt.xlabel('Competitor', fontsize=12)
        plt.ylabel('Number of Attributes', fontsize=12)
        plt.xticks(rotation=45, ha='right')
        plt.legend(title='Attribute')
        plt.tight_layout()

        chart_filename = os.path.join(output_dir, 'competitor_strengths_weaknesses_comparison.png')
        try:
            plt.savefig(chart_filename)
            print(f"Competitor comparison chart saved to: {chart_filename}")
            plt.close() # Close the plot to free memory
            return chart_filename
        except Exception as e:
            print(f"Error saving competitor comparison chart: {e}")
            plt.close()
            return None

