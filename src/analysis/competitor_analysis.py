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
                'avg_weaknesses_per_competitor': 0,
                'competitors_by_specialization': {},
                'market_position_distribution': {}
            }

        analysis_results: Dict[str, Any] = {}

        # Total number of competitors
        analysis_results['total_competitors'] = len(self.df)

        # Average number of strengths and weaknesses listed per competitor
        # Assuming 'Strengths' and 'Weaknesses' columns contain comma-separated values or are to be interpreted qualitatively
        # For simplicity, we'll count non-empty entries. A more sophisticated analysis could parse comma-separated values.
        analysis_results['avg_strengths_per_competitor'] = self.df['Strengths'].astype(str).apply(lambda x: len(x.split(',')) if x and x.strip() else 0).mean()
        analysis_results['avg_weaknesses_per_competitor'] = self.df['Weaknesses'].astype(str).apply(lambda x: len(x.split(',')) if x and x.strip() else 0).mean()

        # Competitors by specialization
        if 'Specialization' in self.df.columns:
            analysis_results['competitors_by_specialization'] = self.df['Specialization'].value_counts().to_dict()
        else:
            analysis_results['competitors_by_specialization'] = {}

        # Market position distribution
        if 'Market Position' in self.df.columns:
            analysis_results['market_position_distribution'] = self.df['Market Position'].value_counts().to_dict()
        else:
            analysis_results['market_position_distribution'] = {}
            
        print("Competitor strengths analysis completed.")
        return analysis_results

    def generate_comparison_chart(self, output_dir: str = 'reports') -> Optional[str]:
        """
        Generates a bar chart comparing competitors based on a selected metric (e.g., number of strengths).

        Args:
            output_dir (str): Directory to save the chart.

        Returns:
            Optional[str]: Path to the saved chart file, or None if generation failed.
        """
        if self.df is None or self.df.empty:
            print("No competitor data available to generate comparison chart.")
            return None

        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                print(f"Created directory: {output_dir}")
            except OSError as e:
                print(f"Error creating directory {output_dir}: {e}")
                return None

        chart_filename = os.path.join(output_dir, 'competitor_strength_comparison.png')

        try:
            # Prepare data for chart: Count strengths for each competitor
            # This assumes 'Strengths' column contains comma-separated strengths or is qualitative.
            # We'll count non-empty entries for simplicity.
            df_chart = self.df.copy()
            df_chart['Strength_Count'] = df_chart['Strengths'].astype(str).apply(lambda x: len(x.split(',')) if x and x.strip() else 0)
            
            # Sort by strength count for better visualization
            df_chart = df_chart.sort_values('Strength_Count', ascending=False)

            plt.figure(figsize=(12, 8))
            sns.barplot(x='Strength_Count', y='Competitor', data=df_chart, palette='viridis')
            plt.title('Number of Strengths Listed Per Competitor', fontsize=16)
            plt.xlabel('Number of Strengths', fontsize=12)
            plt.ylabel('Competitor', fontsize=12)
            plt.tight_layout()
            plt.savefig(chart_filename)
            plt.close()
            print(f"Competitor comparison chart saved to: {chart_filename}")
            return chart_filename
        except Exception as e:
            print(f"Error generating competitor comparison chart: {e}")
            return None

