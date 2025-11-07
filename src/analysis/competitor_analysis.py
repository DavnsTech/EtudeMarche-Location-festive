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
                'most_common_specializations': [],
                'market_position_distribution': {}
            }

        results: Dict[str, Any] = {}
        results['total_competitors'] = len(self.df)

        # Calculate average number of strengths/weaknesses listed (simple count)
        # Assuming strengths/weaknesses are comma-separated strings
        results['avg_strengths_per_competitor'] = self.df['Strengths'].str.split(',').str.len().mean() if 'Strengths' in self.df.columns else 0
        results['avg_weaknesses_per_competitor'] = self.df['Weaknesses'].str.split(',').str.len().mean() if 'Weaknesses' in self.df.columns else 0

        # Most common specializations
        if 'Specialization' in self.df.columns:
            # Handle potential NaN values and ensure string type
            specializations = self.df['Specialization'].dropna().astype(str).str.lower().str.split(',')
            all_specializations = [item.strip() for sublist in specializations for item in sublist if item.strip()]
            if all_specializations:
                from collections import Counter
                spec_counts = Counter(all_specializations)
                results['most_common_specializations'] = spec_counts.most_common(5) # Top 5
            else:
                results['most_common_specializations'] = []
        else:
            results['most_common_specializations'] = []

        # Market position distribution
        if 'Market Position' in self.df.columns:
            # Handle potential NaN values and ensure string type
            market_position_counts = self.df['Market Position'].dropna().astype(str).str.lower().value_counts()
            results['market_position_distribution'] = market_position_counts.to_dict()
        else:
            results['market_position_distribution'] = {}

        print("Competitor strengths analysis completed.")
        return results

    def generate_comparison_chart(self, output_dir: str = 'reports') -> Optional[str]:
        """
        Generates a bar chart comparing competitors based on a selected metric (e.g., number of strengths or pricing).
        This is a simplified example. A more robust analysis would involve feature engineering.

        Args:
            output_dir (str): Directory to save the chart.

        Returns:
            Optional[str]: Path to the saved chart image, or None if generation failed.
        """
        if self.df is None or self.df.empty:
            print("No competitor data available for chart generation.")
            return None

        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                print(f"Created directory: {output_dir}")
            except OSError as e:
                print(f"Error creating directory {output_dir}: {e}")
                return None

        plt.figure(figsize=(12, 8))
        sns.set_theme(style="whitegrid")

        # Example: Bar chart of number of listed strengths
        if 'Strengths' in self.df.columns:
            strengths_count = self.df['Strengths'].str.split(',').str.len().fillna(0)
            plot_df = pd.DataFrame({
                'Competitor': self.df['Competitor'].str.strip(),
                'Number of Strengths': strengths_count
            })
            plot_df = plot_df.sort_values('Number of Strengths', ascending=False)

            ax = sns.barplot(x='Number of Strengths', y='Competitor', data=plot_df, palette='viridis')
            ax.set_title('Number of Strengths Listed by Competitor', fontsize=16, fontweight='bold')
            ax.set_xlabel("Number of Strengths", fontsize=12)
            ax.set_ylabel("Competitor", fontsize=12)
            plt.tight_layout()

            chart_filename = os.path.join(output_dir, 'competitor_strengths_comparison.png')
            try:
                plt.savefig(chart_filename)
                print(f"Competitor strengths comparison chart saved to '{chart_filename}'.")
                plt.close() # Close the plot to free memory
                return chart_filename
            except Exception as e:
                print(f"Error saving chart to {chart_filename}: {e}")
                plt.close()
                return None
        else:
            print("No 'Strengths' column found for chart generation.")
            return None

