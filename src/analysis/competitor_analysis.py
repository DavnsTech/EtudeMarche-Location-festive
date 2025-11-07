"""
Competitor analysis module
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

class CompetitorAnalysis:
    def __init__(self, data_file='data/competitor_research.xlsx'):
        self.data_file = data_file
        try:
            self.df = pd.read_excel(data_file)
        except FileNotFoundError:
            print(f"Data file {data_file} not found. Creating empty dataframe.")
            self.df = pd.DataFrame()
    
    def analyze_competitor_strengths(self):
        """Analyze competitor strengths and weaknesses"""
        if self.df.empty:
            return None
            
        # Count non-empty entries for strengths/weaknesses
        strength_counts = self.df['Strengths'].dropna().apply(lambda x: len(str(x).split(',')) if pd.notnull(x) else 0)
        weakness_counts = self.df['Weaknesses'].dropna().apply(lambda x: len(str(x).split(',')) if pd.notnull(x) else 0)
        
        analysis = {
            'total_competitors': len(self.df),
            'avg_strengths_per_competitor': strength_counts.mean() if not strength_counts.empty else 0,
            'avg_weaknesses_per_competitor': weakness_counts.mean() if not weakness_counts.empty else 0
        }
        
        return analysis
    
    def generate_comparison_chart(self):
        """Generate a comparison chart of competitors"""
        if self.df.empty:
            return None
            
        # Create a simple visualization
        plt.figure(figsize=(12, 8))
        plt.title("Competitor Analysis Overview")
        plt.xlabel("Competitors")
        plt.ylabel("Metrics")
        plt.xticks(rotation=45, ha='right')
        
        # This is a placeholder - actual implementation would depend on collected data
        competitors = self.df['Competitor'].tolist()[:10]  # Limit to first 10 for display
        metrics = list(range(len(competitors)))
        
        plt.bar(competitors, metrics, color='skyblue')
        plt.tight_layout()
        plt.savefig('reports/competitor_comparison.png')
        plt.close()
        
        return 'reports/competitor_comparison.png'

if __name__ == "__main__":
    analysis = CompetitorAnalysis()
    results = analysis.analyze_competitor_strengths()
    if results:
        print("Competitor Analysis Results:")
        print(f"Total Competitors: {results['total_competitors']}")
        print(f"Avg Strengths per Competitor: {results['avg_strengths_per_competitor']:.2f}")
        print(f"Avg Weaknesses per Competitor: {results['avg_weaknesses_per_competitor']:.2f}")
        
        chart_path = analysis.generate_comparison_chart()
        if chart_path:
            print(f"Comparison chart saved to {chart_path}")
