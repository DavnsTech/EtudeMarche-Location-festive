"""
Main entry point for the Location Festive Niort Market Study project
"""

from data.competitor_data import CompetitorDataCollector
from data.market_data import MarketDataHandler
from analysis.market_analysis import MarketAnalyzer
from reports.generate_reports import generate_all_reports

def main():
    print("=== Location Festive Niort - Étude de Marché ===")
    print("Initialisation du projet...\n")
    
    # Step 1: Setup data collection templates
    print("1. Configuration des modèles de collecte de données...")
    competitor_collector = CompetitorDataCollector()
    competitor_collector.create_competitor_template()
    print("   ✓ Modèle de recherche concurrentielle créé")
    
    market_handler = MarketDataHandler()
    market_handler.create_market_summary_excel()
    market_handler.save_market_data()
    print("   ✓ Modèle de données de marché créé")
    
    # Step 2: Run market analysis
    print("\n2. Analyse du marché...")
    analyzer = MarketAnalyzer()
    analysis_results = analyzer.run_full_analysis()
    print("   ✓ Analyse complète effectuée")
    
    # Step 3: Generate reports
    print("\n3. Génération des rapports...")
    excel_file, ppt_file = generate_all_reports()
    print("   ✓ Rapports générés")
    
    print("\n=== Processus terminé avec succès ===")
    print("\nFichiers créés :")
    print(f"  - {excel_file}")
    print(f"  - {ppt_file}")
    print(f"  - data/competitor_research.xlsx")
    print(f"  - data/market_overview.xlsx")

if __name__ == "__main__":
    main()
