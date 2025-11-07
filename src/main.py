"""
Main entry point for the Location Festive Niort Market Study project
"""

import os
from data.competitor_data import CompetitorDataCollector
from data.market_data import MarketDataHandler
from analysis.market_analysis import MarketAnalyzer
from reports.generate_reports import generate_all_reports

def main():
    print("=== Location Festive Niort - Étude de Marché ===")
    print("Initialisation du projet...\n")
    
    # Ensure data and reports directories exist
    for dir_name in ['data', 'reports']:
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)
            print(f"Created directory: {dir_name}")

    # Step 1: Setup data collection templates
    print("1. Configuration des modèles de collecte de données...")
    
    competitor_collector = CompetitorDataCollector()
    competitor_template_path = competitor_collector.create_competitor_template()
    if competitor_template_path:
        print(f"   ✓ Modèle de recherche concurrentielle créé à : {competitor_template_path}")
    else:
        print("   ✗ Échec de la création du modèle de recherche concurrentielle.")
    
    market_handler = MarketDataHandler()
    market_excel_path = market_handler.create_market_summary_excel()
    market_json_status = market_handler.save_market_data()
    
    if market_excel_path and "saved successfully" in market_json_status:
        print(f"   ✓ Modèles de données de marché créés et sauvegardés.")
    else:
        print("   ✗ Échec de la création ou sauvegarde des modèles de données de marché.")
    
    # Step 2: Run market analysis
    print("\n2. Analyse du marché...")
    analyzer = MarketAnalyzer()
    analysis_results = analyzer.run_full_analysis()
    if analysis_results:
        print("   ✓ Analyse complète effectuée.")
        print(f"     - Résumé de l'analyse sauvegardé à : {analysis_results.get('full_analysis_summary_report', 'N/A')}")
    else:
        print("   ✗ Échec de l'exécution de l'analyse complète.")
    
    # Step 3: Generate reports
    print("\n3. Génération des rapports...")
    excel_file, ppt_file = generate_all_reports()
    
    if excel_file and ppt_file:
        print("   ✓ Rapports générés avec succès.")
        print("\n=== Processus terminé avec succès ===")
        print("\nFichiers créés :")
        print(f"  - {excel_file}")
        print(f"  - {ppt_file}")
        print(f"  - data/competitor_research.xlsx")
        print(f"  - data/market_overview.xlsx")
        print(f"  - reports/full_analysis_summary.xlsx")
    else:
        print("   ✗ Échec de la génération d'un ou plusieurs rapports.")
        print("\n=== Processus terminé avec des erreurs ===")

if __name__ == "__main__":
    main()
