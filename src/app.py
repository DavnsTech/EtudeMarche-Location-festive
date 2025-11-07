"""
Main entry point for the Location Festive Niort Market Study project
"""

import os
from data.competitor_data import CompetitorDataCollector
from data.market_data import MarketDataHandler
from analysis.market_analysis import MarketAnalyzer
from reports.generate_reports import generate_all_reports
import json

def main():
    """
    Orchestrates the market study process from data setup to report generation.
    """
    print("=== Location Festive Niort - Étude de Marché ===")
    print("Initialisation du projet...\n")

    # Ensure data and reports directories exist
    for dir_name in ['data', 'reports']:
        if not os.path.exists(dir_name):
            try:
                os.makedirs(dir_name)
                print(f"Created directory: {dir_name}")
            except OSError as e:
                print(f"Error creating directory {dir_name}: {e}")
                # Depending on severity, you might want to exit or handle this more robustly
                if dir_name == 'data':
                    print("Critical error: Could not create data directory. Exiting.")
                    return

    # Step 1: Setup data collection templates and initial data
    print("1. Configuration des modèles de collecte de données et données initiales...")

    # Competitor Data Setup
    competitor_collector = CompetitorDataCollector()
    competitor_template_path = competitor_collector.create_competitor_template()
    if competitor_template_path:
        print(f"   ✓ Modèle de recherche concurrentielle créé à : {competitor_template_path}")
    else:
        print("   ✗ Échec de la création du modèle de recherche concurrentielle.")

    # Market Data Setup
    market_handler = MarketDataHandler()
    market_excel_path = market_handler.create_market_summary_excel()
    market_json_status = market_handler.save_market_data()

    if market_excel_path and "saved successfully" in market_json_status:
        print(f"   ✓ Modèles de données de marché créés et sauvegardés (Excel: {market_excel_path}, JSON status: {market_json_status}).")
    else:
        print(f"   ✓ Modèles de données de marché créés (Excel: {market_excel_path}, JSON status: {market_json_status}).")

    # Step 2: Run market and competitor analysis
    print("\n2. Analyse du marché et de la concurrence...")
    analyzer = MarketAnalyzer()
    analysis_results = analyzer.run_full_analysis() # Changed to capture results

    if analysis_results:
        print("   ✓ Analyse complète terminée.")
        # You can print or process analysis_results here if needed
        # print("\nAnalyse Results Summary:")
        # print(json.dumps(analysis_results, indent=2))
    else:
        print("   ✗ L'analyse complète a échoué.")

    # Step 3: Generate reports
    print("\n3. Génération des rapports...")
    excel_report_path, ppt_report_path = generate_all_reports()

    if excel_report_path:
        print(f"   ✓ Rapport Excel généré : {excel_report_path}")
    else:
        print("   ✗ Échec de la génération du rapport Excel.")

    if ppt_report_path:
        print(f"   ✓ Rapport PowerPoint généré : {ppt_report_path}")
    else:
        print("   ✗ Échec de la génération du rapport PowerPoint.")

    print("\n=== Étude de marché terminée. ===")

if __name__ == "__main__":
    main()
