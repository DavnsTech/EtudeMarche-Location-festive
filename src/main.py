"""
Main entry point for the Location Festive Niort Market Study project
"""

import os
from data.competitor_data import CompetitorDataCollector
from data.market_data import MarketDataHandler
from analysis.market_analysis import MarketAnalyzer
from analysis.financial_analysis import FinancialAnalyzer
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
        print(f"   ✗ Échec de la création des modèles de données de marché.")

    # Step 2: Run market analysis
    print("\n2. Exécution de l'analyse de marché...")
    market_analyzer = MarketAnalyzer()
    market_analysis_results = market_analyzer.run_full_analysis()
    print("   ✓ Analyse de marché terminée.")

    # Step 3: Run financial analysis
    print("\n3. Exécution de l'analyse financière...")
    financial_analyzer = FinancialAnalyzer()
    financial_analysis_results = financial_analyzer.run_full_financial_analysis()
    print("   ✓ Analyse financière terminée.")
    
    # Generate executive summary
    executive_summary = financial_analyzer.generate_executive_summary(financial_analysis_results)
    print("\n" + executive_summary)

    # Step 4: Generate reports
    print("\n4. Génération des rapports...")
    excel_report_path, ppt_report_path = generate_all_reports()
    
    if excel_report_path and ppt_report_path:
        print(f"   ✓ Rapport Excel généré : {excel_report_path}")
        print(f"   ✓ Présentation PowerPoint générée : {ppt_report_path}")
    else:
        print("   ✗ Échec de la génération des rapports.")
        
    # Step 5: Save financial analysis results
    print("\n5. Sauvegarde des résultats d'analyse financière...")
    try:
        with open('reports/financial_analysis.json', 'w') as f:
            json.dump(financial_analysis_results, f, indent=2)
        print("   ✓ Résultats d'analyse financière sauvegardés.")
    except Exception as e:
        print(f"   ✗ Échec de la sauvegarde des résultats d'analyse financière : {e}")

    print("\n=== Étude de marché terminée ===")

if __name__ == "__main__":
    main()
