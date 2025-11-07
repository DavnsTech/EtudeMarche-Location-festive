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
            os.makedirs(dir_name)
            print(f"Created directory: {dir_name}")

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
    analysis_results = analyzer.run_full_analysis()
    
    if analysis_results:
        print("   ✓ Analyse complète effectuée.")
        print(f"     - Résumé de l'analyse : {analysis_results.get('full_analysis_summary_report', 'N/A')}")
        print(f"     - Chemin graphique comparaison concurrents : {analysis_results.get('competitor_comparison_chart', 'N/A')}")
    else:
        print("   ✗ Échec de l'exécution de l'analyse complète.")

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

    # Generate QA report
    print("\n4. Génération du rapport QA...")
    qa_report = generate_qa_report(analysis_results, excel_report_path, ppt_report_path)
    qa_report_path = "qa_report.json"
    try:
        with open(qa_report_path, 'w', encoding='utf-8') as f:
            json.dump(qa_report, f, indent=4)
        print(f"   ✓ Rapport QA généré : {qa_report_path}")
    except Exception as e:
        print(f"   ✗ Erreur lors de la génération du rapport QA : {e}")

    print("\n=== Fin de l'étude de marché ===")

def generate_qa_report(analysis_results: Dict[str, Any], excel_path: Optional[str], ppt_path: Optional[str]) -> Dict[str, Any]:
    """
    Generates a QA report summarizing the process and outcomes.
    """
    qa_status = {
        "project_name": "Location Festive Niort - Market Study",
        "overall_status": "Passed" if analysis_results and excel_path and ppt_path else "Failed",
        "steps": {
            "data_setup": {
                "status": "Passed" if analysis_results is not None else "Failed",
                "message": "Data templates and initial data setup completed." if analysis_results is not None else "Data setup failed."
            },
            "analysis": {
                "status": "Passed" if analysis_results and analysis_results.get("full_analysis_summary_report") else "Failed",
                "message": analysis_results.get("full_analysis_summary_report", "Analysis failed.") if analysis_results else "Analysis not performed."
            },
            "report_generation": {
                "status": "Passed" if excel_path and ppt_path else "Failed",
                "message": f"Excel report: {excel_path if excel_path else 'Failed'}. PowerPoint report: {ppt_path if ppt_path else 'Failed'}."
            }
        },
        "generated_files": {
            "excel_report": excel_path if excel_path else "Not generated",
            "powerpoint_report": ppt_path if ppt_path else "Not generated",
            "qa_report": "qa_report.json" # This file itself
        },
        "analysis_details": analysis_results.get("competitor_analysis", {}), # Include some analysis details
        "errors": [] # Placeholder for specific errors if any occurred
    }

    if not analysis_results:
        qa_status["errors"].append("Market analysis did not return results.")
    if not excel_path:
        qa_status["errors"].append("Excel report generation failed.")
    if not ppt_path:
        qa_status["errors"].append("PowerPoint report generation failed.")
        
    if qa_status["errors"]:
        qa_status["overall_status"] = "Failed"

    return qa_status

if __name__ == "__main__":
    main()
