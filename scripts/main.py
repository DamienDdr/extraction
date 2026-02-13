#!/usr/bin/env python3
"""
Script principal du DailyRH Scraper

Ce script orchestre l'ensemble du processus :
1. Scraping des données DailyRH via Playwright
2. Export CSV des données brutes
3. Analyse et génération du rapport Excel

Prérequis :
- Session SSO sauvegardée (exécuter save_session.py d'abord)
- Dépendances installées (pip install -r requirements.txt)
- Navigateur Playwright installé (playwright install chromium)

Utilisation :
    python scripts/main.py

Fichiers générés :
- output/leave_planning_2026.csv : Données brutes
- output/rapport_conges_2026.xlsx : Rapport Excel formaté
- dailyrh_scraper.log : Journal d'exécution
"""

import sys
import pandas as pd
from pathlib import Path

# Ajouter le répertoire parent au path pour pouvoir importer src
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.config import OUTPUT_DIR, OUTPUT_CSV, OUTPUT_EXCEL, TARGET_YEAR
from src.logging import setup_logger
from src.scraper import scrape_all_months
from src.excel import analyze_leave_data, create_excel_report


def main():
    """Fonction principale du programme"""
    
    # Configuration du logging
    logger = setup_logger(
        name="dailyrh_scraper",
        log_file="dailyrh_scraper.log",
        level="DEBUG"  # Changer en "DEBUG" pour plus de détails
    )
    
    try:
        logger.info("="*60)
        logger.info("DailyRH Leave Planning Scraper - Démarrage")
        logger.info("="*60)
        
        # Créer le répertoire de sortie
        output_path = Path(OUTPUT_DIR)
        output_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"Répertoire de sortie : {output_path.absolute()}")
        
        # Étape 1 : Scraping
        logger.info(f"Étape 1/3 : Scraping des données pour l'année {TARGET_YEAR}")
        all_records = scrape_all_months(TARGET_YEAR)
        
        if not all_records:
            logger.error("Aucune donnée collectée - Arrêt du programme")
            sys.exit(1)
        
        # Étape 2 : Export CSV
        csv_path = output_path / OUTPUT_CSV
        logger.info(f"Étape 2/3 : Export CSV ({len(all_records)} lignes)")
        df = pd.DataFrame(all_records)
        df.to_csv(csv_path, index=False, encoding="utf-8")
        logger.info(f"CSV créé : {csv_path}")
        
        # Statistiques de collecte
        months_scraped = sorted(set(r['date'][:7] for r in all_records))
        logger.info(f"Mois collectés : {len(months_scraped)}/12")
        for month_str in months_scraped:
            month_data = [r for r in all_records if r['date'][:7] == month_str]
            logger.info(f"  {month_str} : {len(month_data)} lignes")
        
        # Étape 3 : Génération Excel
        excel_path = output_path / OUTPUT_EXCEL
        logger.info("Étape 3/3 : Génération du rapport Excel")
        stats = analyze_leave_data(str(csv_path))
        create_excel_report(stats, str(csv_path), str(excel_path))
        
        logger.info("="*60)
        logger.info("✅ Traitement terminé avec succès")
        logger.info("="*60)
        logger.info(f"Fichiers générés :")
        logger.info(f"  - CSV : {csv_path}")
        logger.info(f"  - Excel : {excel_path}")
        logger.info(f"  - Log : dailyrh_scraper.log")
        
    except KeyboardInterrupt:
        logger.warning("\n⚠️ Interruption manuelle détectée")
        sys.exit(130)
    
    except Exception as e:
        logger.exception(f"❌ Erreur fatale : {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
