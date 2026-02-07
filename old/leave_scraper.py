# leave_scraper_complete.py
from playwright.sync_api import sync_playwright, TimeoutError
import pandas as pd
import time
from datetime import datetime

SESSION_FILE = "../bnpparibas_session.json"
OUTPUT_FILE = "leave_data.csv"

def parse_dates(date_text):
    """
    Convertit une chaîne de dates au format standard ISO (YYYY-MM-DD).
    Exemple : "12/02/2026 - 14/02/2026" -> "2026-02-12", "2026-02-14"
    """
    try:
        if " - " in date_text:
            start_str, end_str = date_text.split(" - ")
        else:
            start_str = end_str = date_text.strip()
        start_date = datetime.strptime(start_str.strip(), "%d/%m/%Y").date()
        end_date = datetime.strptime(end_str.strip(), "%d/%m/%Y").date()
        return start_date.isoformat(), end_date.isoformat()
    except:
        return date_text, date_text

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state=SESSION_FILE)
    page = context.new_page()

    # Aller sur la page Leave Planning
    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")
    page.wait_for_load_state("networkidle")
    time.sleep(3)  # attendre le rendu complet

    print("✅ Page chargée, récupération des données...")

    # Récupérer toutes les cartes et lignes de congés
    cards = page.query_selector_all("div.leave-item")
    rows = page.query_selector_all("[role='row']")
    elements = cards + rows

    if not elements:
        print("⚠️ Aucun élément trouvé. Vérifie que la page est bien chargée et que ta session est valide.")
        browser.close()
        exit()

    data = []
    for el in elements:
        # Extraction des informations, avec fallback si l'élément n'existe pas
        type_el = el.query_selector(".leave-type")
        date_el = el.query_selector(".leave-dates")
        statut_el = el.query_selector(".leave-status")

        type_conge = type_el.inner_text().strip() if type_el else ""
        dates = date_el.inner_text().strip() if date_el else ""
        statut = statut_el.inner_text().strip() if statut_el else ""

        start_date, end_date = parse_dates(dates)

        data.append({
            "type": type_conge,
            "start_date": start_date,
            "end_date": end_date,
            "statut": statut
        })

    # Export CSV
    df = pd.DataFrame(data)
    df.to_csv(OUTPUT_FILE, index=False)
    print(f"✅ Données extraites et sauvegardées dans {OUTPUT_FILE}")

    browser.close()
