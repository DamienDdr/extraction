# leave_scraper_locator.py
from playwright.sync_api import sync_playwright, TimeoutError
import pandas as pd
from datetime import datetime
import time

SESSION_FILE = "../bnpparibas_session.json"
OUTPUT_FILE = "leave_data.csv"

def parse_dates(date_text):
    """
    Convertit une chaîne de dates au format ISO.
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

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")
    page.wait_for_load_state("networkidle")
    time.sleep(3)  # attendre rendu complet SPA

    print("✅ Page chargée, récupération des données...")

    # Récupérer toutes les cartes de congés via locator
    leave_cards = page.locator("div[class*='leave-card']")  # à adapter selon ton DOM
    count = leave_cards.count()
    print(f"🔹 {count} éléments trouvés.")

    if count == 0:
        print("⚠️ Aucun élément trouvé. Vérifie le DOM ou adapte le locator.")
        browser.close()
        exit()

    data = []
    for i in range(count):
        el = leave_cards.nth(i)

        # Extraction avec fallback
        type_conge = el.locator("span[class*='type']").inner_text() if el.locator("span[class*='type']").count() > 0 else ""
        dates = el.locator("span[class*='dates']").inner_text() if el.locator("span[class*='dates']").count() > 0 else ""
        statut = el.locator("span[class*='status']").inner_text() if el.locator("span[class*='status']").count() > 0 else ""

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
