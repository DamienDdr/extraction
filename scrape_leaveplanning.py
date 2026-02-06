from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime, timedelta
import calendar
import time

SESSION_FILE = "bnpparibas_session.json"
OUTPUT_FILE = "leave_planning_final.csv"

YEAR = 2026
MONTH = 2

START_DATE = datetime(YEAR, MONTH, 1)
DAYS_IN_MONTH = calendar.monthrange(YEAR, MONTH)[1]

def date_str(d):
    return d.strftime("%Y/%m/%d")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state=SESSION_FILE)
    page = context.new_page()

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")
    page.wait_for_load_state("networkidle")
    time.sleep(4)

    print("✅ Page chargée")

    # ===== NOMS COLLABORATEURS (colonne gauche)
    name_cells = page.locator("td.dhx_matrix_scell")
    collaborators = [
        name_cells.nth(i).inner_text().strip()
        for i in range(name_cells.count())
        if name_cells.nth(i).inner_text().strip()
    ]

    rows = page.locator("tr.dhx_row_item")

    print(f"👥 {len(collaborators)} collaborateurs détectés")

    data = []

    for r in range(rows.count()):
        collab = collaborators[r]
        tr = rows.nth(r)
        cells = tr.locator("td")

        for day_idx in range(DAYS_IN_MONTH):
            day = START_DATE + timedelta(days=day_idx)

            # 🔹 WEEKEND (calcul calendaire)
            if day.weekday() >= 5:
                data.append({
                    "collaborateur": collab,
                    "date": date_str(day),
                    "type": "WEEKEND",
                    "title": ""
                })
                continue

            cell = cells.nth(day_idx)

            # 🔹 CONGÉS = bandeau vert
            leave_events = cell.locator(".dhx_matrix_line > div")
            if leave_events.count() > 0:
                title = (leave_events.first.get_attribute("title") or "").strip()
                data.append({
                    "collaborateur": collab,
                    "date": date_str(day),
                    "type": "CONGES",
                    "title": title
                })
                continue

            # 🔹 TÉLÉTRAVAIL = icône maison
            home_icon = cell.locator("i[class*='home'], span[class*='home'], div[class*='home']")
            if home_icon.count() > 0:
                data.append({
                    "collaborateur": collab,
                    "date": date_str(day),
                    "type": "TELETRAVAIL",
                    "title": ""
                })
                continue

            # 🔹 PRÉSENT SUR SITE
            data.append({
                "collaborateur": collab,
                "date": date_str(day),
                "type": "PRESENT",
                "title": ""
            })

    df = pd.DataFrame(data)
    df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8")

    print(f"📁 Export terminé → {OUTPUT_FILE}")
    browser.close()
