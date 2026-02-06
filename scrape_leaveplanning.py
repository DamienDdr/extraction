from playwright.sync_api import sync_playwright
import pandas as pd
import time
from datetime import datetime, timedelta

SESSION_FILE = "bnpparibas_session.json"
OUTPUT_FILE = "leave_planning.csv"

START_DATE = datetime(2026, 2, 1)

def col_to_date(col_idx):
    return (START_DATE + timedelta(days=col_idx)).strftime("%Y/%m/%d")

def classify(classes, title):
    c = (classes or "").lower()
    t = (title or "").lower()

    if "weekend" in c:
        return "WEEKEND"
    if "telework" in c or "teletravail" in t:
        return "TELETRAVAIL"
    if "cong" in t:
        return "CONGES"
    return "AUTRE"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state=SESSION_FILE)
    page = context.new_page()

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")
    page.wait_for_load_state("networkidle")
    time.sleep(4)

    print("✅ Page chargée")

    # =====================================================
    # 1️⃣ NOMS – colonne gauche FIXE
    # =====================================================
    name_cells = page.locator("td.dhx_matrix_scell, div.dhx_scell")
    names = []

    for i in range(name_cells.count()):
        txt = (name_cells.nth(i).inner_text() or "").strip()
        if txt:
            names.append(txt)

    print(f"👤 {len(names)} noms détectés")

    # =====================================================
    # 2️⃣ PLANNING – lignes principales
    # =====================================================
    rows = page.locator("tr.dhx_row_item")
    row_count = rows.count()

    print(f"👥 {row_count} collaborateurs détectés")

    data = []

    for r in range(row_count):
        row = rows.nth(r)
        collaborateur = names[r] if r < len(names) else f"COLLAB_{r+1}"

        cells = row.locator("td")

        for day_idx in range(cells.count()):
            cell = cells.nth(day_idx)
            events = cell.locator(".dhx_matrix_line > div")
            cell_class = (cell.get_attribute("class") or "").lower()

            # WEEKEND sans événement
            if "weekend" in cell_class and events.count() == 0:
                data.append({
                    "collaborateur": collaborateur,
                    "date": col_to_date(day_idx),
                    "type": "WEEKEND",
                    "title": ""
                })

            for e in range(events.count()):
                ev = events.nth(e)

                classes = ev.get_attribute("class")
                title = ev.get_attribute("title") or ev.text_content()

                data.append({
                    "collaborateur": collaborateur,
                    "date": "" if "telework" in (classes or "").lower() else col_to_date(day_idx),
                    "type": classify(classes, title),
                    "title": (title or "").strip()
                })

    df = pd.DataFrame(data)
    df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8")

    print(f"📁 Export terminé → {OUTPUT_FILE}")
    browser.close()
