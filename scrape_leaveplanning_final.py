from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date, timedelta
import time
import re

SESSION_FILE = "bnpparibas_session.json"
OUTPUT_FILE = "leave_planning_fevrier_2026_final.csv"

MONTH_START = date(2026, 2, 1)
MONTH_END = date(2026, 2, 28)
NB_DAYS = (MONTH_END - MONTH_START).days + 1

DAY_WIDTH_PX = 35

def px_to_day_range(left_px, width_px):
    start = int(round(left_px / DAY_WIDTH_PX))
    end = int(round((left_px + width_px) / DAY_WIDTH_PX)) - 1
    return start, end

def is_weekend(d):
    return d.weekday() >= 5

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

    # ===== Récupération NOMS (méthode fiable du script 2)
    name_cells = page.locator("td.dhx_matrix_scell")
    collaborators = [
        name_cells.nth(i).inner_text().strip()
        for i in range(name_cells.count())
        if name_cells.nth(i).inner_text().strip()
    ]

    rows = page.locator("tr.dhx_row_item")
    row_count = rows.count()

    if row_count == 0:
        raise RuntimeError("❌ Aucune ligne détectée")

    print(f"👥 {row_count} collaborateurs détectés")

    records = []

    for r in range(row_count):
        row = rows.nth(r)

        # =====================
        # NOM COLLABORATEUR
        # =====================
        if r < len(collaborators):
            name = collaborators[r]
        else:
            name = "INCONNU"

        print(f"🧩 Ligne {r} → {name}")

        # =====================
        # INIT MOIS
        # =====================
        planning = {}
        for i in range(NB_DAYS):
            d = MONTH_START + timedelta(days=i)
            planning[i] = {
                "date": date_str(d),
                "type": "WEEKEND" if is_weekend(d) else "PRESENT",
                "title": ""
            }

        matrix = row.locator(".dhx_matrix_line")
        if matrix.count() == 0:
            continue

        events = matrix.locator("div")

        for e in range(events.count()):
            ev = events.nth(e)
            cls = ev.get_attribute("class") or ""
            title = ev.get_attribute("title") or ""
            style = ev.get_attribute("style") or ""

            left_match = re.search(r"left:\s*(\d+)px", style)
            width_match = re.search(r"width:\s*(\d+)px", style)

            if not left_match or not width_match:
                continue

            start_day, end_day = px_to_day_range(
                int(left_match.group(1)),
                int(width_match.group(1))
            )

            for day_idx in range(start_day, end_day + 1):
                if day_idx < 0 or day_idx >= NB_DAYS:
                    continue

                if "grey_cell_weekend" in cls:
                    planning[day_idx]["type"] = "WEEKEND"

                elif "validated_vcell" in cls:
                    planning[day_idx]["type"] = "CONGES"
                    planning[day_idx]["title"] = title

                elif "telework" in cls:
                    if planning[day_idx]["type"] == "PRESENT":
                        planning[day_idx]["type"] = "TELETRAVAIL"
                        planning[day_idx]["title"] = title

        for i in range(NB_DAYS):
            info = planning[i]
            records.append({
                "collaborateur": name,
                "date": info["date"],
                "type": info["type"],
                "title": info["title"]
            })

    df = pd.DataFrame(records)
    df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8")

    print(f"📁 Export terminé → {OUTPUT_FILE}")
    browser.close()
