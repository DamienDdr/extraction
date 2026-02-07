import re
import math
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date, timedelta
import time

SESSION_FILE = "bnpparibas_session.json"
OUTPUT_FILE = "leave_planning_fevrier_2026_final.csv"

MONTH_START = date(2026, 2, 1)
MONTH_END = date(2026, 2, 28)
NB_DAYS = (MONTH_END - MONTH_START).days + 1


def is_weekend(d):
    return d.weekday() >= 5


def date_str(d):
    return d.strftime("%Y/%m/%d")


def determine_event_type(cls):
    """Détermine le type d'événement selon les classes CSS"""
    if "grey_cell_weekend" in cls:
        return None

    if "telework" in cls and "validated_vcell" in cls:
        return "TELETRAVAIL"
    elif "validated_vcell" in cls:
        return "CONGES"

    return None


def pixels_to_days(left_px, width_px, col_width, nb_days):
    """Convertit position/largeur en pixels vers indices de jours"""
    center_px = left_px + width_px / 2
    center_day = round(center_px / col_width)

    if width_px < col_width * 0.8:
        return max(0, min(nb_days - 1, center_day)), max(0, min(nb_days - 1, center_day))

    start_idx = max(0, int(math.floor(left_px / col_width)))
    end_idx = min(nb_days - 1, int(math.ceil((left_px + width_px) / col_width)) - 1)

    return start_idx, end_idx


with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state=SESSION_FILE)
    page = context.new_page()

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")
    page.wait_for_load_state("networkidle")
    time.sleep(4)

    print("✅ Page chargée")

    rows = page.locator("tr.dhx_row_item")
    row_count = rows.count()
    if row_count == 0:
        raise RuntimeError("❌ Aucune ligne détectée")

    print(f"👥 {row_count} collaborateurs détectés\n")

    records = []

    for r in range(row_count):
        row = rows.nth(r)

        name_cell = row.locator("td.dhx_matrix_scell").first
        name = name_cell.inner_text().strip() if name_cell.count() > 0 else "INCONNU"

        if name == "Mes Collègues" or not name:
            print(f"⏭️  Ligne {r} → '{name}' (ignorée)\n")
            continue

        print(f"🧩 Ligne {r} → {name}")

        # INIT
        planning = {}
        for i in range(NB_DAYS):
            d = MONTH_START + timedelta(days=i)
            planning[i] = {
                "date": date_str(d),
                "type": "WEEKEND" if is_weekend(d) else "PRESENT",
                "title": ""
            }

        matrix_div = row.locator(".dhx_matrix_line").first
        matrix_width = matrix_div.evaluate("el => el.offsetWidth")
        col_width = matrix_width / NB_DAYS

        events = matrix_div.locator("div[class*='cell'], div[class*='event']")
        event_count = events.count()
        print(f"   📦 {event_count} événements détectés")

        count_teletravail = 0
        count_conges = 0

        for e in range(event_count):
            ev = events.nth(e)
            cls = ev.get_attribute("class") or ""
            title = ev.get_attribute("title") or ""
            style = ev.get_attribute("style") or ""

            left_match = re.search(r"left:\s*([\d.]+)px", style)
            width_match = re.search(r"width:\s*([\d.]+)px", style)
            if not left_match or not width_match:
                continue

            left_px = float(left_match.group(1))
            width_px = float(width_match.group(1))

            event_type = determine_event_type(cls)
            if not event_type:
                continue

            start_idx, end_idx = pixels_to_days(left_px, width_px, col_width, NB_DAYS)

            if event_type == "TELETRAVAIL":
                count_teletravail += 1
            elif event_type == "CONGES":
                count_conges += 1

            for day_idx in range(start_idx, end_idx + 1):
                # ✅ NE JAMAIS ÉCRASER UN WEEKEND !
                if planning[day_idx]["type"] == "WEEKEND":
                    continue

                # Les congés ont priorité sur le télétravail et présent
                if event_type == "CONGES" or planning[day_idx]["type"] == "PRESENT":
                    planning[day_idx]["type"] = event_type
                    planning[day_idx]["title"] = title

        print(f"   ✅ 🏠 TT: {count_teletravail} | 🏖️ CP: {count_conges}")

        for i in range(NB_DAYS):
            info = planning[i]
            records.append({
                "collaborateur": name,
                "date": info["date"],
                "type": info["type"],
                "title": info["title"]
            })
        print()

    df = pd.DataFrame(records)
    df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8")
    print(f"📁 Export terminé → {OUTPUT_FILE}")

    browser.close()