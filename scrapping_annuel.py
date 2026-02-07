import re
import math
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date, timedelta
import time
import calendar

SESSION_FILE = "bnpparibas_session.json"
OUTPUT_FILE = "leave_planning_2026_complete.csv"


def is_weekend(d):
    return d.weekday() >= 5


def date_str(d):
    return d.strftime("%Y/%m/%d")


def determine_event_type(cls):
    if "grey_cell_weekend" in cls:
        return None
    if "telework" in cls and "validated_vcell" in cls:
        return "TELETRAVAIL"
    elif "validated_vcell" in cls:
        return "CONGES"
    return None


def pixels_to_days(left_px, width_px, col_width, nb_days):
    center_px = left_px + width_px / 2
    center_day = round(center_px / col_width)

    if width_px < col_width * 0.8:
        return max(0, min(nb_days - 1, center_day)), max(0, min(nb_days - 1, center_day))

    start_idx = max(0, int(math.floor(left_px / col_width)))
    end_idx = min(nb_days - 1, int(math.ceil((left_px + width_px) / col_width)) - 1)

    return start_idx, end_idx


def scrape_month(page, year, month):
    """Scrape un mois donné et retourne les records"""
    month_start = date(year, month, 1)
    _, last_day = calendar.monthrange(year, month)
    month_end = date(year, month, last_day)
    nb_days = (month_end - month_start).days + 1

    print(f"\n📆 Traitement : {month_start.strftime('%B %Y').upper()}")

    # Attendre que le calendrier soit chargé
    time.sleep(2)

    rows = page.locator("tr.dhx_row_item")
    row_count = rows.count()

    if row_count == 0:
        print(f"   ⚠️ Aucune ligne détectée pour {month_start.strftime('%B %Y')}")
        return []

    print(f"   👥 {row_count} collaborateurs")

    records = []

    for r in range(row_count):
        row = rows.nth(r)

        name_cell = row.locator("td.dhx_matrix_scell").first
        name = name_cell.inner_text().strip() if name_cell.count() > 0 else "INCONNU"

        if name == "Mes Collègues" or not name:
            continue

        # INIT planning du mois
        planning = {}
        for i in range(nb_days):
            d = month_start + timedelta(days=i)
            planning[i] = {
                "date": date_str(d),
                "type": "WEEKEND" if is_weekend(d) else "PRESENT",
                "title": ""
            }

        matrix_div = row.locator(".dhx_matrix_line").first
        matrix_width = matrix_div.evaluate("el => el.offsetWidth")
        col_width = matrix_width / nb_days

        events = matrix_div.locator("div[class*='cell'], div[class*='event']")

        for e in range(events.count()):
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

            start_idx, end_idx = pixels_to_days(left_px, width_px, col_width, nb_days)

            for day_idx in range(start_idx, end_idx + 1):
                if planning[day_idx]["type"] == "WEEKEND":
                    continue

                if event_type == "CONGES" or planning[day_idx]["type"] == "PRESENT":
                    planning[day_idx]["type"] = event_type
                    planning[day_idx]["title"] = title

        # Ajouter aux records
        for i in range(nb_days):
            info = planning[i]
            records.append({
                "collaborateur": name,
                "date": info["date"],
                "type": info["type"],
                "title": info["title"]
            })

    print(f"   ✅ {len(records)} lignes extraites")
    return records


def get_current_month_text(page):
    """Récupère le texte du mois affiché - ex: 'février 2026'"""
    try:
        date_elem = page.locator("#date_now")

        # ✅ Attendre que l'élément soit ATTACHÉ au DOM (pas forcément visible)
        date_elem.wait_for(state="attached", timeout=15000)

        # text_content() fonctionne même sur les éléments cachés
        text = date_elem.text_content(timeout=5000)

        if text and text.strip():
            return text.strip()

        raise RuntimeError("Le texte de #date_now est vide")

    except Exception as e:
        raise RuntimeError(f"❌ Impossible de trouver le texte du mois (#date_now) : {e}")


def parse_month_year(text):
    """Parse 'janvier 2026' → (1, 2026) ou 'février 2026' → (2, 2026)"""
    french_months = {
        "janvier": 1, "février": 2, "mars": 3, "avril": 4,
        "mai": 5, "juin": 6, "juillet": 7, "août": 8,
        "septembre": 9, "octobre": 10, "novembre": 11, "décembre": 12
    }

    text_lower = text.lower()

    # Extraire mois et année
    for month_name, month_num in french_months.items():
        if month_name in text_lower:
            year_match = re.search(r'(\d{4})', text)
            if year_match:
                year = int(year_match.group(1))
                return month_num, year

    return None, None


def navigate_to_january(page, year):
    """Revient à janvier de l'année donnée avec détection intelligente"""
    target_month = 1  # Janvier
    target_year = year

    current_text = get_current_month_text(page)
    current_month, current_year = parse_month_year(current_text)

    print(f"\n🎯 NAVIGATION VERS JANVIER {year}")
    print(f"   📍 Position actuelle : {current_text} (mois {current_month}, année {current_year})")

    if current_month is None or current_year is None:
        raise RuntimeError(f"❌ Impossible de parser le mois actuel : {current_text}")

    max_clicks = 50
    clicks = 0

    while (current_month != target_month or current_year != target_year) and clicks < max_clicks:

        # Décider d'aller en arrière ou en avant
        if current_year > target_year or (current_year == target_year and current_month > target_month):
            # Aller en arrière ◀️
            print(f"   ◀️  Clic {clicks + 1} : {current_text} → précédent")
            prev_button = page.locator("div.dhx_cal_prev_button.prev-month").first
            prev_button.click()
        elif current_year < target_year or (current_year == target_year and current_month < target_month):
            # Aller en avant ▶️
            print(f"   ▶️  Clic {clicks + 1} : {current_text} → suivant")
            next_button = page.locator("div.dhx_cal_next_button.next-month").first
            next_button.click()

        time.sleep(1.5)

        current_text = get_current_month_text(page)
        current_month, current_year = parse_month_year(current_text)
        clicks += 1

    if clicks >= max_clicks:
        raise RuntimeError(f"❌ Impossible d'atteindre Janvier {year} après {clicks} clics")

    print(f"   ✅ Navigation réussie en {clicks} clics → {current_text}\n")


def go_to_next_month(page):
    """Avance d'un mois (flèche droite)"""
    next_button = page.locator("div.dhx_cal_next_button.next-month").first
    next_button.click()
    time.sleep(1.5)


# ============== SCRIPT PRINCIPAL ==============

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state=SESSION_FILE)
    page = context.new_page()

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")
    page.wait_for_load_state("networkidle")

    print("✅ Page chargée, attente du chargement complet...")
    time.sleep(6)

    # 🔍 DEBUG : Afficher le texte du mois détecté
    print("\n🔍 DEBUG : Test de détection du mois...")
    try:
        detected_month = get_current_month_text(page)
        print(f"   ✅ Mois détecté : '{detected_month}'")
    except Exception as e:
        print(f"   ❌ Échec détection : {e}")
        browser.close()
        exit(1)

    # ⏪ ÉTAPE 1 : Aller à janvier 2026
    try:
        navigate_to_january(page, 2026)
    except Exception as e:
        print(f"❌ Impossible d'aller à janvier : {e}")
        import traceback

        traceback.print_exc()
        browser.close()
        exit(1)

    all_records = []

    # ▶️ ÉTAPE 2 : Boucle sur les 12 mois de 2026
    for month in range(1, 13):
        try:
            # Scraper le mois actuel
            month_records = scrape_month(page, 2026, month)
            all_records.extend(month_records)

            # Aller au mois suivant (sauf pour décembre)
            if month < 12:
                go_to_next_month(page)

        except Exception as e:
            print(f"❌ Erreur pour {calendar.month_name[month]} 2026 : {e}")
            import traceback

            traceback.print_exc()

            # Essayer de continuer quand même
            if month < 12:
                try:
                    go_to_next_month(page)
                except:
                    print("⚠️ Impossible d'avancer au mois suivant, arrêt du script")
                    break
            continue

    # Export final
    if all_records:
        df = pd.DataFrame(all_records)
        df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8")
        print(f"\n🎉 Export terminé → {OUTPUT_FILE}")
        print(f"📊 Total : {len(all_records)} lignes")

        # Statistiques par mois
        months_scraped = sorted(set(r['date'][:7] for r in all_records))
        print(f"📅 Mois collectés : {len(months_scraped)}/12")
        for month_str in months_scraped:
            month_data = [r for r in all_records if r['date'][:7] == month_str]
            print(f"   • {month_str} : {len(month_data)} lignes")
    else:
        print("\n❌ Aucune donnée collectée")

    browser.close()