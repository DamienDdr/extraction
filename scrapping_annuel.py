import re
import math
from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date, timedelta
import time
import calendar

SESSION_FILE = "bnpparibas_session.json"
OUTPUT_FILE = "leave_planning_2026_complete.csv"


def date_str(d):
    return d.strftime("%Y/%m/%d")


def determine_event_type_and_status(cls):
    """Détermine le type d'événement ET le statut de validation selon les classes CSS"""
    event_type = None
    status = None

    if "grey_cell_weekend" in cls:
        event_type = "JOUR_NON_OUVRE"
        status = None  # Pas de statut pour les jours non ouvrés
    elif "telework" in cls:
        event_type = "TELETRAVAIL"
        if "to_validate_vcell" in cls:
            status = "À valider"
        elif "validated_vcell" in cls:
            status = "Validé"
    elif "validated_vcell" in cls or "to_validate_vcell" in cls:
        event_type = "CONGES"
        if "to_validate_vcell" in cls:
            status = "À valider"
        elif "validated_vcell" in cls:
            status = "Validé"

    return event_type, status


def is_half_day(width_px, col_width):
    """Détecte si c'est une demi-journée (width très petite)"""
    # ✅ CORRECTION : Seuil adaptatif basé sur col_width
    # Une vraie demi-journée fait environ 1-2px
    return width_px <= 2


def pixels_to_days(left_px, width_px, col_width, nb_days, event_type=None):
    """Convertit position/largeur en pixels vers indices de jours"""
    center_px = left_px + width_px / 2
    # ✅ CORRECTION : Utiliser int() au lieu de round() pour éviter le décalage d'un jour
    # round(26.65) = 27 mais on veut 26 (jour 27 en index 0-based)
    center_day = int(center_px / col_width)

    # ✅ CORRECTION : Événements très courts (width < 30% d'un jour)
    # sont traités comme un seul jour basé sur leur position centrale
    # Cela gère les événements isolés avec width:4px ou width:6px généré par DailyRH
    if width_px < col_width * 0.3 or event_type == "JOUR_NON_OUVRE":
        return max(0, min(nb_days - 1, center_day)), max(0, min(nb_days - 1, center_day))

    start_idx = max(0, int(math.floor(left_px / col_width)))
    end_idx = min(nb_days - 1, int(math.ceil((left_px + width_px) / col_width)) - 1)

    return start_idx, end_idx


def build_detail(title, status):
    """Construit le texte du detail en combinant title et status"""
    if title and status:
        return f"{title} ({status})"
    elif title:
        return title
    elif status:
        return status
    else:
        return ""


def scrape_month(page, year, month):
    """Scrape un mois donné et retourne les records"""
    month_start = date(year, month, 1)
    _, last_day = calendar.monthrange(year, month)
    month_end = date(year, month, last_day)
    nb_days = (month_end - month_start).days + 1

    print(f"\n📆 Traitement : {month_start.strftime('%B %Y').upper()}")

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

        if (
            name == "Mes Collègues"
            or (isinstance(name,str) and name.startswith("Signataire"))
            or (isinstance(name,str) and name.startswith("Total"))
            or not name
        ):
            continue

        # ✅ NOUVEAU : Extraire l'UID depuis data-corp-id
        uid = ""
        try:
            # Chercher l'élément avec data-corp-id dans la ligne
            corp_id_elem = row.locator("[data-corp-id]").first
            if corp_id_elem.count() > 0:
                corp_id_full = corp_id_elem.get_attribute("data-corp-id") or ""
                # Extraire la partie numérique (ex: "HRF354710" → "354710")
                uid_match = re.search(r'(\d+)$', corp_id_full)
                if uid_match:
                    uid = uid_match.group(1)
        except Exception as e:
            print(f"   ⚠️  Impossible d'extraire l'UID pour {name}: {e}")
            uid = ""

        planning = {}
        for i in range(nb_days):
            d = month_start + timedelta(days=i)
            planning[i] = {
                "date": date_str(d),
                "type_am": "PRESENT",
                "type_pm": "PRESENT",
                "detail_am": "",
                "detail_pm": ""
            }

        matrix_div = row.locator(".dhx_matrix_line").first
        matrix_width = matrix_div.evaluate("el => el.offsetWidth")
        col_width = matrix_width / nb_days

        events = matrix_div.locator("div[class*='cell'], div[class*='event']")

        # Collecter tous les événements
        all_events = []

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

            # ✅ Détecter type ET statut
            event_type, status = determine_event_type_and_status(cls)
            if not event_type:
                continue

            # ✅ Construire le detail (title + status)
            detail = build_detail(title, status)

            start_idx, end_idx = pixels_to_days(left_px, width_px, col_width, nb_days, event_type)
            half_day = is_half_day(width_px, col_width)

            all_events.append({
                'type': event_type,
                'detail': detail,  # ✅ Detail complet (title + status)
                'start_idx': start_idx,
                'end_idx': end_idx,
                'half_day': half_day,
                'order': e
            })

        # Regrouper les demi-journées par jour
        half_days_by_day = {}
        for evt in all_events:
            if evt['half_day']:
                for day_idx in range(evt['start_idx'], evt['end_idx'] + 1):
                    if day_idx not in half_days_by_day:
                        half_days_by_day[day_idx] = []
                    half_days_by_day[day_idx].append(evt)

        count_jno = 0
        count_teletravail = 0
        count_conges = 0
        count_demi = 0

        # ÉTAPE 1 : Traiter les DEMI-JOURNÉES EN PREMIER
        for day_idx, day_events in half_days_by_day.items():
            day_events.sort(key=lambda x: x['order'])

            for idx, evt in enumerate(day_events[:2]):
                period = "am" if idx == 0 else "pm"
                event_type = evt['type']
                detail = evt['detail']

                count_demi += 1

                if period == "am":
                    current = planning[day_idx]["type_am"]
                    if event_type == "JOUR_NON_OUVRE" or (event_type == "CONGES" and current != "JOUR_NON_OUVRE") or (
                            event_type == "TELETRAVAIL" and current == "PRESENT"):
                        planning[day_idx]["type_am"] = event_type
                        planning[day_idx]["detail_am"] = detail
                else:
                    current = planning[day_idx]["type_pm"]
                    if event_type == "JOUR_NON_OUVRE" or (event_type == "CONGES" and current != "JOUR_NON_OUVRE") or (
                            event_type == "TELETRAVAIL" and current == "PRESENT"):
                        planning[day_idx]["type_pm"] = event_type
                        planning[day_idx]["detail_pm"] = detail

        # ÉTAPE 2 : Traiter les JOURNÉES ENTIÈRES
        for evt in all_events:
            event_type = evt['type']
            detail = evt['detail']

            if event_type == "JOUR_NON_OUVRE":
                count_jno += 1
            elif event_type == "TELETRAVAIL":
                count_teletravail += 1
            elif event_type == "CONGES":
                count_conges += 1

            if evt['half_day']:
                continue

            for day_idx in range(evt['start_idx'], evt['end_idx'] + 1):
                current_am = planning[day_idx]["type_am"]
                current_pm = planning[day_idx]["type_pm"]

                if event_type == "JOUR_NON_OUVRE":
                    planning[day_idx]["type_am"] = "JOUR_NON_OUVRE"
                    planning[day_idx]["type_pm"] = "JOUR_NON_OUVRE"
                    planning[day_idx]["detail_am"] = detail
                    planning[day_idx]["detail_pm"] = detail

                elif event_type == "CONGES":
                    # ✅ CORRECTION CRITIQUE : Toujours appliquer CONGES sauf si JOUR_NON_OUVRE
                    # Cela permet de gérer les saisies multiples (ex: RTT 27-29 puis RTT 30)
                    # CONGES a priorité sur PRESENT ET TELETRAVAIL
                    if current_am != "JOUR_NON_OUVRE":
                        planning[day_idx]["type_am"] = "CONGES"
                        planning[day_idx]["detail_am"] = detail
                    if current_pm != "JOUR_NON_OUVRE":
                        planning[day_idx]["type_pm"] = "CONGES"
                        planning[day_idx]["detail_pm"] = detail

                elif event_type == "TELETRAVAIL":
                    # Télétravail seulement si PRESENT (ne remplace pas CONGES)
                    if current_am == "PRESENT":
                        planning[day_idx]["type_am"] = "TELETRAVAIL"
                        planning[day_idx]["detail_am"] = detail
                    if current_pm == "PRESENT":
                        planning[day_idx]["type_pm"] = "TELETRAVAIL"
                        planning[day_idx]["detail_pm"] = detail

        print(f"   ✅ JNO: {count_jno} | 🏠 TT: {count_teletravail} | 🏖️ CP: {count_conges} | ⏰ Demi: {count_demi}")

        # Export
        for i in range(nb_days):
            info = planning[i]
            records.append({
                "collaborateur": name,
                "uid": uid,  # ✅ NOUVEAU : Ajout de l'UID
                "date": info["date"],
                "type_am": info["type_am"],
                "detail_am": info["detail_am"],
                "type_pm": info["type_pm"],
                "detail_pm": info["detail_pm"]
            })

    print(f"   ✅ {len(records)} lignes extraites")
    return records


def get_current_month_text(page):
    """Récupère le texte du mois affiché - ex: 'février 2026'"""
    try:
        date_elem = page.locator("#date_now")
        date_elem.wait_for(state="attached", timeout=15000)
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

    for month_name, month_num in french_months.items():
        if month_name in text_lower:
            year_match = re.search(r'(\d{4})', text)
            if year_match:
                year = int(year_match.group(1))
                return month_num, year

    return None, None


def navigate_to_january(page, year):
    """Revient à janvier de l'année donnée avec détection intelligente"""
    target_month = 1
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

        if current_year > target_year or (current_year == target_year and current_month > target_month):
            print(f"   ◀️  Clic {clicks + 1} : {current_text} → précédent")
            prev_button = page.locator("div.dhx_cal_prev_button.prev-month").first
            prev_button.click()
        elif current_year < target_year or (current_year == target_year and current_month < target_month):
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

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/teamplanning")
    page.wait_for_load_state("networkidle")

    print("✅ Page chargée, attente du chargement complet...")
    time.sleep(10)

    print("\n🔍 DEBUG : Test de détection du mois...")
    try:
        detected_month = get_current_month_text(page)
        print(f"   ✅ Mois détecté : '{detected_month}'")
    except Exception as e:
        print(f"   ❌ Échec détection : {e}")
        browser.close()
        exit(1)

    try:
        navigate_to_january(page, 2026)
    except Exception as e:
        print(f"❌ Impossible d'aller à janvier : {e}")
        import traceback

        traceback.print_exc()
        browser.close()
        exit(1)

    all_records = []

    for month in range(1, 13):
        try:
            month_records = scrape_month(page, 2026, month)
            all_records.extend(month_records)

            if month < 12:
                go_to_next_month(page)

        except Exception as e:
            print(f"❌ Erreur pour {calendar.month_name[month]} 2026 : {e}")
            import traceback

            traceback.print_exc()

            if month < 12:
                try:
                    go_to_next_month(page)
                except:
                    print("⚠️ Impossible d'avancer au mois suivant, arrêt du script")
                    break
            continue

    if all_records:
        df = pd.DataFrame(all_records)
        df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8")
        print(f"\n🎉 Export terminé → {OUTPUT_FILE}")
        print(f"📊 Total : {len(all_records)} lignes")

        months_scraped = sorted(set(r['date'][:7] for r in all_records))
        print(f"📅 Mois collectés : {len(months_scraped)}/12")
        for month_str in months_scraped:
            month_data = [r for r in all_records if r['date'][:7] == month_str]
            print(f"   • {month_str} : {len(month_data)} lignes")
    else:
        print("\n❌ Aucune donnée collectée")

    browser.close()