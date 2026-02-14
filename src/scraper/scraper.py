"""Module de scraping des données de planning DailyRH"""

import re
import math
import time
import calendar
from datetime import date, timedelta
from typing import List, Dict, Tuple, Set, Optional
from playwright.sync_api import Page, sync_playwright

from src.config import (
    SESSION_FILE, DAILYRH_URL, TARGET_YEAR,
    HEADLESS_MODE, NAVIGATION_DELAY, INITIAL_LOAD_DELAY, MAX_NAVIGATION_CLICKS
)
from src.utils import (
    date_to_string, build_detail, extract_uid_from_corp_id,
    extract_date_from_css_class, parse_month_year_text
)
from src.logging import get_logger

logger = get_logger()


def determine_event_type_and_status(css_class: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Détermine le type d'événement et le statut de validation depuis les classes CSS.
    
    Args:
        css_class: Classes CSS de l'élément
        
    Returns:
        Tuple (type_événement, statut)
    """
    event_type = None
    status = None
    
    if "grey_cell_weekend" in css_class:
        event_type = "JOUR_NON_OUVRE"
        status = None
    elif "telework" in css_class:
        event_type = "TELETRAVAIL"
        if "to_validate_vcell" in css_class:
            status = "À valider"
        elif "validated_vcell" in css_class:
            status = "Validé"
    elif "validated_vcell" in css_class or "to_validate_vcell" in css_class:
        event_type = "CONGES"
        if "to_validate_vcell" in css_class:
            status = "À valider"
        elif "validated_vcell" in css_class:
            status = "Validé"
    
    return event_type, status


def is_half_day(width_px: float, col_width: float) -> bool:
    """
    Détecte si un événement est une demi-journée basé sur sa largeur.
    
    Args:
        width_px: Largeur de l'événement en pixels
        col_width: Largeur d'une colonne (jour) en pixels
        
    Returns:
        True si demi-journée
    """
    return width_px <= 2


def pixels_to_days(left_px: float, width_px: float, col_width: float, nb_days: int) -> Tuple[int, int]:
    """
    Convertit position/largeur en pixels vers indices de jours.
    
    Args:
        left_px: Position left en pixels
        width_px: Largeur en pixels
        col_width: Largeur d'une colonne en pixels
        nb_days: Nombre de jours dans le mois
        
    Returns:
        Tuple (start_idx, end_idx) (0-based)
    """
    center_px = left_px + width_px / 2
    center_day = int(center_px / col_width)
    
    # Événements très courts
    if width_px < col_width * 0.3:
        return max(0, min(nb_days - 1, center_day)), max(0, min(nb_days - 1, center_day))
    
    start_idx = max(0, int(math.floor(left_px / col_width)))
    end_idx = min(nb_days - 1, int((left_px + width_px - col_width / 2) / col_width))
    
    return start_idx, end_idx


def extract_non_working_days(page: Page, year: int, month: int) -> Set[int]:
    """
    Extrait les indices des jours non ouvrés pour un mois donné.
    
    Args:
        page: Page Playwright
        year: Année
        month: Mois (1-12)
        
    Returns:
        Set d'indices de jours (0-based)
    """
    jno_dates_set = set()
    all_jno_timespans = page.locator("div.dhx_marked_timespan.grey_cell_weekend")
    
    for i in range(all_jno_timespans.count()):
        css_class = all_jno_timespans.nth(i).get_attribute("class") or ""
        jno_date = extract_date_from_css_class(css_class)
        if jno_date:
            jno_dates_set.add(jno_date)
    
    # Convertir en indices pour le mois courant
    jno_day_indices = set()
    for d_str in jno_dates_set:
        parts = d_str.split("/")
        if len(parts) == 3:
            d_year, d_month, d_day = int(parts[0]), int(parts[1]), int(parts[2])
            if d_year == year and d_month == month:
                jno_day_indices.add(d_day - 1)  # 0-based
    
    return jno_day_indices


def extract_collaborator_events(row, nb_days: int) -> List[Dict]:
    """
    Version robuste basée sur bounding_box().
    """
    matrix_div = row.locator(".dhx_matrix_line").first
    cells = row.locator("td.dhx_matrix_cell")

    # Pré-calcul des bounding boxes des cellules
    day_boxes = []
    for i in range(min(nb_days, cells.count())):
        box = cells.nth(i).bounding_box()
        if box:
            day_boxes.append(box)
        else:
            day_boxes.append(None)

    events_normal = matrix_div.locator(
        "div[class*='cell']:not(.dhx_marked_timespan), "
        "div[class*='event']:not(.dhx_marked_timespan)"
    )

    all_events = []

    for i in range(events_normal.count()):
        ev = events_normal.nth(i)
        css_class = ev.get_attribute("class") or ""
        title = ev.get_attribute("title") or ""

        if "grey_cell_weekend" in css_class:
            continue

        event_box = ev.bounding_box()
        if not event_box:
            continue

        event_type, status = determine_event_type_and_status(css_class)
        if not event_type or event_type == "JOUR_NON_OUVRE":
            continue

        detail = build_detail(title, status)

        impacted_days = []

        for day_idx, cell_box in enumerate(day_boxes):
            if not cell_box:
                continue

            # Vérifier chevauchement horizontal
            overlap = not (
                event_box["x"] + event_box["width"] <= cell_box["x"]
                or event_box["x"] >= cell_box["x"] + cell_box["width"]
            )

            if overlap:
                impacted_days.append(day_idx)

        if not impacted_days:
            continue

        start_idx = min(impacted_days)
        end_idx = max(impacted_days)

        # Détection demi-journée
        first_cell_box = day_boxes[start_idx]
        cell_width = first_cell_box["width"]
        ratio = event_box["width"] / cell_width

        half_day = ratio < 0.65

        period = None
        if half_day:
            event_center = event_box["x"] + event_box["width"] / 2
            cell_center = first_cell_box["x"] + cell_width / 2
            period = "am" if event_center < cell_center else "pm"

        all_events.append({
            "type": event_type,
            "detail": detail,
            "start_idx": start_idx,
            "end_idx": end_idx,
            "half_day": half_day,
            "period": period,
            "order": len(all_events),
        })

    return all_events


def apply_half_day_events(planning: Dict, events: List[Dict]):
    """
    Applique les événements de demi-journée au planning.

    Args:
        planning: Dictionnaire de planning à modifier
        events: Liste d'événements
    """
    # Grouper les demi-journées par jour
    half_days_by_day = {}
    for evt in events:
        if evt['half_day']:
            for day_idx in range(evt['start_idx'], evt['end_idx'] + 1):
                if day_idx not in half_days_by_day:
                    half_days_by_day[day_idx] = []
                half_days_by_day[day_idx].append(evt)

    # Appliquer les demi-journées
    for day_idx, day_events in half_days_by_day.items():
        for evt in day_events:
            period = evt.get('period', 'am')  # Par défaut AM si pas détecté
            event_type = evt['type']
            detail = evt['detail']

            if period == "am":
                current = planning[day_idx]["type_am"]
                if (event_type == "CONGES" and current in ["PRESENT", "TELETRAVAIL"]) or \
                   (event_type == "TELETRAVAIL" and current == "PRESENT"):
                    planning[day_idx]["type_am"] = event_type
                    planning[day_idx]["detail_am"] = detail
            else:  # period == "pm"
                current = planning[day_idx]["type_pm"]
                if (event_type == "CONGES" and current in ["PRESENT", "TELETRAVAIL"]) or \
                   (event_type == "TELETRAVAIL" and current == "PRESENT"):
                    planning[day_idx]["type_pm"] = event_type
                    planning[day_idx]["detail_pm"] = detail


def apply_full_day_events(planning: Dict, events: List[Dict]):
    """
    Applique les événements de journée entière au planning.
    
    Args:
        planning: Dictionnaire de planning à modifier
        events: Liste d'événements
    """
    # D'abord les CONGES
    for evt in events:
        if evt['half_day'] or evt['type'] != 'CONGES':
            continue
        detail = evt['detail']
        for day_idx in range(evt['start_idx'], evt['end_idx'] + 1):
            current_am = planning[day_idx]["type_am"]
            current_pm = planning[day_idx]["type_pm"]
            if current_am in ["PRESENT", "TELETRAVAIL"]:
                planning[day_idx]["type_am"] = "CONGES"
                planning[day_idx]["detail_am"] = detail
            if current_pm in ["PRESENT", "TELETRAVAIL"]:
                planning[day_idx]["type_pm"] = "CONGES"
                planning[day_idx]["detail_pm"] = detail
    
    # Ensuite TELETRAVAIL
    for evt in events:
        if evt['half_day'] or evt['type'] != 'TELETRAVAIL':
            continue
        detail = evt['detail']
        for day_idx in range(evt['start_idx'], evt['end_idx'] + 1):
            current_am = planning[day_idx]["type_am"]
            current_pm = planning[day_idx]["type_pm"]
            if current_am == "PRESENT":
                planning[day_idx]["type_am"] = "TELETRAVAIL"
                planning[day_idx]["detail_am"] = detail
            if current_pm == "PRESENT":
                planning[day_idx]["type_pm"] = "TELETRAVAIL"
                planning[day_idx]["detail_pm"] = detail


def apply_non_working_days(planning: Dict, jno_indices: Set[int]):
    """
    Applique les jours non ouvrés (priorité absolue).
    
    Args:
        planning: Dictionnaire de planning à modifier
        jno_indices: Set d'indices de jours non ouvrés
    """
    for day_idx in jno_indices:
        if day_idx in planning:
            planning[day_idx]["type_am"] = "JOUR_NON_OUVRE"
            planning[day_idx]["detail_am"] = ""
            planning[day_idx]["type_pm"] = "JOUR_NON_OUVRE"
            planning[day_idx]["detail_pm"] = ""


def scrape_month(page: Page, year: int, month: int) -> List[Dict]:
    """
    Scrape les données d'un mois donné (version robuste).
    """
    month_start = date(year, month, 1)
    _, last_day = calendar.monthrange(year, month)
    month_end = date(year, month, last_day)
    nb_days = (month_end - month_start).days + 1

    logger.info(f"Traitement du mois : {month_start.strftime('%B %Y')}")

    # Attendre que les lignes soient présentes
    page.wait_for_selector("tr.dhx_row_item", timeout=15000)

    rows = page.locator("tr.dhx_row_item")
    row_count = rows.count()

    if row_count == 0:
        logger.warning(f"Aucune ligne détectée pour {month_start.strftime('%B %Y')}")
        return []

    logger.info(f"Nombre de collaborateurs : {row_count}")

    # Extraire les jours non ouvrés
    jno_day_indices = extract_non_working_days(page, year, month)
    logger.info(f"Jours non ouvrés : {len(jno_day_indices)} jours")

    records = []

    for r in range(row_count):
        row = rows.nth(r)

        name_cell = row.locator("td.dhx_matrix_scell").first
        name = name_cell.inner_text().strip() if name_cell.count() > 0 else "INCONNU"


        if (
                name == "Mes Collègues"
                or (isinstance(name, str) and name.startswith("Signataire"))
                or (isinstance(name, str) and name.startswith("Total"))
                or not name
        ):
            continue

        # Extraction UID
        uid = ""
        try:
            corp_id_elem = row.locator("[data-corp-id]").first
            if corp_id_elem.count() > 0:
                corp_id_full = corp_id_elem.get_attribute("data-corp-id") or ""
                uid = extract_uid_from_corp_id(corp_id_full)
        except Exception as e:
            logger.warning(f"Impossible d'extraire l'UID pour {name}: {e}")

        # Initialisation planning
        planning = {}
        for i in range(nb_days):
            d = month_start + timedelta(days=i)
            planning[i] = {
                "date": date_to_string(d),
                "type_am": "PRESENT",
                "type_pm": "PRESENT",
                "detail_am": "",
                "detail_pm": ""
            }

        # Vérifier que la matrice est bien rendue
        matrix_div = row.locator(".dhx_matrix_line").first
        matrix_div.wait_for(state="attached", timeout=10000)

        # Extraction événements
        try:
            events = extract_collaborator_events(row, nb_days)
        except Exception as e:
            logger.error(f"Erreur extraction événements pour {name}: {e}")
            continue

        # Application logique métier
        try:
            apply_half_day_events(planning, events)
            apply_full_day_events(planning, events)
            apply_non_working_days(planning, jno_day_indices)
        except Exception as e:
            logger.error(f"Erreur application planning pour {name}: {e}")
            continue

        logger.debug(f"Traité : {name} ({uid})")

        # Génération des records
        for i in range(nb_days):
            info = planning[i]
            records.append({
                "collaborateur": name,
                "uid": uid,
                "date": info["date"],
                "type_am": info["type_am"],
                "detail_am": info["detail_am"],
                "type_pm": info["type_pm"],
                "detail_pm": info["detail_pm"]
            })

    logger.info(f"Lignes extraites : {len(records)}")

    # ← NOUVELLE APPROCHE : Extraction des totaux depuis les cellules
    try:
        dailyrh_totals = extract_dailyrh_totals(page, year, month, nb_days)
        logger.info(f"Totaux DailyRH extraits : {len(dailyrh_totals)} jours")

        validation = validate_totals(dailyrh_totals, records, jno_day_indices, year, month)

        if validation['errors_count'] == 0:
            logger.info(f"✅ Validation OK : aucun écart détecté")
        else:
            logger.warning(f"⚠️ Validation : {validation['errors_count']} écarts détectés")
            for error in validation['errors'][:10]:
                logger.warning(
                    f"  Jour {error['day']} : "
                    f"DailyRH={error['dailyrh_count']:.1f}, Scrapé={error['scraped_count']:.1f} "
                    f"(écart: {error['difference']:+.1f})"
                )
            if validation['errors_count'] > 10:
                logger.warning(f"  ... et {validation['errors_count'] - 10} autres écarts")
    except Exception as e:
        logger.error(f"Erreur lors de la validation des totaux : {e}")

    return records

def get_current_month_text(page: Page) -> str:
    """
    Récupère le texte du mois affiché.
    
    Args:
        page: Page Playwright
        
    Returns:
        Texte du mois (ex: 'février 2026')
    """
    try:
        date_elem = page.locator("#date_now")
        date_elem.wait_for(state="attached", timeout=15000)
        text = date_elem.text_content(timeout=5000)
        
        if text and text.strip():
            return text.strip()
        
        raise RuntimeError("Le texte de #date_now est vide")
    
    except Exception as e:
        raise RuntimeError(f"Impossible de trouver le texte du mois (#date_now) : {e}")


def navigate_to_january(page: Page, year: int):
    """
    Navigue vers janvier de l'année cible.
    
    Args:
        page: Page Playwright
        year: Année cible
    """
    target_month = 1
    
    current_text = get_current_month_text(page)
    current_month, current_year = parse_month_year_text(current_text)
    
    logger.info(f"Navigation vers janvier {year}")
    logger.info(f"Position actuelle : {current_text}")
    
    if current_month is None or current_year is None:
        raise RuntimeError(f"Impossible de parser le mois actuel : {current_text}")
    
    clicks = 0
    
    while (current_month != target_month or current_year != year) and clicks < MAX_NAVIGATION_CLICKS:
        
        if current_year > year or (current_year == year and current_month > target_month):
            prev_button = page.locator("div.dhx_cal_prev_button.prev-month").first
            prev_button.click()
        elif current_year < year or (current_year == year and current_month < target_month):
            next_button = page.locator("div.dhx_cal_next_button.next-month").first
            next_button.click()
        
        time.sleep(NAVIGATION_DELAY)
        
        current_text = get_current_month_text(page)
        current_month, current_year = parse_month_year_text(current_text)
        clicks += 1
    
    if clicks >= MAX_NAVIGATION_CLICKS:
        raise RuntimeError(f"Impossible d'atteindre janvier {year} après {clicks} clics")
    
    logger.info(f"Navigation réussie en {clicks} clics")


def go_to_next_month(page: Page):
    """
    Avance d'un mois.
    
    Args:
        page: Page Playwright
    """
    next_button = page.locator("div.dhx_cal_next_button.next-month").first
    next_button.click()
    time.sleep(NAVIGATION_DELAY)


def scrape_all_months(year: int) -> List[Dict]:
    """
    Scrape tous les mois de l'année.
    
    Args:
        year: Année à scraper
        
    Returns:
        Liste de tous les records
    """
    all_records = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS_MODE)
        context = browser.new_context(storage_state=SESSION_FILE)
        page = context.new_page()
        
        logger.info("Chargement de DailyRH...")
        page.goto(DAILYRH_URL)
        page.wait_for_load_state("networkidle")
        
        logger.info(f"Attente du chargement complet ({INITIAL_LOAD_DELAY}s)...")
        time.sleep(INITIAL_LOAD_DELAY)
        
        # Navigation vers janvier
        try:
            navigate_to_january(page, year)
        except Exception as e:
            logger.error(f"Impossible de naviguer vers janvier : {e}")
            browser.close()
            raise
        
        # Scraper chaque mois
        for month in range(1, 13):
            try:
                month_records = scrape_month(page, year, month)
                all_records.extend(month_records)
                
                if month < 12:
                    go_to_next_month(page)
            
            except Exception as e:
                logger.error(f"Erreur pour {calendar.month_name[month]} {year} : {e}")
                
                if month < 12:
                    try:
                        go_to_next_month(page)
                    except:
                        logger.error("Impossible d'avancer au mois suivant, arrêt du script")
                        break
                continue
        
        browser.close()
    
    return all_records


def validate_totals(dailyrh_totals: Dict[str, float], scraped_data: List[Dict], jno_indices: Set[int], year: int,
                       month: int) -> Dict:
    """
    Compare les totaux DailyRH (chiffres bruts) avec les données scrapées.
    EXCLUT les jours non ouvrés.
    """
    errors = []

    # Calculer les totaux scrapés par jour
    scraped_by_day = {}
    for record in scraped_data:
        date_str = record['date']
        if date_str not in scraped_by_day:
            scraped_by_day[date_str] = 0.0

        type_am = record['type_am']
        type_pm = record['type_pm']

        # Compter les événements (pas PRESENT ni JOUR_NON_OUVRE)
        if type_am not in ['PRESENT', 'JOUR_NON_OUVRE']:
            scraped_by_day[date_str] += 0.5
        if type_pm not in ['PRESENT', 'JOUR_NON_OUVRE']:
            scraped_by_day[date_str] += 0.5

    # Comparer jour par jour
    for date_str, dr_count in dailyrh_totals.items():
        # Ignorer les jours non ouvrés
        day_idx = int(date_str.split('/')[2]) - 1  # 0-based
        if day_idx in jno_indices:
            continue

        sc_count = scraped_by_day.get(date_str, 0.0)

        if abs(dr_count - sc_count) > 0.01:  # Tolérance pour les flottants
            day_num = int(date_str.split('/')[2])
            errors.append({
                'date': date_str,
                'day': day_num,
                'dailyrh_count': dr_count,
                'scraped_count': sc_count,
                'difference': sc_count - dr_count
            })

    return {
        'errors_count': len(errors),
        'errors': sorted(errors, key=lambda x: x['day'])
    }


def extract_dailyrh_totals(page: Page, year: int, month: int, nb_days: int) -> Dict[str, float]:
    """
    Extrait les totaux DailyRH depuis les cellules teamTotal_cell.

    Returns:
        Dictionnaire {date_string: nombre_événements}
    """
    totals = {}

    # Sélectionner toutes les cellules de total
    total_cells = page.locator("td.teamTotal_cell")

    for i in range(total_cells.count()):
        cell = total_cells.nth(i)
        css_class = cell.get_attribute("class") or ""

        # Extraire la date depuis la classe (format : 2026-02-01)
        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', css_class)
        if not date_match:
            continue

        date_str = date_match.group(1).replace('-', '/')  # 2026/02/01

        # Vérifier que c'est bien le mois en cours
        if not date_str.startswith(f"{year}/{month:02d}"):
            continue

        # Extraire le nombre depuis le <div> intérieur
        try:
            text = cell.inner_text().strip()
            # Remplacer la virgule par un point pour les décimales
            count = float(text.replace(',', '.')) if text else 0.0
            totals[date_str] = count
        except:
            totals[date_str] = 0.0

    return totals
