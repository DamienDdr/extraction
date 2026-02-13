"""
Module de fonctions utilitaires

Ce module contient toutes les fonctions réutilisables pour le traitement
des données de planning : parsing de dates, validation des événements,
génération de codes de statut, etc.

Les fonctions sont organisées par thématique :
- Conversion de dates
- Validation et parsing des détails d'événements
- Génération de codes de statut
- Extraction d'informations depuis le HTML/CSS
- Comptage d'événements
"""

import re
from datetime import date
from typing import Optional, Tuple


# ============================================================
# CONVERSION DE DATES
# ============================================================

def date_to_string(d: date) -> str:
    """
    Convertit une date Python en string au format YYYY/MM/DD.
    
    Args:
        d: Date à convertir
        
    Returns:
        String au format "2026/01/15"
        
    Exemple:
        >>> from datetime import date
        >>> date_to_string(date(2026, 1, 15))
        '2026/01/15'
    """
    return d.strftime("%Y/%m/%d")


# ============================================================
# VALIDATION ET PARSING DES DÉTAILS
# ============================================================

def is_validated(detail: str) -> Optional[bool]:
    """
    Détermine si un événement est validé à partir du texte de détail.
    
    Cette fonction analyse le texte pour détecter les mentions de validation.
    Elle retourne True si validé, False si à valider, None si indéterminé.
    
    Args:
        detail: Texte de détail de l'événement (ex: "Congés (validé)")
        
    Returns:
        True si validé, False si à valider, None si indéterminé
        
    Exemples:
        >>> is_validated("Congés (validé)")
        True
        >>> is_validated("RTT (à valider)")
        False
        >>> is_validated("")
        None
    """
    if not detail or detail == '':
        return None
    
    detail_lower = str(detail).lower()
    
    if 'validé' in detail_lower or '(validé)' in detail_lower:
        return True
    elif 'à valider' in detail_lower or 'a valider' in detail_lower:
        return False
    
    return None


def is_rtt(detail: str) -> bool:
    """
    Détermine si un congé est un RTT.
    
    Args:
        detail: Texte de détail de l'événement
        
    Returns:
        True si l'événement contient "rtt", False sinon
        
    Exemples:
        >>> is_rtt("RTT (validé)")
        True
        >>> is_rtt("Congés payés")
        False
    """
    if not detail or detail == '':
        return False
    
    return 'rtt' in str(detail).lower()


def build_detail(title: str, status: str) -> str:
    """
    Construit le texte de détail en combinant titre et statut.
    
    Args:
        title: Titre de l'événement (ex: "Congés")
        status: Statut de validation (ex: "Validé")
        
    Returns:
        Texte combiné (ex: "Congés (Validé)")
        
    Exemples:
        >>> build_detail("Congés", "Validé")
        'Congés (Validé)'
        >>> build_detail("RTT", "")
        'RTT'
        >>> build_detail("", "Validé")
        'Validé'
    """
    if title and status:
        return f"{title} ({status})"
    elif title:
        return title
    elif status:
        return status
    else:
        return ""


# ============================================================
# EXTRACTION D'INFORMATIONS DEPUIS HTML/CSS
# ============================================================

def extract_uid_from_corp_id(corp_id: str) -> str:
    """
    Extrait l'UID depuis un attribut data-corp-id.
    
    L'UID est un identifiant de 6 caractères alphanumérique extrait
    depuis le format "HRF344256-0_HRF460606".
    
    Args:
        corp_id: Valeur de data-corp-id (ex: "HRF344256-0_HRF460606")
        
    Returns:
        UID extrait (ex: "460606") ou chaîne vide si non trouvé
        
    Exemple:
        >>> extract_uid_from_corp_id("HRF344256-0_HRF460606")
        '460606'
    """
    if not corp_id:
        return ""
    
    uid_match = re.search(r'HRF([A-Za-z0-9]{6})', corp_id)
    if uid_match:
        return uid_match.group(1)
    
    return ""


def extract_date_from_css_class(css_class: str) -> Optional[str]:
    """
    Extrait la date depuis une classe CSS de dhx_marked_timespan.
    
    Les jours non ouvrés ont leur date encodée dans la classe CSS,
    au format YYYY/MM/DD.
    
    Args:
        css_class: Classe CSS contenant la date
        
    Returns:
        Date au format YYYY/MM/DD ou None si non trouvée
        
    Exemple:
        >>> extract_date_from_css_class("dhx_marked_timespan grey_cell_weekend 2026/04/06")
        '2026/04/06'
    """
    if not css_class:
        return None
    
    date_match = re.search(r'(\d{4}/\d{2}/\d{2})', css_class)
    if date_match:
        return date_match.group(1)
    
    return None


def parse_month_year_text(text: str) -> Tuple[Optional[int], Optional[int]]:
    """
    Parse un texte de mois en français vers (mois, année).
    
    Args:
        text: Texte à parser (ex: "janvier 2026", "février 2026")
        
    Returns:
        Tuple (numéro_mois, année) ou (None, None) si échec
        
    Exemples:
        >>> parse_month_year_text("janvier 2026")
        (1, 2026)
        >>> parse_month_year_text("février 2026")
        (2, 2026)
        >>> parse_month_year_text("texte invalide")
        (None, None)
    """
    from src.config import FRENCH_MONTHS
    
    if not text:
        return None, None
    
    text_lower = text.lower()
    
    for month_name, month_num in FRENCH_MONTHS.items():
        if month_name in text_lower:
            year_match = re.search(r'(\d{4})', text)
            if year_match:
                year = int(year_match.group(1))
                return month_num, year
    
    return None, None


# ============================================================
# GÉNÉRATION DE CODES DE STATUT
# ============================================================

def get_status_code(type_am: str, type_pm: str, detail_am: str, detail_pm: str) -> str:
    """
    Génère un code de statut pour une journée complète.
    
    Cette fonction combine les informations du matin et de l'après-midi
    pour produire un code synthétique représentant l'état de la journée.
    
    Codes possibles :
    - "CV" : Congés validés (journée entière)
    - "TV-AM" : Télétravail validé le matin uniquement
    - "CV/TV" : Congés le matin, télétravail l'après-midi
    - "W" : Week-end / Jour férié
    - "" : Présent au bureau
    
    Args:
        type_am: Type d'événement matin (PRESENT, CONGES, TELETRAVAIL, JOUR_NON_OUVRE)
        type_pm: Type d'événement après-midi
        detail_am: Détail matin (pour déterminer validation et RTT)
        detail_pm: Détail après-midi
        
    Returns:
        Code de statut (ex: "CV", "TV-AM", "CV/TV", "W", "")
        
    Exemples:
        >>> get_status_code("CONGES", "CONGES", "Congés (Validé)", "Congés (Validé)")
        'CV'
        >>> get_status_code("TELETRAVAIL", "PRESENT", "TT (Validé)", "")
        'TV-AM'
        >>> get_status_code("JOUR_NON_OUVRE", "JOUR_NON_OUVRE", "", "")
        'W'
    """
    # Journée complète non ouvrée
    if type_am == 'JOUR_NON_OUVRE' and type_pm == 'JOUR_NON_OUVRE':
        return 'W'
    
    # Journée complète présent
    if type_am == 'PRESENT' and type_pm == 'PRESENT':
        return ''
    
    def half_day_code(type_val: str, detail_val: str) -> str:
        """Génère le code pour une demi-journée"""
        if type_val == 'PRESENT':
            return 'P'
        elif type_val == 'TELETRAVAIL':
            return 'TV' if is_validated(detail_val) else 'TP'
        elif type_val == 'CONGES':
            if is_rtt(detail_val):
                return 'RV' if is_validated(detail_val) else 'RP'
            else:
                return 'CV' if is_validated(detail_val) else 'CP'
        elif type_val == 'JOUR_NON_OUVRE':
            return 'W'
        return ''
    
    code_am = half_day_code(type_am, detail_am)
    code_pm = half_day_code(type_pm, detail_pm)
    
    # Codes identiques matin et après-midi
    if code_am == code_pm:
        return '' if code_am == 'P' else code_am
    
    # Demi-journée matin ou après-midi
    if code_am == 'P' and code_pm != '':
        return f'{code_pm}-PM'
    if code_pm == 'P' and code_am != '':
        return f'{code_am}-AM'
    
    # Journée mixte (matin différent de l'après-midi)
    if code_am != '' and code_pm != '':
        return f'{code_am}/{code_pm}'
    
    return code_am if code_am else code_pm


# ============================================================
# COMPTAGE D'ÉVÉNEMENTS
# ============================================================

def count_event_weight(code: str, prefixes: Optional[list] = None) -> float:
    """
    Compte le poids d'un code en tenant compte des demi-journées.
    
    Cette fonction est utilisée pour calculer les totaux dans les feuilles Excel.
    Elle retourne le nombre de jours représentés par le code.
    
    Logique :
    - Journée entière (ex: "CV") = 1 jour
    - Demi-journée (ex: "CV-AM") = 0.5 jour
    - Journée mixte (ex: "CV/TV") = 0.5 jour pour chaque code
    - Week-end ("W") ou présent ("") = 0 jour
    
    Args:
        code: Code d'événement (ex: "CV", "TV-AM", "CV/TV")
        prefixes: Liste de préfixes à compter (None = tous sauf W et vide)
        
    Returns:
        Poids en jours (0, 0.5, ou 1)
        
    Exemples:
        >>> count_event_weight("CV")
        1.0
        >>> count_event_weight("CV-AM")
        0.5
        >>> count_event_weight("CV/TV")  # Sans filtre = compte tout
        1.0
        >>> count_event_weight("CV/TV", ["CV"])  # Avec filtre = compte seulement CV
        0.5
        >>> count_event_weight("W")
        0.0
    """
    if not code or code == '' or code == 'W':
        return 0
    
    # Cas journée mixte : "CV/TV", "RV/TP", etc.
    if '/' in code:
        parts = code.split('/')
        total = 0
        for part in parts:
            p = part.strip()
            if prefixes is None:
                total += 0.5
            else:
                for prefix in prefixes:
                    if p.startswith(prefix):
                        total += 0.5
                        break
        return total
    
    # Cas demi-journée : "CV-AM", "TV-PM", etc.
    if '-AM' in code or '-PM' in code:
        base = code.replace('-AM', '').replace('-PM', '')
        if prefixes is None:
            return 0.5
        for prefix in prefixes:
            if base.startswith(prefix):
                return 0.5
        return 0
    
    # Cas journée complète : "CV", "TV", "RV", etc.
    if prefixes is None:
        return 1
    for prefix in prefixes:
        if code.startswith(prefix):
            return 1
    return 0
