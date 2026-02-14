import calendar
from typing import List

def get_days_per_month(year: int) -> List[int]:
    """
    Retourne une liste contenant le nombre de jours
    pour chaque mois de l'année donnée.
    Gère automatiquement les années bissextiles.
    """
    return [calendar.monthrange(year, month)[1] for month in range(1, 13)]
