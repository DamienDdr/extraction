"""
DailyRH Leave Planning Scraper

Package principal contenant tous les modules du projet.
"""

__version__ = "1.0.0"
__author__ = "BNP Paribas"

from .config import *
from .logging import setup_logger, get_logger
from .utils import *
from .scraper import scrape_all_months
from .excel import analyze_leave_data, create_excel_report
