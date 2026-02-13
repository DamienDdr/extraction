"""
Module de gestion du logging

Ce module configure un système de logging centralisé pour l'ensemble du projet.
Il permet d'enregistrer les messages dans un fichier et/ou dans la console,
avec différents niveaux de verbosité (DEBUG, INFO, WARNING, ERROR, CRITICAL).

Utilisation :
    from src.logging import get_logger
    
    logger = get_logger()
    logger.info("Message d'information")
    logger.warning("Message d'avertissement")
    logger.error("Message d'erreur")
"""

import logging
import sys
from pathlib import Path
from typing import Optional


def setup_logger(
    name: str = "dailyrh_scraper",
    log_file: Optional[str] = None,
    level: int = logging.INFO
) -> logging.Logger:
    """
    Configure et retourne un logger.
    
    Cette fonction crée un logger avec :
    - Un handler console pour afficher les messages dans le terminal
    - Un handler fichier optionnel pour sauvegarder les logs
    - Un format standardisé avec timestamp, niveau et message
    
    Args:
        name: Nom du logger (par défaut "dailyrh_scraper")
        log_file: Chemin du fichier de log (optionnel, None = pas de fichier)
        level: Niveau minimum de log (DEBUG, INFO, WARNING, ERROR, CRITICAL)
        
    Returns:
        Logger configuré et prêt à l'emploi
        
    Exemple:
        >>> logger = setup_logger(name="mon_scraper", log_file="mon_log.log", level=logging.DEBUG)
        >>> logger.info("Démarrage du scraper")
        >>> logger.debug("Variable x = 42")
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # Éviter les handlers dupliqués si la fonction est appelée plusieurs fois
    if logger.handlers:
        return logger
    
    # Format des messages : [2026-02-13 14:30:45] INFO - Message ici
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Handler console : affiche dans le terminal
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # Handler fichier : sauvegarde dans un fichier (optionnel)
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger


def get_logger(name: str = "dailyrh_scraper") -> logging.Logger:
    """
    Récupère un logger existant ou en crée un nouveau.
    
    Cette fonction est un raccourci pour obtenir un logger sans avoir
    à le reconfigurer. Si le logger n'existe pas encore, il sera créé
    avec la configuration par défaut.
    
    Args:
        name: Nom du logger à récupérer
        
    Returns:
        Logger existant ou nouveau logger
        
    Exemple:
        >>> logger = get_logger()
        >>> logger.info("Message simple")
    """
    return logging.getLogger(name)
