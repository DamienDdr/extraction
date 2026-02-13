#!/usr/bin/env python3
"""
Script de sauvegarde de session SSO pour DailyRH

Ce script permet de sauvegarder la session d'authentification SSO BNP Paribas
pour une utilisation ultérieure par le scraper principal.

Fonctionnement :
1. Ouvre un navigateur Chrome
2. Charge la page DailyRH
3. Attend que l'utilisateur se connecte manuellement au SSO
4. Sauvegarde les cookies et le stockage de session dans un fichier JSON

Le fichier de session généré est ensuite utilisé par scripts/main.py pour
éviter de devoir se reconnecter à chaque exécution.

Prérequis :
- Playwright installé (pip install playwright)
- Navigateur Chromium installé (playwright install chromium)
- Accès à DailyRH (compte BNP Paribas)

Utilisation :
    python scripts/save_session.py

Fichier généré :
- bnpparibas_session.json : Session SSO (à ne JAMAIS committer sur Git)

Note importante :
    Ce script doit être exécuté UNE SEULE FOIS avant la première utilisation
    du scraper principal, ou lorsque la session SSO expire (généralement
    après quelques jours/semaines).
"""

import sys
from pathlib import Path
from playwright.sync_api import sync_playwright

# Ajouter le répertoire parent au path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.config import SESSION_FILE, DAILYRH_URL


def main():
    """Sauvegarde la session SSO après authentification manuelle"""
    
    print("="*60)
    print("DailyRH Session Saver")
    print("="*60)
    print("\nCe script va ouvrir un navigateur pour vous permettre")
    print("de vous authentifier manuellement au SSO BNP Paribas.")
    print("\nLe fichier de session sera sauvegardé dans :")
    print(f"  → {SESSION_FILE}")
    print("\n⚠️  IMPORTANT : Ne partagez JAMAIS ce fichier (contient vos cookies)")
    print()
    
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,  # OBLIGATOIRE pour le SSO
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context()
        page = context.new_page()
        
        print(f"Ouverture de {DAILYRH_URL}...")
        page.goto(DAILYRH_URL)
        
        print("\n" + "="*60)
        print("👉 CONNECTEZ-VOUS MANUELLEMENT AU SSO")
        print("="*60)
        print("\nÉtapes :")
        print("1. Entrez votre identifiant BNP Paribas")
        print("2. Entrez votre mot de passe")
        print("3. Validez l'authentification à deux facteurs si demandée")
        print("4. Attendez d'être sur la page DailyRH (Team Planning)")
        print("\nUne fois connecté et sur la page DailyRH,")
        print("appuyez sur ENTRÉE ici pour sauvegarder la session...")
        input()
        
        # Sauvegarde de la session
        context.storage_state(path=SESSION_FILE)
        print(f"\n✅ Session sauvegardée dans {SESSION_FILE}")
        print("\nVous pouvez maintenant exécuter le scraper principal")
        print("avec la commande : python scripts/main.py")
        
        browser.close()
    
    print("\n" + "="*60)
    print("Session sauvegardée avec succès")
    print("="*60)
    print("\n💡 Conseils :")
    print("  - Cette session est valide pendant quelques jours/semaines")
    print("  - Ré-exécutez ce script si vous obtenez des erreurs d'authentification")
    print("  - Ne committez JAMAIS le fichier de session sur Git")


if __name__ == "__main__":
    main()
