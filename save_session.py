# save_session.py
from playwright.sync_api import sync_playwright

SESSION_FILE = "bnpparibas_session.json"

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=False,  # OBLIGATOIRE pour le SSO
        args=["--disable-blink-features=AutomationControlled"]
    )
    context = browser.new_context()
    page = context.new_page()

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")

    print("👉 Connecte-toi MANUELLEMENT au SSO, puis appuie sur Entrée ici...")
    input()

    # Sauvegarde cookies + storage
    context.storage_state(path=SESSION_FILE)
    print(f"✅ Session sauvegardée dans {SESSION_FILE}")

    browser.close()
