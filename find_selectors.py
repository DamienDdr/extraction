from playwright.sync_api import sync_playwright
import time

SESSION_FILE = "bnpparibas_session.json"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state=SESSION_FILE)
    page = context.new_page()

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")
    page.wait_for_load_state("networkidle")
    time.sleep(4)

    print("=" * 80)
    print("🔍 RECHERCHE DES SÉLECTEURS DE NAVIGATION")
    print("=" * 80)

    # ===== 1. CHERCHER LE MOIS AFFICHÉ =====
    print("\n📅 1. TEXTE DU MOIS AFFICHÉ :")
    print("-" * 80)

    # Chercher tous les éléments contenant "Février" ou "2026"
    candidates_month = [
        ".dhx_cal_date",
        ".dhx_cal_tab",
        "div:has-text('Février')",
        "div:has-text('2026')",
        "[class*='date']",
        "[class*='month']",
        "[class*='calendar']"
    ]

    for selector in candidates_month:
        try:
            elements = page.locator(selector)
            count = elements.count()
            if count > 0:
                for i in range(min(count, 3)):  # Max 3 premiers éléments
                    text = elements.nth(i).inner_text().strip()
                    if text and ("2026" in text or "Février" in text or "février" in text):
                        print(f"✅ Sélecteur: {selector}")
                        print(f"   Texte: {text}")
                        print(f"   HTML: {elements.nth(i).get_attribute('class')}")
                        print()
        except:
            pass

    # ===== 2. CHERCHER LES BOUTONS DE NAVIGATION =====
    print("\n◀️▶️ 2. BOUTONS DE NAVIGATION :")
    print("-" * 80)

    # Chercher les boutons avec flèches
    candidates_buttons = [
        "button.dhx_cal_next_button",
        "button.dhx_cal_prev_button",
        "div.dhx_cal_next_button",
        "div.dhx_cal_prev_button",
        "[class*='next']",
        "[class*='prev']",
        "button:has-text('›')",
        "button:has-text('‹')",
        "button:has-text('>')",
        "button:has-text('<')",
        "div:has-text('›')",
        "div:has-text('‹')",
        "[title*='suivant']",
        "[title*='précédent']",
        "[aria-label*='next']",
        "[aria-label*='previous']"
    ]

    for selector in candidates_buttons:
        try:
            elements = page.locator(selector)
            count = elements.count()
            if count > 0:
                for i in range(min(count, 2)):
                    elem = elements.nth(i)
                    text = elem.inner_text().strip() if elem.inner_text() else "(vide)"
                    classes = elem.get_attribute('class') or ""
                    title = elem.get_attribute('title') or ""

                    # Vérifier si c'est visible
                    is_visible = elem.is_visible()

                    if is_visible or "next" in classes or "prev" in classes:
                        print(f"✅ Sélecteur: {selector}")
                        print(f"   Texte: {text}")
                        print(f"   Classes: {classes}")
                        print(f"   Title: {title}")
                        print(f"   Visible: {is_visible}")
                        print()
        except:
            pass

    # ===== 3. INSPECTION VISUELLE DANS LE NAVIGATEUR =====
    print("\n🖱️ 3. INSPECTION MANUELLE :")
    print("-" * 80)
    print("Le navigateur reste ouvert. Fais ceci :")
    print("1. Clique droit sur la flèche '>' (mois suivant)")
    print("2. Choisis 'Inspecter' (Inspect)")
    print("3. Note la classe CSS (ex: class='dhx_cal_next_button')")
    print()
    print("4. Fais pareil pour le texte 'Février 2026'")
    print("5. Note le sélecteur (ex: class='dhx_cal_date')")
    print()
    print("Appuie sur ENTRÉE quand tu as noté les sélecteurs...")

    input()

    browser.close()

    print("\n✅ Script terminé. Donne-moi les sélecteurs que tu as trouvés !")