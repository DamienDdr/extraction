from playwright.sync_api import sync_playwright
import time

SESSION_FILE = "bnpparibas_session.json"

ROW_INDEX = 0          # collaborateur (0 = premier)
DAYS_TO_INSPECT = [1, 2, 3, 4, 5, 6]  # colonnes jours à analyser

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state=SESSION_FILE)
    page = context.new_page()

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")
    page.wait_for_load_state("networkidle")
    time.sleep(4)

    print("✅ Page chargée – inspection DOM")

    rows = page.locator("tr.dhx_row_item")
    if rows.count() == 0:
        print("❌ Aucune ligne DHTMLX détectée")
        browser.close()
        exit()

    row = rows.nth(ROW_INDEX)
    cells = row.locator("td")

    print(f"\n👤 Inspection ligne {ROW_INDEX}")
    print("=" * 80)

    for day_idx in DAYS_TO_INSPECT:
        if day_idx >= cells.count():
            continue

        cell = cells.nth(day_idx)

        print(f"\n📅 COLONNE JOUR INDEX = {day_idx}")
        print("-" * 60)

        # HTML complet
        outer_html = cell.evaluate("el => el.outerHTML")
        print("🔹 outerHTML:")
        print(outer_html)

        # Classes
        classes = cell.get_attribute("class")
        print("\n🔹 classes:")
        print(classes)

        # Texte visible
        text = (cell.text_content() or "").strip()
        print("\n🔹 textContent:")
        print(repr(text))

        # Enfants directs
        children = cell.evaluate("""
            el => Array.from(el.children).map(c => ({
                tag: c.tagName,
                class: c.className,
                title: c.getAttribute('title'),
                html: c.outerHTML
            }))
        """)

        print("\n🔹 enfants directs:")
        for c in children:
            print(c)

    browser.close()
