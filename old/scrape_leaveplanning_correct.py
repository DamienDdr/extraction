from playwright.sync_api import sync_playwright
import pandas as pd
import re
import time

SESSION_FILE = "../bnpparibas_session.json"
OUTPUT_FILE = "../leave_planning.csv"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)
    context = browser.new_context(storage_state=SESSION_FILE)
    page = context.new_page()

    page.goto("https://dailyrh.hr.bnpparibas/app/foryou/#/demarches/leaveplanning")
    page.wait_for_load_state("networkidle")
    time.sleep(5)

    print("✅ Page chargée")

    # ==============================
    # 1️⃣ EXTRACTION COLLABORATEURS
    # ==============================
    collab_map = {}

    collab_rows = page.locator("div.dhx_matrix_scell.MYCOLLEAGUES")
    nb = collab_rows.count()
    print(f"👤 {nb} collaborateurs détectés (colonne gauche)")

    for i in range(nb):
        row = collab_rows.nth(i)
        classes = row.get_attribute("class") or ""
        name = row.inner_text().strip()

        m = re.search(r"(HRF[a-zA-Z0-9]+)", classes)
        if not m:
            continue

        hrf_id = m.group(1)
        collab_map[hrf_id] = name

    # ==============================
    # 2️⃣ EXTRACTION ÉVÉNEMENTS
    # ==============================
    events = page.locator("div.dhx_cal_event")
    count = events.count()
    print(f"📅 {count} événements détectés")

    rows = []

    for i in range(count):
        ev = events.nth(i)

        classes = ev.get_attribute("class") or ""
        title = ev.inner_text().strip()

        m = re.search(r"(HRF[a-zA-Z0-9]+)", classes)
        if not m:
            continue

        hrf_id = m.group(1)
        nom = collab_map.get(hrf_id, hrf_id)

        rows.append({
            "collaborateur": nom,
            "hrf": hrf_id,
            "titre": title,
            "classes": classes
        })

    # ==============================
    # 3️⃣ EXPORT CSV
    # ==============================
    df = pd.DataFrame(rows)
    df.to_csv(OUTPUT_FILE, index=False, encoding="utf-8-sig")

    print(f"📁 Export terminé → {OUTPUT_FILE}")
    browser.close()
