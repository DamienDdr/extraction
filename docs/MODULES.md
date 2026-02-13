# 📚 Guide des modules

Ce document explique l'architecture et le rôle de chaque module du projet pour faciliter la maintenance et les évolutions.

## 📁 Structure du projet

```
dailyrh_scraper/
├── src/                      # Code source
│   ├── config/              # Configuration
│   ├── logging/             # Gestion des logs
│   ├── utils/               # Fonctions utilitaires
│   ├── scraper/             # Extraction des données
│   └── excel/               # Génération Excel
├── scripts/                  # Scripts exécutables
│   ├── main.py              # Script principal
│   └── save_session.py      # Sauvegarde session
├── output/                   # Fichiers générés
└── docs/                     # Documentation
```

---

## 🔧 Module: src/config/

**Fichier** : `src/config/config.py`

**Rôle** : Centralise toutes les constantes et paramètres de configuration du projet.

### Sections principales

#### 1. Configuration des fichiers

```python
SESSION_FILE = "bnpparibas_session.json"  # Fichier de session SSO
OUTPUT_DIR = "output"                      # Répertoire de sortie
OUTPUT_CSV = "leave_planning_2026.csv"     # Nom du CSV
OUTPUT_EXCEL = "rapport_conges_2026.xlsx"  # Nom de l'Excel
```

**Pourquoi** : Permet de changer les chemins sans toucher au code.

#### 2. Configuration du scraping

```python
DAILYRH_URL = "https://..."        # URL de DailyRH
TARGET_YEAR = 2026                  # Année à extraire
HEADLESS_MODE = False               # Navigateur visible/invisible
NAVIGATION_DELAY = 1.5              # Délai entre les mois
INITIAL_LOAD_DELAY = 10             # Délai initial
MAX_NAVIGATION_CLICKS = 50          # Limite de clics
```

**Pourquoi** : Ajuster les délais selon la performance du réseau.

#### 3. Règles RH

```python
RULE_START_DATE = datetime(2026, 5, 15)   # Début période
RULE_END_DATE = datetime(2026, 10, 15)     # Fin période
RULE_MIN_CONSECUTIVE_DAYS = 10             # 10j consécutifs
RULE_MIN_TOTAL_DAYS = 20                   # 20j total
```

**Pourquoi** : Les règles RH peuvent changer d'une année à l'autre.

#### 4. Styles Excel

Toutes les couleurs, bordures, et dimensions des cellules Excel sont définies ici.

**Pourquoi** : Modifier l'apparence du rapport en un seul endroit.

### 💡 Cas d'usage

**Changer l'année cible** :
```python
TARGET_YEAR = 2027
```

**Activer le mode serveur (sans navigateur)** :
```python
HEADLESS_MODE = True
```

**Modifier les règles RH** :
```python
RULE_MIN_CONSECUTIVE_DAYS = 15
RULE_MIN_TOTAL_DAYS = 25
```

---

## 📝 Module: src/logging/

**Fichier** : `src/logging/logger.py`

**Rôle** : Gère l'affichage et l'enregistrement des messages du programme.

### Fonctions principales

#### `setup_logger(name, log_file, level)`

Configure un logger avec affichage console et fichier.

```python
logger = setup_logger(
    name="dailyrh_scraper",
    log_file="dailyrh_scraper.log",
    level=logging.INFO
)
```

**Niveaux disponibles** :
- `DEBUG` : Messages très détaillés (debugging)
- `INFO` : Messages informatifs normaux
- `WARNING` : Avertissements
- `ERROR` : Erreurs
- `CRITICAL` : Erreurs critiques

#### `get_logger(name)`

Récupère un logger déjà configuré.

```python
logger = get_logger()
logger.info("Message d'information")
logger.warning("Attention !")
logger.error("Erreur détectée")
```

### 💡 Cas d'usage

**Activer le mode debug** :

Dans `scripts/main.py`, changer :
```python
logger = setup_logger(level=logging.DEBUG)  # Plus verbeux
```

**Désactiver le fichier de log** :
```python
logger = setup_logger(log_file=None)  # Uniquement console
```

---

## 🛠️ Module: src/utils/

**Fichier** : `src/utils/utils.py`

**Rôle** : Regroupe toutes les fonctions réutilisables pour le traitement des données.

### Catégories de fonctions

#### 1. Conversion de dates

```python
date_to_string(date(2026, 1, 15))
# → "2026/01/15"
```

#### 2. Validation des événements

```python
is_validated("Congés (validé)")      # → True
is_validated("RTT (à valider)")      # → False
is_rtt("RTT (validé)")               # → True
```

#### 3. Construction de détails

```python
build_detail("Congés", "Validé")
# → "Congés (Validé)"
```

#### 4. Extraction HTML/CSS

```python
extract_uid_from_corp_id("HRF344256-0_HRF460606")
# → "460606"

extract_date_from_css_class("grey_cell_weekend 2026/04/06")
# → "2026/04/06"

parse_month_year_text("janvier 2026")
# → (1, 2026)
```

#### 5. Génération de codes

```python
get_status_code("CONGES", "CONGES", "CP (V)", "CP (V)")
# → "CV" (Congés Validés)

get_status_code("TELETRAVAIL", "PRESENT", "TT (V)", "")
# → "TV-AM" (Télétravail le matin)
```

#### 6. Comptage d'événements

```python
count_event_weight("CV")          # → 1.0 (journée)
count_event_weight("CV-AM")       # → 0.5 (demi-journée)
count_event_weight("CV/TV")       # → 1.0 (journée mixte)
count_event_weight("CV/TV", ["CV"])  # → 0.5 (seulement CV)
```

### 💡 Cas d'usage

**Ajouter un nouveau type d'événement** :

Modifier `get_status_code()` pour gérer le nouveau type.

**Changer la logique de validation** :

Modifier `is_validated()` pour détecter de nouveaux patterns.

---

## 🌐 Module: src/scraper/

**Fichier** : `src/scraper/scraper.py`

**Rôle** : Extrait les données de planning depuis DailyRH via Playwright.

### Architecture du scraping

```
scrape_all_months()
    └─► scrape_month(year, month)
           ├─► extract_non_working_days()          # Jours fériés/WE
           ├─► extract_collaborator_events()       # Événements normaux
           ├─► apply_half_day_events()             # Demi-journées
           ├─► apply_full_day_events()             # Journées entières
           └─► apply_non_working_days()            # Priorité absolue JNO
```

### Fonctions principales

#### `scrape_all_months(year)`

Scrape tous les mois d'une année complète.

**Workflow** :
1. Ouvre le navigateur avec la session SSO
2. Charge DailyRH
3. Navigue vers janvier
4. Pour chaque mois (1-12) :
   - Scrape le mois
   - Passe au mois suivant
5. Retourne tous les records

#### `scrape_month(page, year, month)`

Scrape un mois donné.

**Workflow** :
1. Compte le nombre de collaborateurs
2. Extrait les jours non ouvrés (une seule fois pour tout le mois)
3. Pour chaque collaborateur :
   - Extrait les événements (CONGES, TELETRAVAIL)
   - Applique les événements dans l'ordre de priorité
   - Génère les records CSV

#### Ordre de priorité des événements

```
1. Demi-journées          (AM/PM spécifiques)
2. Journées entières      (CONGES puis TELETRAVAIL)
3. Jours non ouvrés       (Écrase TOUT, priorité absolue)
```

**Pourquoi cet ordre** : Un congé qui chevauche un week-end ne compte PAS le week-end comme congé.

### Détection des événements

#### Classes CSS importantes

```python
"grey_cell_weekend"      # Jour non ouvré
"telework"               # Télétravail
"validated_vcell"        # Événement validé
"to_validate_vcell"      # Événement à valider
```

#### Calcul des positions

Les événements sont positionnés en pixels dans le DOM. Le scraper :
1. Calcule la largeur d'une colonne (jour)
2. Convertit la position pixel en indice de jour
3. Détecte les demi-journées (largeur < 3px)

### 💡 Cas d'usage

**Ajouter un nouveau type d'événement** :

Modifier `determine_event_type_and_status()` pour détecter les nouvelles classes CSS.

**Ajuster la détection des demi-journées** :

Modifier `is_half_day()` :
```python
def is_half_day(width_px, col_width):
    return width_px <= 3  # Au lieu de 2
```

**Scraper une autre année** :

Dans `src/config/config.py` :
```python
TARGET_YEAR = 2027
```

---

## 📊 Module: src/excel/

**Fichier** : `src/excel/excel_generator.py`

**Rôle** : Analyse les données et génère le rapport Excel formaté.

### Fonctions principales

#### `analyze_leave_data(csv_file)`

Analyse le CSV et calcule les statistiques.

**Retourne** :
```python
{
    "Dupont Jean": {
        "uid": "123456",
        "teletravail_valide_am": 5,
        "conges_valides_am": 12,
        "rtt_valides_am": 3,
        "regle_10j_consecutifs": True,
        "regle_20j_total": True,
        "jours_consecutifs_max": 15,
        "jours_total_periode": 22,
        ...
    },
    ...
}
```

**Logique des règles RH** :
1. Filtre les données sur la période (15 mai - 15 octobre)
2. Identifie les jours de congés
3. Calcule les jours consécutifs (en ignorant les week-ends)
4. Calcule le total de jours

#### `create_excel_report(stats, csv_file, output_file)`

Crée le fichier Excel avec 3 types de feuilles.

##### 1. Feuille "Synthèse"

```
┌─────────────┬─────┬──────────────┬─────────┬──────────┐
│ Collaborateur│ UID │ Télétravail │ Congés │ Règle 10j│
├─────────────┼─────┼──────────────┼─────────┼──────────┤
│ Dupont Jean │12345│    5.0 j    │ 12.0 j │    ✓     │
│ Martin Paul │67890│    3.0 j    │  8.0 j │    ✗     │
└─────────────┴─────┴──────────────┴─────────┴──────────┘
```

##### 2. Feuilles mensuelles

```
Janvier 2026
┌─────────────┬───┬───┬───┬───┬───┬
│Collaborateur│ 1 │ 2 │ 3 │ 4 │...│
├─────────────┼───┼───┼───┼───┼───┤
│ Dupont Jean │   │ W │ W │CV │TV │
│ Martin Paul │TV │ W │ W │   │   │
├─────────────┼───┼───┼───┼───┼───┤
│TOTAL (j)    │0.5│ 0 │ 0 │0.5│0.5│
└─────────────┴───┴───┴───┴───┴───┘
```

##### 3. Feuilles individuelles

```
PLANNING 2026 - Dupont Jean
┌─────────┬───┬───┬───┬───┬───┬
│  Mois   │ 1 │ 2 │ 3 │ 4 │...│
├─────────┼───┼───┼───┼───┼───┤
│ Janvier │   │ W │ W │CV │TV │
│ Février │   │   │ W │ W │   │
│ Mars    │RV │   │   │ W │ W │
└─────────┴───┴───┴───┴───┴───┘
```

### Application des styles

```python
apply_cell_style(cell, code, is_even_row)
```

**Logique** :
- Code "CV" → Remplissage vert
- Code "TV" → Remplissage bleu
- Code "W" → Remplissage gris
- Code "CV/TV" → Remplissage mixte
- Lignes paires → Fond légèrement coloré

### Comptage des totaux

```python
count_event_weight(code, prefixes)
```

**Exemple** :
```python
# Total de tous les événements
count_event_weight("CV")  # → 1.0
count_event_weight("CV-AM")  # → 0.5

# Total seulement des congés
count_event_weight("CV", ["CV", "CP"])  # → 1.0
count_event_weight("TV", ["CV", "CP"])  # → 0.0
```

### 💡 Cas d'usage

**Modifier les couleurs** :

Dans `src/config/config.py` :
```python
CV_FILL = PatternFill(start_color="00FF00", ...)  # Vert plus vif
```

**Ajouter une colonne dans la synthèse** :

1. Calculer la métrique dans `analyze_leave_data()`
2. Ajouter la colonne dans `create_summary_sheet()`

**Changer les règles RH** :

Dans `src/config/config.py` :
```python
RULE_MIN_CONSECUTIVE_DAYS = 15  # Au lieu de 10
```

---

## ▶️ Scripts: scripts/

### `scripts/main.py`

**Rôle** : Orchestre l'ensemble du processus de bout en bout.

**Workflow** :
```
1. Configuration du logging
2. Création du répertoire output/
3. Scraping de tous les mois
4. Export CSV
5. Analyse des données
6. Génération Excel
7. Affichage du résumé
```

**Gestion d'erreurs** :
- `KeyboardInterrupt` : Interruption manuelle (Ctrl+C)
- `Exception` : Toute autre erreur → log et exit

### `scripts/save_session.py`

**Rôle** : Sauvegarde la session SSO pour éviter de se reconnecter à chaque fois.

**Workflow** :
```
1. Ouvre un navigateur non-headless
2. Charge DailyRH
3. Attend la connexion manuelle de l'utilisateur
4. Sauvegarde cookies + localStorage dans un fichier JSON
5. Ferme le navigateur
```

**Fichier généré** : `bnpparibas_session.json`

⚠️ **Ce fichier contient des données sensibles et ne doit JAMAIS être committé.**

---

## 🔄 Flux de données

```
DailyRH (web)
    ↓
[Playwright Browser]
    ↓
scraper.py → Records CSV
    ↓
Pandas DataFrame
    ↓
excel_generator.py → Analyse
    ↓
Openpyxl Workbook
    ↓
Fichier Excel
```

---

## 🧪 Debugging

### Activer le mode DEBUG

Dans `scripts/main.py` :
```python
logger = setup_logger(level=logging.DEBUG)
```

Vous verrez alors :
```
DEBUG - Traité : Dupont Jean (123456)
DEBUG - Event: left=150px, width=30px, type=CONGES
DEBUG - → start_idx=5 (jour 6), end_idx=7 (jour 8)
```

### Consulter les logs

```bash
cat dailyrh_scraper.log
```

Rechercher les erreurs :
```bash
grep ERROR dailyrh_scraper.log
grep WARNING dailyrh_scraper.log
```

### Tester un seul mois

Modifier `scraper.py` temporairement :
```python
# Au lieu de :
for month in range(1, 13):

# Faire :
for month in range(1, 2):  # Seulement janvier
```

---

## 📦 Dépendances

**Fichier** : `requirements.txt`

```
playwright==1.40.0     # Automatisation navigateur
pandas==2.1.4          # Manipulation de données
openpyxl==3.1.2        # Génération Excel
```

**Installation** :
```bash
pip install -r requirements.txt
playwright install chromium
```

---

## 🎯 Points d'attention pour la maintenance

### 1. Changements de l'interface DailyRH

Si DailyRH change son interface, vérifier :
- Les sélecteurs CSS dans `scraper.py`
- Les classes d'événements dans `determine_event_type_and_status()`

### 2. Nouvelles règles RH

Modifier dans `src/config/config.py` et potentiellement dans `analyze_leave_data()`.

### 3. Performance

Si le scraping est lent :
- Augmenter les délais dans `src/config/config.py`
- Vérifier la qualité de la connexion réseau
- Activer `HEADLESS_MODE = True` pour gagner du temps

### 4. Nouveaux types d'événements

Si DailyRH ajoute de nouveaux types (ex: télétravail partiel) :
1. Ajouter la détection dans `determine_event_type_and_status()`
2. Ajouter les couleurs dans `src/config/config.py`
3. Ajouter la logique dans `get_status_code()`
4. Mettre à jour `apply_cell_style()`

---

## ✅ Checklist de maintenance

Avant de modifier le code :

- [ ] Lire ce document (MODULES.md)
- [ ] Identifier le module concerné
- [ ] Comprendre le flux de données
- [ ] Tester sur un seul mois d'abord
- [ ] Vérifier les logs en mode DEBUG
- [ ] Tester avec plusieurs collaborateurs
- [ ] Vérifier le fichier Excel généré

---

Ce guide devrait vous permettre de comprendre et maintenir le projet facilement. Pour toute question, consultez les logs et le code source avec les docstrings détaillées.
