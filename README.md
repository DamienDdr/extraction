# DailyRH Leave Planning Scraper

Solution d'extraction automatisée des données de planning de congés depuis DailyRH (BNP Paribas) avec génération de rapports Excel analytiques.

## 📋 Vue d'ensemble

Ce projet permet de :
- **Extraire** automatiquement les plannings de congés de tous les collaborateurs
- **Analyser** les données selon les règles RH (10j consécutifs, 20j total sur période)
- **Générer** des rapports Excel formatés avec synthèses et calendriers

## 🏗️ Architecture

```
dailyrh_scraper/
├── src/                    # Code source principal
│   ├── config/            # Configuration et constantes
│   ├── logging/           # Système de logging
│   ├── utils/             # Fonctions utilitaires
│   ├── scraper/           # Logique d'extraction
│   └── excel/             # Génération de rapports
├── scripts/               # Scripts exécutables
│   ├── main.py           # Script principal
│   └── save_session.py   # Sauvegarde session SSO
├── output/                # Fichiers générés (CSV, Excel)
├── docs/                  # Documentation
└── requirements.txt       # Dépendances Python
```

## 🚀 Installation

### Prérequis
- Python 3.8 ou supérieur
- Accès à DailyRH (compte BNP Paribas)

### Installation des dépendances

```bash
# 1. Cloner ou extraire le projet
cd dailyrh_scraper

# 2. Installer les dépendances Python
pip install -r requirements.txt

# 3. Installer le navigateur Chromium pour Playwright
playwright install chromium
```

## 📖 Utilisation

### Première utilisation

#### 1. Sauvegarder la session SSO

```bash
python scripts/save_session.py
```

**Ce qu'il se passe :**
1. Un navigateur Chrome s'ouvre
2. Connectez-vous manuellement au SSO BNP Paribas
3. Appuyez sur ENTRÉE dans le terminal une fois connecté
4. La session est sauvegardée

⚠️ **Cette étape n'est à faire qu'une seule fois** (ou quand la session expire).

#### 2. Lancer le scraping

```bash
python scripts/main.py
```

**Le script va :**
1. Charger DailyRH avec votre session sauvegardée
2. Naviguer automatiquement vers janvier 2026
3. Extraire les données de tous les mois
4. Générer un CSV et un Excel dans `output/`

**Durée** : 15-20 minutes selon le nombre de collaborateurs

### Exécutions suivantes

Une fois la session sauvegardée, il suffit d'exécuter :

```bash
python scripts/main.py
```

## 📊 Fichiers générés

Tous les fichiers sont créés dans le répertoire `output/` :

| Fichier | Description |
|---------|-------------|
| `leave_planning_2026.csv` | Données brutes au format CSV |
| `rapport_conges_2026.xlsx` | Rapport Excel complet avec analyses |

### Structure du CSV

```csv
collaborateur,uid,date,type_am,detail_am,type_pm,detail_pm
Dupont Jean,123456,2026/01/15,CONGES,Congés (Validé),CONGES,Congés (Validé)
```

### Contenu du rapport Excel

**Feuille "Synthèse"**
- Compteurs par collaborateur (télétravail, congés, RTT)
- Validation des règles RH (10j consécutifs, 20j total)
- Totaux et statistiques globales

**Feuilles mensuelles** (Janvier, Février, ...)
- Vue par mois avec collaborateurs en lignes
- Codes couleur pour chaque type d'événement
- Totaux en bas de page

**Feuilles individuelles** (une par collaborateur)
- Calendrier annuel complet (12 mois × 31 jours)
- Vue d'ensemble du planning de l'année

## 🎨 Codes Excel

| Code | Signification | Couleur |
|------|---------------|---------|
| `CV` | Congés validés | 🟢 Vert |
| `CP` | Congés à valider | 🟡 Vert clair |
| `RV` | RTT validés | 🟠 Orange |
| `RP` | RTT à valider | 🟡 Orange clair |
| `TV` | Télétravail validé | 🔵 Bleu |
| `TP` | Télétravail à valider | 🔵 Bleu clair |
| `W` | Week-end / Férié | ⚫ Gris |
| `CV-AM` | Congés matin uniquement | 🟢 |
| `TV-PM` | Télétravail après-midi | 🔵 |
| `CV/TV` | Matin ≠ après-midi | ⚪ Mixte |

## ⚙️ Configuration

La configuration se trouve dans `src/config/config.py`. Vous pouvez modifier :

```python
# Année à extraire
TARGET_YEAR = 2026

# Mode sans interface graphique
HEADLESS_MODE = False  # True pour exécution serveur

# Délais de navigation (en secondes)
NAVIGATION_DELAY = 1.5
INITIAL_LOAD_DELAY = 10

# Règles RH
RULE_MIN_CONSECUTIVE_DAYS = 10
RULE_MIN_TOTAL_DAYS = 20

# Noms des fichiers de sortie
OUTPUT_CSV = "leave_planning_2026.csv"
OUTPUT_EXCEL = "rapport_conges_2026.xlsx"
```

## 🔍 Logging

Le système génère un fichier `dailyrh_scraper.log` avec :
- Progression du scraping
- Erreurs et avertissements
- Statistiques de collecte

Pour plus de détails, modifier le niveau dans `scripts/main.py` :

```python
logger = setup_logger(level="DEBUG")  # Au lieu de "INFO"
```

## 🛠️ Maintenance

Pour comprendre le code et apporter des modifications, consultez :

- `docs/MODULES.md` : Explication détaillée de chaque module
- `docs/QUICKSTART.md` : Guide de démarrage rapide

## ❓ Dépannage

### Erreur "Session file not found"

**Solution** : Exécutez d'abord `python scripts/save_session.py`

### Session expirée

**Solution** : Ré-exécutez `python scripts/save_session.py`

### Timeout / Navigation lente

**Solution** : Augmentez les délais dans `src/config/config.py` :
```python
NAVIGATION_DELAY = 3.0
INITIAL_LOAD_DELAY = 15
```

### Données incomplètes

**Solution** : Consultez `dailyrh_scraper.log` pour identifier le mois problématique

## 🔒 Sécurité

**⚠️ IMPORTANT :**
- Ne committez JAMAIS `bnpparibas_session.json` (contient vos cookies)
- Ne partagez JAMAIS les fichiers CSV/Excel (données personnelles)
- Le `.gitignore` est configuré pour protéger ces fichiers

## 📞 Support

Pour toute question :
1. Consultez `dailyrh_scraper.log`
2. Lisez `docs/MODULES.md`
3. Activez le mode DEBUG pour plus de détails

## 📝 Licence

Usage interne BNP Paribas uniquement.
