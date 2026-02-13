# 🚀 Guide de démarrage rapide

Guide pas à pas pour utiliser le DailyRH Scraper en 15 minutes.

## ⏱️ Étape 1 : Installation (5 minutes)

```bash
# Naviguer dans le dossier du projet
cd dailyrh_scraper

# Installer les dépendances Python
pip install -r requirements.txt

# Installer le navigateur Chromium
playwright install chromium
```

✅ **Installation terminée !**

## 🔐 Étape 2 : Sauvegarder la session SSO (5 minutes)

```bash
python scripts/save_session.py
```

### Ce qu'il va se passer :

1. ✅ Un navigateur Chrome s'ouvre automatiquement
2. ✅ La page DailyRH se charge
3. 👉 **Vous devez vous connecter manuellement**
   - Entrez votre identifiant BNP Paribas
   - Entrez votre mot de passe  
   - Validez l'authentification à deux facteurs si demandée
4. ✅ Une fois sur la page Team Planning, retournez au terminal
5. ✅ Appuyez sur **ENTRÉE**
6. ✅ Le message "Session sauvegardée" apparaît

**✨ Cette étape n'est à faire qu'UNE SEULE FOIS !**

(ou lorsque la session expire après quelques jours/semaines)

## ▶️ Étape 3 : Lancer le scraping (15-20 minutes)

```bash
python scripts/main.py
```

### Ce qu'il va se passer :

1. ✅ Chargement de DailyRH avec votre session
2. ✅ Navigation automatique vers janvier 2026
3. ✅ Extraction des données de **tous les mois** de l'année
4. ✅ Génération du CSV dans `output/`
5. ✅ Génération de l'Excel dans `output/`
6. ✅ Message "Traitement terminé avec succès"

### Pendant le scraping

Vous verrez défiler des messages comme :
```
Traitement du mois : Janvier 2026
Nombre de collaborateurs : 25
Lignes extraites : 775
...
```

**Ne fermez pas le terminal pendant l'exécution !**

## 📁 Étape 4 : Récupérer les fichiers

Les fichiers sont dans le dossier `output/` :

```
output/
├── leave_planning_2026.csv         # Données brutes
└── rapport_conges_2026.xlsx        # Rapport Excel
```

### Ouvrir l'Excel

Double-cliquez sur `rapport_conges_2026.xlsx` pour voir :

- **Feuille "Synthèse"** : Vue d'ensemble avec compteurs
- **Feuilles mensuelles** : Janvier, Février, Mars, ...
- **Feuilles individuelles** : Une par collaborateur

## 🎯 Prochaines utilisations

La prochaine fois, il suffit d'exécuter :

```bash
python scripts/main.py
```

**C'est tout !** La session est déjà sauvegardée. ☕

## 🎨 Comprendre les codes Excel

| Code | Signification |
|------|---------------|
| `CV` | Congés validés (journée entière) |
| `CP` | Congés à valider |
| `TV` | Télétravail validé |
| `W` | Week-end ou jour férié |
| `CV-AM` | Congés le matin uniquement |
| `CV/TV` | Congés matin, télétravail après-midi |

**Voir le README.md pour la liste complète**

## ⚙️ Personnalisation rapide

### Changer l'année

Éditez `src/config/config.py` :
```python
TARGET_YEAR = 2027  # Au lieu de 2026
```

### Mode invisible (sans navigateur)

Éditez `src/config/config.py` :
```python
HEADLESS_MODE = True  # Le navigateur ne s'affiche plus
```

### Plus de détails dans les logs

Éditez `scripts/main.py` :
```python
logger = setup_logger(level="DEBUG")  # Au lieu de "INFO"
```

## ❓ Problèmes fréquents

### ❌ "Session file not found"

**Cause** : Vous n'avez pas encore sauvegardé la session

**Solution** :
```bash
python scripts/save_session.py
```

### ❌ "Authentication failed" ou "Login required"

**Cause** : La session a expiré

**Solution** :
```bash
python scripts/save_session.py  # Rafraîchir la session
```

### ❌ Le scraping se bloque ou timeout

**Cause** : Les délais sont trop courts

**Solution** : Augmentez les délais dans `src/config/config.py` :
```python
NAVIGATION_DELAY = 3.0         # Au lieu de 1.5
INITIAL_LOAD_DELAY = 15        # Au lieu de 10
```

### ❌ Données manquantes pour certains mois

**Cause** : Erreur pendant le scraping d'un mois

**Solution** : 
1. Consultez `dailyrh_scraper.log`
2. Identifiez le mois problématique
3. Ré-exécutez le script

## 📖 Aller plus loin

- **README.md** : Documentation complète
- **docs/MODULES.md** : Comprendre l'architecture du code
- **dailyrh_scraper.log** : Journal détaillé d'exécution

## ✅ Récapitulatif

```bash
# Installation (une fois)
pip install -r requirements.txt
playwright install chromium

# Sauvegarder la session (une fois)
python scripts/save_session.py

# Exécuter le scraping (à chaque fois)
python scripts/main.py

# Récupérer les fichiers
ls output/
```

**Durée totale** : ~25 minutes (première fois), ~15 minutes (suivantes)

🎉 **C'est tout !** Vous êtes prêt à utiliser le scraper.
