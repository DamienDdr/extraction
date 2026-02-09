#!/usr/bin/env python3
"""
Script pour générer des statistiques de congés par personne à partir du CSV de planning
Avec planning annuel dynamique - VERSION CORRIGÉE
"""

import pandas as pd
from collections import defaultdict
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule


def is_validated(detail):
    """Détermine si un événement est validé ou à valider"""
    if pd.isna(detail) or detail == '':
        return None
    detail_lower = str(detail).lower()
    if 'validé' in detail_lower or '(validé)' in detail_lower:
        return True
    elif 'à valider' in detail_lower or 'a valider' in detail_lower:
        return False
    return None


def is_rtt(detail):
    """Détermine si c'est un RTT"""
    if pd.isna(detail) or detail == '':
        return False
    return 'rtt' in str(detail).lower()


def get_status_code(type_am, type_pm, detail_am, detail_pm):
    """Retourne le code de statut avec gestion des demi-journées différentes"""
    # Week-end/Férié
    if type_am == 'JOUR_NON_OUVRE' and type_pm == 'JOUR_NON_OUVRE':
        return 'W'

    # Présent toute la journée
    if type_am == 'PRESENT' and type_pm == 'PRESENT':
        return ''

    # Fonction pour obtenir le code d'une demi-journée
    def get_half_day_code(type_val, detail_val):
        if type_val == 'PRESENT':
            return 'P'
        elif type_val == 'TELETRAVAIL':
            return 'TV' if is_validated(detail_val) else 'TP'
        elif type_val == 'CONGES':
            if is_rtt(detail_val):
                return 'RV' if is_validated(detail_val) else 'RP'
            else:
                return 'CV' if is_validated(detail_val) else 'CP'
        elif type_val == 'JOUR_NON_OUVRE':
            return 'W'
        return ''

    # Obtenir les codes pour AM et PM
    code_am = get_half_day_code(type_am, detail_am)
    code_pm = get_half_day_code(type_pm, detail_pm)

    # Journée complète identique
    if code_am == code_pm:
        # Si les deux sont vides (présent)
        if code_am == 'P':
            return ''
        # Sinon retourner le code sans précision
        return code_am

    # Demi-journées différentes
    # Si une des deux est présent, on affiche juste l'autre avec AM ou PM
    if code_am == 'P' and code_pm != '':
        return f'{code_pm}-PM'
    if code_pm == 'P' and code_am != '':
        return f'{code_am}-AM'

    # Si les deux sont des événements différents (pas présent)
    # On affiche les deux séparés par /
    if code_am != '' and code_pm != '':
        return f'{code_am}/{code_pm}'

    # Par défaut
    return code_am if code_am else code_pm


def analyze_leave_data(csv_file):
    """Analyse le fichier CSV et retourne les statistiques par personne"""

    df = pd.read_csv(csv_file)
    df['date_obj'] = pd.to_datetime(df['date'], format='%Y/%m/%d')

    stats = defaultdict(lambda: {
        'teletravail_valide_am': 0,
        'teletravail_valide_pm': 0,
        'teletravail_a_valider_am': 0,
        'teletravail_a_valider_pm': 0,
        'conges_valides_am': 0,
        'conges_valides_pm': 0,
        'conges_a_valider_am': 0,
        'conges_a_valider_pm': 0,
        'rtt_valides_am': 0,
        'rtt_valides_pm': 0,
        'rtt_a_valider_am': 0,
        'rtt_a_valider_pm': 0,
        'regle_10j_consecutifs': False,
        'regle_20j_total': False,
        'jours_consecutifs_max': 0,
        'jours_total_periode': 0,
    })

    # Période de référence : 15/05 au 15/10
    date_debut = datetime(2026, 5, 15)
    date_fin = datetime(2026, 10, 15)

    for _, row in df.iterrows():
        collaborateur = row['collaborateur']

        if row['type_am'] == 'TELETRAVAIL':
            validated = is_validated(row['detail_am'])
            if validated == True:
                stats[collaborateur]['teletravail_valide_am'] += 1
            elif validated == False:
                stats[collaborateur]['teletravail_a_valider_am'] += 1

        elif row['type_am'] == 'CONGES':
            if is_rtt(row['detail_am']):
                validated = is_validated(row['detail_am'])
                if validated == True:
                    stats[collaborateur]['rtt_valides_am'] += 1
                elif validated == False:
                    stats[collaborateur]['rtt_a_valider_am'] += 1
            else:
                validated = is_validated(row['detail_am'])
                if validated == True:
                    stats[collaborateur]['conges_valides_am'] += 1
                elif validated == False:
                    stats[collaborateur]['conges_a_valider_am'] += 1

        if row['type_pm'] == 'TELETRAVAIL':
            validated = is_validated(row['detail_pm'])
            if validated == True:
                stats[collaborateur]['teletravail_valide_pm'] += 1
            elif validated == False:
                stats[collaborateur]['teletravail_a_valider_pm'] += 1

        elif row['type_pm'] == 'CONGES':
            if is_rtt(row['detail_pm']):
                validated = is_validated(row['detail_pm'])
                if validated == True:
                    stats[collaborateur]['rtt_valides_pm'] += 1
                elif validated == False:
                    stats[collaborateur]['rtt_a_valider_pm'] += 1
            else:
                validated = is_validated(row['detail_pm'])
                if validated == True:
                    stats[collaborateur]['conges_valides_pm'] += 1
                elif validated == False:
                    stats[collaborateur]['conges_a_valider_pm'] += 1

    # Analyser les règles RH pour chaque collaborateur
    for collaborateur in stats.keys():
        collab_df = df[df['collaborateur'] == collaborateur].copy()
        collab_df = collab_df[(collab_df['date_obj'] >= date_debut) & (collab_df['date_obj'] <= date_fin)]
        collab_df = collab_df.sort_values('date_obj')

        # Dictionnaire des jours : True si congés/RTT, False sinon
        jours_absence = {}
        total_jours = 0

        for _, row in collab_df.iterrows():
            date = row['date_obj']

            # Vérifier si c'est un jour de congés ou RTT (AM ou PM)
            is_absence_am = row['type_am'] == 'CONGES'
            is_absence_pm = row['type_pm'] == 'CONGES'

            # Compter comme jour d'absence si au moins une demi-journée
            if is_absence_am or is_absence_pm:
                jours_absence[date] = True
                # Compter en jours : 0.5 si une seule demi-journée, 1 si les deux
                if is_absence_am and is_absence_pm:
                    total_jours += 1
                else:
                    total_jours += 0.5
            else:
                jours_absence[date] = False

        # Calculer la plus longue séquence de jours consécutifs
        max_consecutifs = 0
        current_consecutifs = 0

        # Parcourir toutes les dates de la période
        current_date = date_debut
        while current_date <= date_fin:
            if current_date in jours_absence and jours_absence[current_date]:
                current_consecutifs += 1
                max_consecutifs = max(max_consecutifs, current_consecutifs)
            else:
                current_consecutifs = 0
            current_date += pd.Timedelta(days=1)

        # Enregistrer les résultats
        stats[collaborateur]['jours_consecutifs_max'] = max_consecutifs
        stats[collaborateur]['jours_total_periode'] = total_jours
        stats[collaborateur]['regle_10j_consecutifs'] = max_consecutifs >= 10
        stats[collaborateur]['regle_20j_total'] = total_jours >= 20

    return stats


def create_summary_sheet(wb, stats):
    """Crée la feuille de synthèse avec les compteurs"""

    ws = wb.active
    ws.title = "Synthèse"

    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    subheader_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    headers = [
        "Collaborateur",
        "Télétravail\nValidé (j)",
        "Télétravail\nÀ valider (j)",
        "Congés\nValidés (j)",
        "Congés\nÀ valider (j)",
        "RTT\nValidés (j)",
        "RTT\nÀ valider (j)",
        "Règle 10j\nconsécutifs",
        "Règle 20j\ntotal"
    ]

    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border

    ws.row_dimensions[1].height = 30

    sorted_collaborateurs = sorted(stats.keys())

    row = 2
    for collaborateur in sorted_collaborateurs:
        s = stats[collaborateur]

        teletravail_valide = (s['teletravail_valide_am'] + s['teletravail_valide_pm']) / 2
        teletravail_a_valider = (s['teletravail_a_valider_am'] + s['teletravail_a_valider_pm']) / 2
        conges_valides = (s['conges_valides_am'] + s['conges_valides_pm']) / 2
        conges_a_valider = (s['conges_a_valider_am'] + s['conges_a_valider_pm']) / 2
        rtt_valides = (s['rtt_valides_am'] + s['rtt_valides_pm']) / 2
        rtt_a_valider = (s['rtt_a_valider_am'] + s['rtt_a_valider_pm']) / 2

        # Règles RH
        regle_10j = "✓" if s['regle_10j_consecutifs'] else "✗"
        regle_20j = "✓" if s['regle_20j_total'] else "✗"

        data = [
            collaborateur,
            teletravail_valide,
            teletravail_a_valider,
            conges_valides,
            conges_a_valider,
            rtt_valides,
            rtt_a_valider,
            regle_10j,
            regle_20j
        ]

        for col, value in enumerate(data, 1):
            cell = ws.cell(row=row, column=col, value=value)
            cell.border = border

            if col > 1 and col <= 7:
                cell.alignment = Alignment(horizontal='right', vertical='center')
                if isinstance(value, (int, float)):
                    cell.number_format = '0.0'
            elif col > 7:
                # Colonnes des règles RH
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(bold=True, size=14)
                # Couleur verte si ✓, rouge si ✗
                if value == "✓":
                    cell.fill = green_fill
                else:
                    cell.fill = red_fill
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')

        row += 1

    # Ligne de totaux (seulement pour les compteurs, pas les règles)
    ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=row, column=1).fill = subheader_fill

    for col in range(2, 8):  # Colonnes 2 à 7 (télétravail, congés, RTT)
        col_letter = get_column_letter(col)
        formula = f"=SUM({col_letter}2:{col_letter}{row - 1})"
        cell = ws.cell(row=row, column=col, value=formula)
        cell.font = Font(bold=True)
        cell.fill = subheader_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = '0.0'

    # Colonnes des règles RH - afficher un résumé
    cell = ws.cell(row=row, column=8)
    nb_ok_10j = sum(1 for s in stats.values() if s['regle_10j_consecutifs'])
    cell.value = f"{nb_ok_10j}/{len(stats)}"
    cell.font = Font(bold=True)
    cell.fill = subheader_fill
    cell.border = border
    cell.alignment = Alignment(horizontal='center', vertical='center')

    cell = ws.cell(row=row, column=9)
    nb_ok_20j = sum(1 for s in stats.values() if s['regle_20j_total'])
    cell.value = f"{nb_ok_20j}/{len(stats)}"
    cell.font = Font(bold=True)
    cell.fill = subheader_fill
    cell.border = border
    cell.alignment = Alignment(horizontal='center', vertical='center')

    # Ajuster les largeurs
    ws.column_dimensions['A'].width = 25
    for col in range(2, 8):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['I'].width = 12

    # Note explicative
    ws.merge_cells(f'A{row + 2}:I{row + 2}')
    note_cell = ws[f'A{row + 2}']
    note_cell.value = "📋 Règles RH (période 15/05 - 15/10) : 10j consécutifs = au moins 10 jours d'affilée | 20j total = au moins 20 jours (consécutifs ou non)"
    note_cell.font = Font(size=9, italic=True)
    note_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[row + 2].height = 30

    ws.freeze_panes = 'A2'


def create_dynamic_calendar_sheet(wb, csv_file):
    """Crée la feuille de planning dynamique avec formules Excel simplifiées"""

    df = pd.read_csv(csv_file)
    df['date_obj'] = pd.to_datetime(df['date'], format='%Y/%m/%d')

    collaborateurs = sorted(df['collaborateur'].unique())

    ws = wb.create_sheet("Planning Calendrier")

    # Styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    month_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    month_font = Font(bold=True, size=10)
    dropdown_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Couleurs pour formatage conditionnel
    teletravail_fill = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
    teletravail_pending_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    conges_fill = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
    conges_pending_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    rtt_fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
    rtt_pending_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    weekend_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Titre
    ws.merge_cells('A1:AF1')
    title_cell = ws['A1']
    title_cell.value = "PLANNING ANNUEL 2026 - DYNAMIQUE"
    title_cell.font = Font(bold=True, size=14, color="FFFFFF")
    title_cell.fill = header_fill
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 25

    # Sélecteur de collaborateur
    ws['A2'] = "Collaborateur :"
    ws['A2'].font = Font(bold=True, size=11)
    ws['A2'].alignment = Alignment(horizontal='right', vertical='center')

    ws.merge_cells('B2:D2')
    dropdown_cell = ws['B2']
    dropdown_cell.fill = dropdown_fill
    dropdown_cell.alignment = Alignment(horizontal='left', vertical='center')
    dropdown_cell.value = collaborateurs[0]

    # Créer la validation de données
    dv = DataValidation(type="list", formula1=f'"{",".join(collaborateurs)}"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add('B2')

    ws.row_dimensions[2].height = 25

    # Légende améliorée - Plus visuelle
    ws['A3'] = "LÉGENDE"
    ws['A3'].font = Font(bold=True, size=12)
    ws['A3'].alignment = Alignment(horizontal='center', vertical='center')

    # Couleur mixte pour journées avec deux types différents
    mixed_fill_legend = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")

    # Ligne 3 : Télétravail et Congés
    legend_row3 = [
        ("B3:C3", "TV", "Télétravail validé", teletravail_fill, "FFFFFF"),
        ("D3:E3", "TP", "Télétravail à valider", teletravail_pending_fill, "000000"),
        ("F3:G3", "CV", "Congés validés", conges_fill, "FFFFFF"),
        ("H3:I3", "CP", "Congés à valider", conges_pending_fill, "000000"),
    ]

    # Ligne 4 : RTT et autres
    ws['A4'] = ""
    legend_row4 = [
        ("B4:C4", "RV", "RTT validés", rtt_fill, "FFFFFF"),
        ("D4:E4", "RP", "RTT à valider", rtt_pending_fill, "000000"),
        ("F4:G4", "W", "Week-end/Férié", weekend_fill, "000000"),
        ("H4:I4", "CV/TV", "Journée mixte", mixed_fill_legend, "000000"),
    ]

    # Appliquer les styles pour la ligne 3
    for cell_range, code, description, fill, font_color in legend_row3:
        # D'abord accéder à la cellule
        start_cell = cell_range.split(':')[0]
        cell = ws[start_cell]
        cell.value = f"{code}\n{description}"
        cell.fill = fill
        cell.font = Font(bold=True, size=9, color=font_color)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
        # Puis fusionner les cellules
        ws.merge_cells(cell_range)

    # Appliquer les styles pour la ligne 4
    for cell_range, code, description, fill, font_color in legend_row4:
        # D'abord accéder à la cellule
        start_cell = cell_range.split(':')[0]
        cell = ws[start_cell]
        cell.value = f"{code}\n{description}"
        cell.fill = fill
        cell.font = Font(bold=True, size=9, color=font_color)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
        # Puis fusionner les cellules
        ws.merge_cells(cell_range)

    # Ajuster la hauteur des lignes de légende
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[4].height = 30

    # Ligne 5 : Explications des formats spéciaux
    ws.merge_cells('A5:I5')
    explanation_cell = ws['A5']
    explanation_cell.value = "💡 Formats spéciaux : Code-AM/PM = demi-journée | Code1/Code2 = matin≠après-midi | Cellule barrée = jour inexistant"
    explanation_cell.font = Font(size=9, italic=True)
    explanation_cell.alignment = Alignment(horizontal='left', vertical='center')
    explanation_cell.fill = PatternFill(start_color="FFF9E6", end_color="FFF9E6", fill_type="solid")
    ws.row_dimensions[5].height = 20

    # En-têtes : Mois + jours 1-31
    ws['A7'] = "Mois"
    ws['A7'].fill = header_fill
    ws['A7'].font = header_font
    ws['A7'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A7'].border = border

    for day in range(1, 32):
        col = day + 1
        cell = ws.cell(row=7, column=col, value=day)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    ws.row_dimensions[7].height = 20

    # Ajuster les largeurs
    ws.column_dimensions['A'].width = 12
    for col in range(2, 33):
        ws.column_dimensions[get_column_letter(col)].width = 5  # Augmenté pour lisibilité

    # Créer une feuille cachée avec les données pré-calculées
    print("  📝 Préparation des données...")
    data_ws = wb.create_sheet("_Lookup")
    data_ws.sheet_state = 'hidden'

    # En-têtes
    data_ws['A1'] = 'Collaborateur'
    data_ws['B1'] = 'Mois'
    data_ws['C1'] = 'Jour'
    data_ws['D1'] = 'Code'

    for cell in data_ws['1:1']:
        cell.font = Font(bold=True)

    # Remplir les données pré-calculées
    row_idx = 2
    for _, row in df.iterrows():
        month = row['date_obj'].month
        day = row['date_obj'].day
        code = get_status_code(row['type_am'], row['type_pm'], row['detail_am'], row['detail_pm'])

        data_ws.cell(row=row_idx, column=1, value=row['collaborateur'])
        data_ws.cell(row=row_idx, column=2, value=month)
        data_ws.cell(row=row_idx, column=3, value=day)
        data_ws.cell(row=row_idx, column=4, value=code)

        row_idx += 1

    # Mois de l'année avec leur nombre de jours réels
    mois_noms = [
        "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
    ]

    # Nombre de jours par mois (2026 n'est pas bissextile)
    jours_par_mois = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    # Style pour les jours inexistants
    border_diagonal = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
        diagonal=Side(style='thin', color='FF0000'),  # Ligne rouge
        diagonalDown=True
    )
    inexistant_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # Gris très clair

    # Créer les lignes pour chaque mois avec formules VLOOKUP simples
    print("  🔧 Création des formules...")
    for month_num, month_name in enumerate(mois_noms, 1):
        row_num = 8 + month_num - 1  # Commence à la ligne 8 (au lieu de 6)
        max_days_in_month = jours_par_mois[month_num - 1]

        # Nom du mois
        cell = ws.cell(row=row_num, column=1, value=month_name)
        cell.fill = month_fill
        cell.font = month_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

        # Formules pour chaque jour du mois
        for day in range(1, 32):
            col_num = day + 1
            cell = ws.cell(row=row_num, column=col_num)

            # Vérifier si ce jour existe dans ce mois
            if day <= max_days_in_month:
                # Jour valide - formule normale
                cell.border = border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(size=6, bold=True)  # Taille réduite pour "CV/TV"

                # Formule pour afficher le code (qui peut contenir AM, PM ou /)
                formula = f'=IFERROR(INDEX(_Lookup!$D:$D,MATCH($B$2&{month_num}&{day},_Lookup!$A:$A&_Lookup!$B:$B&_Lookup!$C:$C,0)),"")'
                cell.value = formula
            else:
                # Jour inexistant - style avec diagonale
                cell.border = border_diagonal
                cell.fill = inexistant_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # Ajouter le formatage conditionnel
    print("  🎨 Application du formatage conditionnel...")

    # Couleur mixte pour les journées avec deux types différents
    mixed_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")  # Gris clair

    # Zone à formater : B8:AF19 (12 mois x 31 jours) - uniquement jours valides
    for month_num in range(1, 13):
        row_num = 8 + month_num - 1  # Commence à la ligne 8
        max_days_in_month = jours_par_mois[month_num - 1]

        # Appliquer le formatage uniquement aux jours existants
        for day in range(1, max_days_in_month + 1):
            col_letter = get_column_letter(day + 1)
            cell_ref = f"{col_letter}{row_num}"

            # D'abord, détecter les journées mixtes (avec /)
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'ISNUMBER(SEARCH("/",{cell_ref}))'], fill=mixed_fill))

            # Ensuite les codes spécifiques
            # Week-end (priorité haute car exact)
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'{cell_ref}="W"'], fill=weekend_fill))

            # Télétravail validé
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'OR(LEFT({cell_ref},2)="TV",{cell_ref}="TV")'],
                                                      fill=teletravail_fill))

            # Télétravail à valider
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'OR(LEFT({cell_ref},2)="TP",{cell_ref}="TP")'],
                                                      fill=teletravail_pending_fill))

            # Congés validés
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'OR(LEFT({cell_ref},2)="CV",{cell_ref}="CV")'],
                                                      fill=conges_fill))

            # Congés à valider
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'OR(LEFT({cell_ref},2)="CP",{cell_ref}="CP")'],
                                                      fill=conges_pending_fill))

            # RTT validés
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'OR(LEFT({cell_ref},2)="RV",{cell_ref}="RV")'],
                                                      fill=rtt_fill))

            # RTT à valider
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'OR(LEFT({cell_ref},2)="RP",{cell_ref}="RP")'],
                                                      fill=rtt_pending_fill))

    ws.freeze_panes = 'B8'  # Geler jusqu'à la colonne A et ligne 7

    print("  ✓ Planning dynamique créé avec succès!")


def create_excel_report(stats, csv_file, output_file):
    """Crée le fichier Excel complet"""

    wb = openpyxl.Workbook()

    print("📊 Création de la feuille Synthèse...")
    create_summary_sheet(wb, stats)

    print("📅 Création de la feuille Planning Dynamique...")
    create_dynamic_calendar_sheet(wb, csv_file)

    wb.save(output_file)
    print(f"✅ Fichier Excel créé : {output_file}")


def main():
    """Fonction principale"""

    input_file = "leave_planning_2026_complete.csv"
    output_file = "compteurs_conges_2026_dynamique_final.xlsx"

    print("🔍 Analyse du fichier CSV...")
    stats = analyze_leave_data(input_file)

    print(f"📊 Nombre de collaborateurs : {len(stats)}")

    print("📝 Génération du fichier Excel...")
    create_excel_report(stats, input_file, output_file)

    print("\n✨ Terminé !")
    print("\nContenu du fichier :")
    print("  - Feuille 'Synthèse' : compteurs par personne")
    print("  - Feuille 'Planning Calendrier' : vue dynamique avec liste déroulante")
    print("\n💡 Utilisation :")
    print("  Sélectionnez un collaborateur dans la liste déroulante (cellule B2)")
    print("  Le calendrier se met automatiquement à jour avec les bonnes couleurs!")
    print("\n📋 Codes affichés dans les cellules :")
    print("  TV = Télétravail Validé          TP = Télétravail à valider")
    print("  CV = Congés Validés              CP = Congés à valider")
    print("  RV = RTT Validés                 RP = RTT à valider")
    print("  W  = Week-end/Férié              P  = Présent")
    print("\n  📌 Formats spéciaux :")
    print("  • Code-AM ou Code-PM → Seule cette demi-journée est concernée")
    print("    Exemple: CV-AM = Congés le matin, présent l'après-midi")
    print("\n  • Code1/Code2 → Matin et après-midi avec événements différents")
    print("    Exemple: CV/TV = Congés le matin, télétravail l'après-midi")
    print("            (couleur gris clair pour ces journées mixtes)")
    print("\n  • Cellules barrées en diagonale = Jours inexistants dans le mois")
    print("    Exemple: 30 et 31 février, 31 avril, etc.")
    print("\n✅ Compatible Excel 365 ET Google Sheets !")


if __name__ == "__main__":
    main()