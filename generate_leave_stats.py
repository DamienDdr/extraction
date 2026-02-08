#!/usr/bin/env python3
"""
Script pour générer des statistiques de congés par personne à partir du CSV de planning
Avec planning annuel en format matriciel (mois x jours)
"""

import pandas as pd
from collections import defaultdict
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


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


def get_day_type(type_am, type_pm, detail_am, detail_pm):
    """
    Détermine le type de journée en fonction de AM et PM
    Retourne: (type, validated, text_display)
    """
    # Si jour non ouvré
    if type_am == 'JOUR_NON_OUVRE' and type_pm == 'JOUR_NON_OUVRE':
        return ('JOUR_NON_OUVRE', None, '')

    # Si présent toute la journée
    if type_am == 'PRESENT' and type_pm == 'PRESENT':
        return ('PRESENT', None, '')

    # Journée complète identique
    if type_am == type_pm and type_am in ['TELETRAVAIL', 'CONGES']:
        if type_am == 'TELETRAVAIL':
            validated = is_validated(detail_am)
            return ('TELETRAVAIL', validated, '')
        elif type_am == 'CONGES':
            if is_rtt(detail_am):
                validated = is_validated(detail_am)
                return ('RTT', validated, '')
            else:
                validated = is_validated(detail_am)
                return ('CONGES', validated, '')

    # Demi-journées différentes - on priorise dans l'ordre: CONGES > RTT > TELETRAVAIL > PRESENT
    types_priority = []

    # Analyser AM
    if type_am == 'CONGES':
        if is_rtt(detail_am):
            types_priority.append(('RTT', is_validated(detail_am), 'AM'))
        else:
            types_priority.append(('CONGES', is_validated(detail_am), 'AM'))
    elif type_am == 'TELETRAVAIL':
        types_priority.append(('TELETRAVAIL', is_validated(detail_am), 'AM'))

    # Analyser PM
    if type_pm == 'CONGES':
        if is_rtt(detail_pm):
            types_priority.append(('RTT', is_validated(detail_pm), 'PM'))
        else:
            types_priority.append(('CONGES', is_validated(detail_pm), 'PM'))
    elif type_pm == 'TELETRAVAIL':
        types_priority.append(('TELETRAVAIL', is_validated(detail_pm), 'PM'))

    # Si on a des types différents, on prend le plus prioritaire
    if types_priority:
        # Prioriser CONGES > RTT > TELETRAVAIL
        for ptype in ['CONGES', 'RTT', 'TELETRAVAIL']:
            for t, v, period in types_priority:
                if t == ptype:
                    # Si demi-journée, afficher AM ou PM
                    text = period if len(types_priority) > 1 or (type_am == 'PRESENT' or type_pm == 'PRESENT') else ''
                    return (t, v, text)

    return ('PRESENT', None, '')


def analyze_leave_data(csv_file):
    """Analyse le fichier CSV et retourne les statistiques par personne"""

    df = pd.read_csv(csv_file)

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
    })

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

    return stats


def create_summary_sheet(wb, stats):
    """Crée la feuille de synthèse avec les compteurs"""

    ws = wb.active
    ws.title = "Synthèse"

    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    subheader_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
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
        "Total\nValidé (j)",
        "Total\nÀ valider (j)",
        "TOTAL\nGÉNÉRAL (j)"
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

        total_valide = teletravail_valide + conges_valides + rtt_valides
        total_a_valider = teletravail_a_valider + conges_a_valider + rtt_a_valider
        total_general = total_valide + total_a_valider

        data = [
            collaborateur,
            teletravail_valide,
            teletravail_a_valider,
            conges_valides,
            conges_a_valider,
            rtt_valides,
            rtt_a_valider,
            total_valide,
            total_a_valider,
            total_general
        ]

        for col, value in enumerate(data, 1):
            cell = ws.cell(row=row, column=col, value=value)
            cell.border = border

            if col > 1:
                cell.alignment = Alignment(horizontal='right', vertical='center')
                if isinstance(value, (int, float)):
                    cell.number_format = '0.0'
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')

        row += 1

    ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=row, column=1).fill = subheader_fill

    for col in range(2, 11):
        col_letter = get_column_letter(col)
        formula = f"=SUM({col_letter}2:{col_letter}{row - 1})"
        cell = ws.cell(row=row, column=col, value=formula)
        cell.font = Font(bold=True)
        cell.fill = subheader_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = '0.0'

    ws.column_dimensions['A'].width = 25
    for col in range(2, 11):
        ws.column_dimensions[get_column_letter(col)].width = 14

    ws.freeze_panes = 'A2'


def create_calendar_planning_sheet_for_collaborator(wb, df, collaborateur, collab_num, total_collabs):
    """Crée une feuille de planning calendrier pour un collaborateur spécifique"""

    # Créer un nom de feuille court (limite de 31 caractères)
    sheet_name = collaborateur[:28] if len(collaborateur) > 28 else collaborateur
    ws = wb.create_sheet(sheet_name)

    # Styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    month_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    month_font = Font(bold=True, size=10)

    # Couleurs pour les types
    teletravail_fill = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")  # Bleu
    teletravail_pending_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")  # Bleu clair
    conges_fill = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")  # Vert
    conges_pending_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")  # Vert clair
    rtt_fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")  # Orange
    rtt_pending_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")  # Orange clair
    weekend_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")  # Gris

    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Titre
    ws.merge_cells('A1:AF1')
    title_cell = ws['A1']
    title_cell.value = f"PLANNING ANNUEL 2026 - {collaborateur}"
    title_cell.font = Font(bold=True, size=14, color="FFFFFF")
    title_cell.fill = header_fill
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 25

    # Légende
    ws['A3'] = "Légende :"
    ws['A3'].font = Font(bold=True)

    legend_items = [
        ("B3", "Télétravail ✓", teletravail_fill),
        ("D3", "Télétravail (à valider)", teletravail_pending_fill),
        ("G3", "Congés ✓", conges_fill),
        ("I3", "Congés (à valider)", conges_pending_fill),
        ("L3", "RTT ✓", rtt_fill),
        ("N3", "RTT (à valider)", rtt_pending_fill),
        ("Q3", "Week-end/Férié", weekend_fill),
        ("S3", "Présent = blanc", None),
    ]

    for cell_ref, text, fill in legend_items:
        cell = ws[cell_ref]
        cell.value = text
        if fill:
            cell.fill = fill
        cell.font = Font(size=9)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    # En-têtes : Mois + jours 1-31
    ws['A5'] = "Mois"
    ws['A5'].fill = header_fill
    ws['A5'].font = header_font
    ws['A5'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A5'].border = border

    for day in range(1, 32):
        col = day + 1  # Colonne B = jour 1
        cell = ws.cell(row=5, column=col, value=day)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    ws.row_dimensions[5].height = 20

    # Ajuster les largeurs
    ws.column_dimensions['A'].width = 12
    for col in range(2, 33):
        ws.column_dimensions[get_column_letter(col)].width = 4

    # Mois de l'année
    mois_noms = [
        "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
    ]

    # Récupérer les données du collaborateur
    collab_df = df[df['collaborateur'] == collaborateur].copy()
    collab_data = {}

    for _, row in collab_df.iterrows():
        date_obj = row['date_obj']
        month = date_obj.month
        day = date_obj.day

        if month not in collab_data:
            collab_data[month] = {}

        day_type, validated, text = get_day_type(
            row['type_am'], row['type_pm'],
            row['detail_am'], row['detail_pm']
        )

        collab_data[month][day] = (day_type, validated, text)

    # Remplir le planning
    for month_num, month_name in enumerate(mois_noms, 1):
        row_num = 6 + month_num - 1

        # Nom du mois
        cell = ws.cell(row=row_num, column=1, value=month_name)
        cell.fill = month_fill
        cell.font = month_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

        # Jours du mois
        for day in range(1, 32):
            col_num = day + 1
            cell = ws.cell(row=row_num, column=col_num)
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')

            # Vérifier si ce jour existe pour ce mois
            if month_num in collab_data and day in collab_data[month_num]:
                day_type, validated, text = collab_data[month_num][day]

                # Appliquer la couleur
                if day_type == 'JOUR_NON_OUVRE':
                    cell.fill = weekend_fill
                elif day_type == 'TELETRAVAIL':
                    cell.fill = teletravail_fill if validated else teletravail_pending_fill
                    cell.value = text if text else ''
                    cell.font = Font(size=8)
                elif day_type == 'CONGES':
                    cell.fill = conges_fill if validated else conges_pending_fill
                    cell.value = text if text else ''
                    cell.font = Font(size=8)
                elif day_type == 'RTT':
                    cell.fill = rtt_fill if validated else rtt_pending_fill
                    cell.value = text if text else ''
                    cell.font = Font(size=8)
                # PRESENT reste blanc (pas de fill)
            else:
                # Jour n'existe pas pour ce mois (ex: 31 février)
                cell.fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")

    ws.freeze_panes = 'B6'

    print(f"  ✓ Feuille créée pour {collaborateur} ({collab_num}/{total_collabs})")


def create_calendar_planning_sheets(wb, csv_file):
    """Crée une feuille de planning pour chaque collaborateur"""

    df = pd.read_csv(csv_file)
    df['date_obj'] = pd.to_datetime(df['date'], format='%Y/%m/%d')

    collaborateurs = sorted(df['collaborateur'].unique())
    total = len(collaborateurs)

    print(f"📅 Création de {total} feuilles de planning calendrier...")

    for idx, collaborateur in enumerate(collaborateurs, 1):
        create_calendar_planning_sheet_for_collaborator(wb, df, collaborateur, idx, total)


def create_excel_report(stats, csv_file, output_file):
    """Crée le fichier Excel complet"""

    wb = openpyxl.Workbook()

    print("📊 Création de la feuille Synthèse...")
    create_summary_sheet(wb, stats)

    print("📅 Création des feuilles de planning calendrier...")
    create_calendar_planning_sheets(wb, csv_file)

    wb.save(output_file)
    print(f"✅ Fichier Excel créé : {output_file}")


def main():
    """Fonction principale"""

    input_file = "leave_planning_2026_complete.csv"
    output_file = "compteurs_conges_2026_calendrier.xlsx"

    print("🔍 Analyse du fichier CSV...")
    stats = analyze_leave_data(input_file)

    print(f"📊 Nombre de collaborateurs : {len(stats)}")

    print("📝 Génération du fichier Excel...")
    create_excel_report(stats, input_file, output_file)

    print("\n✨ Terminé !")
    print("\nContenu du fichier :")
    print("  - Feuille 'Synthèse' : compteurs par personne")
    print("  - Feuilles individuelles : planning annuel calendrier pour chaque collaborateur")


if __name__ == "__main__":
    main()