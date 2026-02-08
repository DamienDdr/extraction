#!/usr/bin/env python3
"""
Script pour générer des statistiques de congés par personne à partir du CSV de planning
Avec planning annuel dynamique (liste déroulante + formules Excel)
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


def create_dynamic_calendar_sheet(wb, csv_file):
    """Crée la feuille de planning dynamique avec formules Excel"""

    df = pd.read_csv(csv_file)
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
    ws['A3'] = "Collaborateur :"
    ws['A3'].font = Font(bold=True, size=11)
    ws['A3'].alignment = Alignment(horizontal='right', vertical='center')

    ws.merge_cells('B3:D3')
    dropdown_cell = ws['B3']
    dropdown_cell.fill = dropdown_fill
    dropdown_cell.alignment = Alignment(horizontal='left', vertical='center')
    dropdown_cell.value = collaborateurs[0]

    # Créer la validation de données
    dv = DataValidation(type="list", formula1=f'"{",".join(collaborateurs)}"', allow_blank=False)
    dv.error = 'Veuillez sélectionner un collaborateur dans la liste'
    dv.errorTitle = 'Entrée invalide'
    ws.add_data_validation(dv)
    dv.add(dropdown_cell)

    # Légende
    ws['F3'] = "Légende :"
    ws['F3'].font = Font(bold=True)

    legend_items = [
        ("G3", "Télétravail ✓", teletravail_fill),
        ("I3", "Télétravail à val.", teletravail_pending_fill),
        ("L3", "Congés ✓", conges_fill),
        ("N3", "Congés à val.", conges_pending_fill),
        ("Q3", "RTT ✓", rtt_fill),
        ("S3", "RTT à val.", rtt_pending_fill),
        ("V3", "Week-end/Férié", weekend_fill),
    ]

    for cell_ref, text, fill in legend_items:
        cell = ws[cell_ref]
        cell.value = text
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
        col = day + 1
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

    # Créer une feuille cachée avec toutes les données
    data_ws = wb.create_sheet("_Données")
    data_ws.sheet_state = 'hidden'

    # Copier les données dans la feuille cachée
    for r_idx, (idx, row) in enumerate(df.iterrows(), 2):
        data_ws.cell(row=r_idx, column=1, value=row['collaborateur'])
        data_ws.cell(row=r_idx, column=2, value=row['date'])
        data_ws.cell(row=r_idx, column=3, value=row['type_am'])
        data_ws.cell(row=r_idx, column=4, value=row['detail_am'])
        data_ws.cell(row=r_idx, column=5, value=row['type_pm'])
        data_ws.cell(row=r_idx, column=6, value=row['detail_pm'])

    # En-têtes de la feuille de données
    headers_data = ['collaborateur', 'date', 'type_am', 'detail_am', 'type_pm', 'detail_pm']
    for c_idx, header in enumerate(headers_data, 1):
        data_ws.cell(row=1, column=c_idx, value=header)
        data_ws.cell(row=1, column=c_idx).font = Font(bold=True)

    # Mois de l'année
    mois_noms = [
        "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
    ]

    # Créer les lignes pour chaque mois
    for month_num, month_name in enumerate(mois_noms, 1):
        row_num = 6 + month_num - 1

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
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(size=8)

            # Formule complexe pour récupérer le statut du jour
            # Format de date : 2026/MM/DD
            date_str = f"2026/{month_num:02d}/{day:02d}"

            # Formule SUMPRODUCT pour trouver la ligne correspondante
            formula = f'''=IFERROR(
IF(
    AND(_Données!$C:$C<>"",_Données!$A:$A=$B$3,_Données!$B:$B="{date_str}"),
    IF(
        AND(INDEX(_Données!$C:$C,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))="JOUR_NON_OUVRE"),
        "W",
        IF(
            AND(INDEX(_Données!$C:$C,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))="TELETRAVAIL",INDEX(_Données!$E:$E,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))="TELETRAVAIL"),
            IF(OR(ISNUMBER(SEARCH("Validé",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0)))),ISNUMBER(SEARCH("validé",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))))),"TV","TP"),
            IF(
                AND(INDEX(_Données!$C:$C,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))="CONGES",INDEX(_Données!$E:$E,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))="CONGES"),
                IF(OR(ISNUMBER(SEARCH("RTT",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0)))),ISNUMBER(SEARCH("rtt",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))))),
                    IF(OR(ISNUMBER(SEARCH("Validé",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0)))),ISNUMBER(SEARCH("validé",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))))),"RV","RP"),
                    IF(OR(ISNUMBER(SEARCH("Validé",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0)))),ISNUMBER(SEARCH("validé",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))))),"CV","CP")
                ),
                ""
            )
        )
    ),
    ""
),"")'''

            # Formule simplifiée avec codes
            # TV = Télétravail Validé, TP = Télétravail Pending
            # CV = Congés Validé, CP = Congés Pending
            # RV = RTT Validé, RP = RTT Pending
            # W = Weekend/Férié

            # Version simplifiée de la formule (à cause de la complexité, on va utiliser une approche différente)
            # On va mettre une formule qui retourne juste le type
            cell.value = ''  # On va remplir avec du code Python

    # Remplir avec Python au lieu de formules Excel trop complexes
    print("  📝 Remplissage des données du calendrier...")
    for month_num in range(1, 13):
        row_num = 6 + month_num - 1
        for day in range(1, 32):
            col_num = day + 1
            cell = ws.cell(row=row_num, column=col_num)

            # Créer une formule qui retourne un code
            date_str = f"2026/{month_num:02d}/{day:02d}"

            # Formule pour récupérer le type_am et type_pm
            # Si les deux sont identiques et valides, on affiche le code correspondant
            formula = f'=IFERROR(IF(COUNTIFS(_Données!$A:$A,$B$3,_Données!$B:$B,"{date_str}")>0,' \
                      f'IF(AND(INDEX(_Données!$C:$C,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))="JOUR_NON_OUVRE"),"W",' \
                      f'IF(INDEX(_Données!$C:$C,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))="TELETRAVAIL",' \
                      f'IF(ISNUMBER(SEARCH("Validé",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0)))),"TV","TP"),' \
                      f'IF(INDEX(_Données!$C:$C,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0))="CONGES",' \
                      f'IF(ISNUMBER(SEARCH("RTT",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0)))),' \
                      f'IF(ISNUMBER(SEARCH("Validé",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0)))),"RV","RP"),' \
                      f'IF(ISNUMBER(SEARCH("Validé",INDEX(_Données!$D:$D,MATCH(1,(_Données!$A:$A=$B$3)*(_Données!$B:$B="{date_str}"),0)))),"CV","CP")),"")),"X"),"")'

            cell.value = formula

    # Ajouter le formatage conditionnel
    print("  🎨 Application du formatage conditionnel...")

    # Zone à formater : B6:AF17 (12 mois x 31 jours)
    for month_num in range(1, 13):
        row_num = 6 + month_num - 1
        for day in range(1, 32):
            col_letter = get_column_letter(day + 1)
            cell_ref = f"{col_letter}{row_num}"

            # Télétravail validé
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'{cell_ref}="TV"'], fill=teletravail_fill))

            # Télétravail à valider
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'{cell_ref}="TP"'], fill=teletravail_pending_fill))

            # Congés validés
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'{cell_ref}="CV"'], fill=conges_fill))

            # Congés à valider
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'{cell_ref}="CP"'], fill=conges_pending_fill))

            # RTT validés
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'{cell_ref}="RV"'], fill=rtt_fill))

            # RTT à valider
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'{cell_ref}="RP"'], fill=rtt_pending_fill))

            # Week-end/Férié
            ws.conditional_formatting.add(cell_ref,
                                          FormulaRule(formula=[f'{cell_ref}="W"'], fill=weekend_fill))

    ws.freeze_panes = 'B6'

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
    output_file = "compteurs_conges_2026_dynamique.xlsx"

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
    print("  Sélectionnez un collaborateur dans la liste déroulante (cellule B3)")
    print("  Le calendrier se met automatiquement à jour avec les bonnes couleurs!")


if __name__ == "__main__":
    main()