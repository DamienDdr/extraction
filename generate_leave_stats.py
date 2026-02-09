#!/usr/bin/env python3
"""
Version ULTRA-COMPATIBLE pour Excel 365 Entreprise
Une feuille par collaborateur - AUCUNE formule complexe
"""

import pandas as pd
from collections import defaultdict
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


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
    """Retourne le code de statut avec gestion des demi-journées"""
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
        if code_am == 'P':
            return ''
        return code_am

    # Demi-journées différentes
    if code_am == 'P' and code_pm != '':
        return f'{code_pm}-PM'
    if code_pm == 'P' and code_am != '':
        return f'{code_am}-AM'

    # Si les deux sont des événements différents
    if code_am != '' and code_pm != '':
        return f'{code_am}/{code_pm}'

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

        jours_type = {}
        total_jours = 0

        for _, row in collab_df.iterrows():
            date = row['date_obj']

            is_conges_am = row['type_am'] == 'CONGES'
            is_conges_pm = row['type_pm'] == 'CONGES'
            is_weekend = row['type_am'] == 'JOUR_NON_OUVRE' and row['type_pm'] == 'JOUR_NON_OUVRE'

            if is_conges_am or is_conges_pm:
                jours_type[date] = 'CONGES'
                if is_conges_am and is_conges_pm:
                    total_jours += 1
                else:
                    total_jours += 0.5
            elif is_weekend:
                jours_type[date] = 'WEEKEND'
            else:
                jours_type[date] = 'AUTRE'

        # Calculer la plus longue séquence
        max_consecutifs = 0
        current_consecutifs = 0

        current_date = date_debut
        while current_date <= date_fin:
            jour_type = jours_type.get(current_date, 'AUTRE')

            if jour_type == 'CONGES':
                current_consecutifs += 1
                max_consecutifs = max(max_consecutifs, current_consecutifs)
            elif jour_type == 'WEEKEND':
                pass  # Ne casse pas la séquence
            else:
                current_consecutifs = 0

            current_date += pd.Timedelta(days=1)

        stats[collaborateur]['jours_consecutifs_max'] = max_consecutifs
        stats[collaborateur]['jours_total_periode'] = total_jours
        stats[collaborateur]['regle_10j_consecutifs'] = max_consecutifs >= 10
        stats[collaborateur]['regle_20j_total'] = total_jours >= 20

    return stats


def create_summary_sheet(wb, stats):
    """Crée la feuille de synthèse"""

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
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = Font(bold=True, size=14)
                if value == "✓":
                    cell.fill = green_fill
                else:
                    cell.fill = red_fill
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')

        row += 1

    # Totaux
    ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=row, column=1).fill = subheader_fill
    ws.cell(row=row, column=1).border = border

    for col in range(2, 8):
        col_letter = get_column_letter(col)
        formula = f"=SUM({col_letter}2:{col_letter}{row - 1})"
        cell = ws.cell(row=row, column=col, value=formula)
        cell.font = Font(bold=True)
        cell.fill = subheader_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = '0.0'

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

    ws.column_dimensions['A'].width = 25
    for col in range(2, 8):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions['H'].width = 12
    ws.column_dimensions['I'].width = 12

    ws.merge_cells(f'A{row + 2}:I{row + 2}')
    note_cell = ws[f'A{row + 2}']
    note_cell.value = "📋 Règles RH (période 15/05 - 15/10) : 10j consécutifs = au moins 10 jours d'affilée | 20j total = au moins 20 jours (consécutifs ou non)"
    note_cell.font = Font(size=9, italic=True)
    note_cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws.row_dimensions[row + 2].height = 30

    ws.freeze_panes = 'A2'


def create_calendar_sheets(wb, csv_file):
    """Crée une feuille de calendrier par collaborateur"""

    df = pd.read_csv(csv_file)
    df['date_obj'] = pd.to_datetime(df['date'], format='%Y/%m/%d')

    collaborateurs = sorted(df['collaborateur'].unique())

    # Styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    month_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    month_font = Font(bold=True, size=10)

    # Couleurs
    teletravail_fill = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
    teletravail_pending_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    conges_fill = PatternFill(start_color="A9D08E", end_color="A9D08E", fill_type="solid")
    conges_pending_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    rtt_fill = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
    rtt_pending_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    weekend_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    mixed_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    inexistant_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    border_diagonal = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
        diagonal=Side(style='thin', color='FF0000'),
        diagonalDown=True
    )

    mois_noms = [
        "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
        "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
    ]

    jours_par_mois = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]

    for collaborateur in collaborateurs:
        ws = wb.create_sheet(collaborateur[:31])  # Limite de 31 caractères pour nom de feuille

        print(f"  📅 Création du planning pour {collaborateur}...")

        # Titre
        ws.merge_cells('A1:AF1')
        title_cell = ws['A1']
        title_cell.value = f"PLANNING 2026 - {collaborateur}"
        title_cell.font = Font(bold=True, size=14, color="FFFFFF")
        title_cell.fill = header_fill
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 25

        # Légende améliorée
        ws['A2'] = "LÉGENDE"
        ws['A2'].font = Font(bold=True, size=12)
        ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

        # Ligne 2 : Télétravail et Congés
        legend_row2 = [
            ("B2:C2", "TV", "Télétravail validé", teletravail_fill, "FFFFFF"),
            ("D2:E2", "TP", "Télétravail à valider", teletravail_pending_fill, "000000"),
            ("F2:G2", "CV", "Congés validés", conges_fill, "FFFFFF"),
            ("H2:I2", "CP", "Congés à valider", conges_pending_fill, "000000"),
        ]

        # Ligne 3 : RTT et autres
        ws['A3'] = ""
        legend_row3 = [
            ("B3:C3", "RV", "RTT validés", rtt_fill, "FFFFFF"),
            ("D3:E3", "RP", "RTT à valider", rtt_pending_fill, "000000"),
            ("F3:G3", "W", "Week-end/Férié", weekend_fill, "000000"),
            ("H3:I3", "CV/TV", "Journée mixte", mixed_fill, "000000"),
        ]

        # Appliquer les styles pour la ligne 2
        for cell_range, code, description, fill, font_color in legend_row2:
            start_cell = cell_range.split(':')[0]
            cell = ws[start_cell]
            cell.value = f"{code}\n{description}"
            cell.fill = fill
            cell.font = Font(bold=True, size=9, color=font_color)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
            ws.merge_cells(cell_range)

        # Appliquer les styles pour la ligne 3
        for cell_range, code, description, fill, font_color in legend_row3:
            start_cell = cell_range.split(':')[0]
            cell = ws[start_cell]
            cell.value = f"{code}\n{description}"
            cell.fill = fill
            cell.font = Font(bold=True, size=9, color=font_color)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
            ws.merge_cells(cell_range)

        # Ajuster la hauteur des lignes de légende
        ws.row_dimensions[2].height = 30
        ws.row_dimensions[3].height = 30

        # Ligne 4 : Explications des formats spéciaux
        ws.merge_cells('A4:I4')
        explanation_cell = ws['A4']
        explanation_cell.value = "💡 Formats spéciaux : Code-AM/PM = demi-journée | Code1/Code2 = matin≠après-midi | Cellule barrée = jour inexistant"
        explanation_cell.font = Font(size=9, italic=True)
        explanation_cell.alignment = Alignment(horizontal='left', vertical='center')
        explanation_cell.fill = PatternFill(start_color="FFF9E6", end_color="FFF9E6", fill_type="solid")
        ws.row_dimensions[4].height = 20

        # En-têtes (décalé à ligne 6 maintenant)
        ws['A6'] = "Mois"
        ws['A6'].fill = header_fill
        ws['A6'].font = header_font
        ws['A6'].alignment = Alignment(horizontal='center', vertical='center')
        ws['A6'].border = border

        for day in range(1, 32):
            col = day + 1
            cell = ws.cell(row=6, column=col, value=day)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = border

        ws.row_dimensions[6].height = 20

        # Largeurs
        ws.column_dimensions['A'].width = 12
        for col in range(2, 33):
            ws.column_dimensions[get_column_letter(col)].width = 5

        # Données du collaborateur
        collab_df = df[df['collaborateur'] == collaborateur].copy()
        collab_data = {}

        for _, row in collab_df.iterrows():
            month = row['date_obj'].month
            day = row['date_obj'].day

            if month not in collab_data:
                collab_data[month] = {}

            code = get_status_code(row['type_am'], row['type_pm'], row['detail_am'], row['detail_pm'])
            collab_data[month][day] = code

        # Remplir le calendrier (commence à ligne 7 maintenant)
        for month_num, month_name in enumerate(mois_noms, 1):
            row_num = 7 + month_num - 1  # Commence à la ligne 7 au lieu de 4
            max_days_in_month = jours_par_mois[month_num - 1]

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

                if day <= max_days_in_month:
                    # Jour valide
                    cell.border = border
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(size=6, bold=True)

                    if month_num in collab_data and day in collab_data[month_num]:
                        code = collab_data[month_num][day]
                        cell.value = code

                        # Appliquer la couleur
                        if '/' in code:
                            cell.fill = mixed_fill
                        elif code.startswith('W'):
                            cell.fill = weekend_fill
                        elif code.startswith('TV'):
                            cell.fill = teletravail_fill
                        elif code.startswith('TP'):
                            cell.fill = teletravail_pending_fill
                        elif code.startswith('CV'):
                            cell.fill = conges_fill
                        elif code.startswith('CP'):
                            cell.fill = conges_pending_fill
                        elif code.startswith('RV'):
                            cell.fill = rtt_fill
                        elif code.startswith('RP'):
                            cell.fill = rtt_pending_fill
                else:
                    # Jour inexistant
                    cell.border = border_diagonal
                    cell.fill = inexistant_fill

        ws.freeze_panes = 'B7'  # Geler jusqu'à la colonne A et ligne 6


def create_excel_report(stats, csv_file, output_file):
    """Crée le fichier Excel complet"""

    wb = openpyxl.Workbook()

    print("📊 Création de la feuille Synthèse...")
    create_summary_sheet(wb, stats)

    print("📅 Création des feuilles Planning par collaborateur...")
    create_calendar_sheets(wb, csv_file)

    wb.save(output_file)
    print(f"✅ Fichier Excel créé : {output_file}")


def main():
    """Fonction principale"""

    input_file = "/mnt/user-data/uploads/1770580647304_leave_planning_2026_complete.csv"
    output_file = "/mnt/user-data/outputs/compteurs_conges_2026_compatible.xlsx"

    print("🔍 Analyse du fichier CSV...")
    stats = analyze_leave_data(input_file)

    print(f"📊 Nombre de collaborateurs : {len(stats)}")

    print("📝 Génération du fichier Excel...")
    create_excel_report(stats, input_file, output_file)

    print("\n✨ Terminé !")
    print("\nContenu du fichier :")
    print("  - Feuille 'Synthèse' : compteurs + règles RH")
    print("  - Une feuille par collaborateur avec son calendrier annuel")
    print("\n✅ Version ULTRA-COMPATIBLE - Fonctionne sur TOUS les Excel !")
    print("   (Pas de formules complexes, pas de liste déroulante)")


if __name__ == "__main__":
    main()