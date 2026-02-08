#!/usr/bin/env python3
"""
Script pour générer des statistiques de congés par personne à partir du CSV de planning
"""

import pandas as pd
from collections import defaultdict
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


def analyze_leave_data(csv_file):
    """Analyse le fichier CSV et retourne les statistiques par personne"""

    # Lire le CSV
    df = pd.read_csv(csv_file)

    # Structure pour stocker les compteurs par personne
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

    # Parcourir chaque ligne
    for _, row in df.iterrows():
        collaborateur = row['collaborateur']

        # Traiter AM (matin)
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

        # Traiter PM (après-midi)
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


def create_excel_report(stats, output_file):
    """Crée un fichier Excel formaté avec les statistiques"""

    # Créer un nouveau classeur
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Compteurs Congés 2026"

    # Définir les styles
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    subheader_fill = PatternFill(start_color="B4C7E7", end_color="B4C7E7", fill_type="solid")
    subheader_font = Font(bold=True, size=10)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # En-têtes
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

    # Écrire les en-têtes
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border

    # Ajuster la hauteur de la ligne d'en-tête
    ws.row_dimensions[1].height = 30

    # Trier les collaborateurs par ordre alphabétique
    sorted_collaborateurs = sorted(stats.keys())

    # Remplir les données
    row = 2
    for collaborateur in sorted_collaborateurs:
        s = stats[collaborateur]

        # Calculer les jours (2 demi-journées = 1 jour)
        teletravail_valide = (s['teletravail_valide_am'] + s['teletravail_valide_pm']) / 2
        teletravail_a_valider = (s['teletravail_a_valider_am'] + s['teletravail_a_valider_pm']) / 2
        conges_valides = (s['conges_valides_am'] + s['conges_valides_pm']) / 2
        conges_a_valider = (s['conges_a_valider_am'] + s['conges_a_valider_pm']) / 2
        rtt_valides = (s['rtt_valides_am'] + s['rtt_valides_pm']) / 2
        rtt_a_valider = (s['rtt_a_valider_am'] + s['rtt_a_valider_pm']) / 2

        # Totaux
        total_valide = teletravail_valide + conges_valides + rtt_valides
        total_a_valider = teletravail_a_valider + conges_a_valider + rtt_a_valider
        total_general = total_valide + total_a_valider

        # Écrire les données
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

            # Aligner les nombres à droite
            if col > 1:
                cell.alignment = Alignment(horizontal='right', vertical='center')
                # Formater les nombres avec 1 décimale
                if isinstance(value, (int, float)):
                    cell.number_format = '0.0'
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')

        row += 1

    # Ajouter une ligne de totaux
    ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=row, column=1).fill = subheader_fill

    for col in range(2, 11):
        # Somme de la colonne
        col_letter = get_column_letter(col)
        formula = f"=SUM({col_letter}2:{col_letter}{row - 1})"
        cell = ws.cell(row=row, column=col, value=formula)
        cell.font = Font(bold=True)
        cell.fill = subheader_fill
        cell.border = border
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = '0.0'

    # Ajuster la largeur des colonnes
    ws.column_dimensions['A'].width = 25  # Collaborateur
    for col in range(2, 11):
        ws.column_dimensions[get_column_letter(col)].width = 14

    # Figer la première ligne
    ws.freeze_panes = 'A2'

    # Sauvegarder le fichier
    wb.save(output_file)
    print(f"✅ Fichier Excel créé : {output_file}")
    print(f"📊 Nombre de collaborateurs : {len(sorted_collaborateurs)}")


def main():
    """Fonction principale"""

    input_file = "leave_planning_2026_complete.csv"
    output_file = "compteurs_conges_2026.xlsx"

    print("🔍 Analyse du fichier CSV...")
    stats = analyze_leave_data(input_file)

    print("📝 Génération du fichier Excel...")
    create_excel_report(stats, output_file)

    print("\n✨ Terminé !")


if __name__ == "__main__":
    main()