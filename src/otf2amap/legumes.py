"""Lecture d'un tableur de distribution Légumes (.xls AmapJ) → comptage paniers.

Le tableur exporté par AmapJ a, sur sa première feuille :
  - une ligne d'en-tête avec les colonnes « petit(s) panier(s) »,
    « moyen(s) panier(s) », « grand(s) panier(s) » ;
  - une ligne « Cumul » donnant le total à livrer pour chaque type.

`compter_paniers` lit cette ligne Cumul et renvoie le nombre de paniers par
type. C'est ce nombre qui sert à créer la vente sur OuvreTaFerme en amont.
"""

import xlrd

TYPES = ("petit", "moyen", "grand")


def _to_int(value):
    """Convertit une cellule (float xls ou '') en entier ; 0 si vide/non numérique."""
    if isinstance(value, (int, float)):
        return int(round(value))
    return 0


def compter_paniers(xls_path):
    """Renvoie {'petit': n, 'moyen': n, 'grand': n} depuis la ligne Cumul.

    Repère dynamiquement la colonne de chaque type via l'en-tête, puis lit la
    ligne dont la première cellule vaut « Cumul ». Les types absents du tableur
    valent 0. Lève ValueError si l'en-tête ou la ligne Cumul est introuvable.
    """
    book = xlrd.open_workbook(xls_path)
    sheet = book.sheet_by_index(0)

    # Colonne de chaque type, repérée sur la ligne d'en-tête « … panier(s) ».
    col_par_type = {}
    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            cell = sheet.cell_value(r, c)
            if not isinstance(cell, str) or "panier" not in cell.lower():
                continue
            for t in TYPES:
                if t in cell.lower():
                    col_par_type[t] = c
        if col_par_type:
            break
    if not col_par_type:
        raise ValueError("en-tête « … panier(s) » introuvable dans le tableur")

    # Ligne « Cumul » : première cellule de colonne 0 contenant « cumul ».
    cumul_row = None
    for r in range(sheet.nrows):
        cell = sheet.cell_value(r, 0)
        if isinstance(cell, str) and "cumul" in cell.strip().lower():
            cumul_row = r
            break
    if cumul_row is None:
        raise ValueError("ligne « Cumul » introuvable dans le tableur")

    return {t: _to_int(sheet.cell_value(cumul_row, col_par_type[t])) if t in col_par_type else 0
            for t in TYPES}
