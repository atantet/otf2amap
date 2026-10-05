"""Lecture d'un tableur de distribution Pommes / Poires (.xls AmapJ) → poids total.

Sur la première feuille du tableur AmapJ :
  - D9  : unité de vente (ex. « 500 g », « 1 kg ») ;
  - D11 : ligne « Cumul », nombre d'unités commandées pour la livraison.

`poids_total` renvoie le nombre d'unités, l'unité et le poids total en kg.
"""

import re

import xlrd

_CELLULE_UNITE = (8, 3)    # D9
_CELLULE_CUMUL = (10, 3)   # D11

_UNITE_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(kg|g)\b", re.IGNORECASE)


def _unite_en_kg(texte):
    """« 500 g » → 0.5, « 1 kg » → 1.0 ; None si illisible."""
    m = _UNITE_RE.search(texte or "")
    if not m:
        return None
    valeur = float(m.group(1).replace(",", "."))
    return valeur if m.group(2).lower() == "kg" else valeur / 1000


def poids_total(xls_path):
    """Renvoie (quantité, unité_texte, poids_kg) ; poids_kg vaut None si l'unité est illisible."""
    sheet = xlrd.open_workbook(xls_path).sheet_by_index(0)
    cumul = sheet.cell_value(*_CELLULE_CUMUL)
    quantite = cumul if isinstance(cumul, (int, float)) else 0
    unite = str(sheet.cell_value(*_CELLULE_UNITE)).strip()
    kg = _unite_en_kg(unite)
    return quantite, unite, (quantite * kg if kg is not None else None)
