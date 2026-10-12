"""Lecture d'un tableur de distribution Pommes / Poires (.xls AmapJ) → poids total.

Sur la première feuille du tableur AmapJ :
  - D9  : unité de vente (ex. « 500 g », « 1 kg ») ;
  - D11 : ligne « Cumul », nombre d'unités commandées pour la livraison.

`poids_total` renvoie le nombre d'unités, l'unité et le poids total en kg.
`lignes_fruits` en tire la ligne « Pommes … Poires … » de la feuille paniers.
"""

import re
from pathlib import Path

import xlrd

from .util import fmt

_CELLULE_UNITE = (8, 3)    # D9
_CELLULE_CUMUL = (10, 3)   # D11

_UNITE_RE = re.compile(r"(\d+(?:[.,]\d+)?)\s*(kg|g)\b", re.IGNORECASE)


FRUITS = ("Pommes", "Poires")   # ordre d'affichage


def nom_fruit(texte):
    """« Pommes » / « Poires » si `texte` désigne ce contrat, sinon None."""
    t = (texte or "").lower()
    for fruit in FRUITS:
        if fruit[:-1].lower() in t:
            return fruit
    return None


def tableurs_fruits(dossier):
    """Couples (fruit, chemin) des tableurs .xls Pommes / Poires du dossier, Pommes d'abord."""
    trouves = [(nom_fruit(c.name), c) for c in sorted(Path(dossier).glob("*.xls"))]
    return sorted((t for t in trouves if t[0]), key=lambda t: FRUITS.index(t[0]))


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


def lignes_fruits(dossier):
    """Ligne des fruits à ajouter à la feuille paniers (liste de 0 ou 1 ligne).

    Lit les tableurs de distribution .xls du dossier de la semaine (téléchargés
    par mail2amap). Les fruits ne sont pas répartis par panier : le premier
    occupe les colonnes produit et TOTAL, le second (champ `suite` : nom,
    quantité, unité) s'écrit sur la même ligne, nom dans la première colonne
    de panier et quantité dans la suivante. Un tableur illisible ou sans commande est ignoré.
    """
    fruits = []
    for fruit, chemin in tableurs_fruits(dossier):
        try:
            quantite, unite, kg = poids_total(chemin)
        except Exception as e:                                # robustesse lecture xls
            print(f"  AVERTISSEMENT : lecture {fruit} ({chemin.name}) impossible : {e}")
            continue
        if not quantite:
            continue
        qty_num, unite = (fmt(kg), "kg") if kg is not None else (fmt(quantite), unite)
        fruits.append((fruit, qty_num, unite))
    if not fruits:
        return []
    prod, qty_num, unite = fruits[0]
    ligne = {"prod": prod, "qty_num": qty_num, "unite": unite, "mon": "", "cells": {}}
    if len(fruits) > 1:
        ligne["suite"] = fruits[1]
    return [ligne]
