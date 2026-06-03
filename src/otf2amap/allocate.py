"""
Attribution des quantités aux paniers.

Pour chaque produit, on remplit une case par type de panier. Deux sources :
  - lecture directe depuis la page 2 (préférée, exacte) ;
  - sinon, repli sur les tokens "1 x N unité" de la colonne COMMANDES de la
    page 1, attribués par permutation optimale (score lexicographique sur
    l'erreur d'arrondi des ratios qté/n_panier à 2, 1 puis 0 décimales).
"""

import itertools
import re

from .util import fmt


def match_produit(nom_p1, page2_data):
    """
    Cherche le nom de produit de la page 1 dans les données de la page 2.
    Correspondance exacte d'abord, puis partielle (inclusion).
    """
    if nom_p1 in page2_data:
        return page2_data[nom_p1]
    # Correspondance partielle : le nom p2 est contenu dans le nom p1 ou vice versa
    for nom_p2, cells in page2_data.items():
        if nom_p2 in nom_p1 or nom_p1 in nom_p2:
            return cells
    return None


def allocate(rows, paniers, page2_data):
    """
    Complète chaque ligne de `rows` (modifiée sur place) avec :
      - qty_num : quantité totale formatée
      - unite   : unité (kg, bte, u.)
      - cells   : {clé_panier: "qté unité"} par panier
    """
    for r in rows:
        parts = r['qty'].split()
        qty_total = float(parts[0]) if parts else 0
        r['qty_num'] = fmt(qty_total)
        r['unite']   = parts[1] if len(parts) > 1 else ''

        tokens = re.findall(r'1\s*x\s*([\d.]+)\s*(\S+)', r['cmd'])
        cells  = {p['key']: '' for p in paniers}
        qtys   = [float(q) for q, u in tokens]
        units  = [u for q, u in tokens]

        # Lecture directe depuis la page 2 si disponible
        p2_cells = match_produit(r['prod'], page2_data) if page2_data else None
        if p2_cells:
            for pan in paniers:
                if pan['key'] in p2_cells:
                    qty_p2, unit_p2 = p2_cells[pan['key']]
                    cells[pan['key']] = f"{fmt(qty_p2 / pan['n'])} {unit_p2}"
        elif qtys:
            # Fallback : permutation optimale (score lexicographique)
            best_score  = None
            best_assign = list(range(len(qtys)))
            for perm in itertools.permutations(range(len(paniers)), len(qtys)):
                ratios = [qtys[i] / paniers[perm[i]]['n'] for i in range(len(qtys))]
                score  = tuple(round(sum(abs(r2 - round(r2, dp)) for r2 in ratios), 10)
                               for dp in [2, 1, 0])
                if best_score is None or score < best_score:
                    best_score  = score
                    best_assign = perm
            for i, (q, u) in enumerate(zip(qtys, units, strict=False)):
                pan = paniers[best_assign[i]]
                cells[pan['key']] = f"{fmt(q / pan['n'])} {u}"

        r['cells'] = cells

    return rows
