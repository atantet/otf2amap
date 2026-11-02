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

from .extract import CALIBRES
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
            if len(qtys) > len(paniers):
                raise ValueError(
                    f"{r['prod']} : {len(qtys)} quantités commandées pour "
                    f"{len(paniers)} panier(s) reconnu(s) — extraction des "
                    "paniers incomplète ?")
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


# ── Fusion des formats d'un même produit à la pièce ──────────────────────────

# Unités « à la pièce » concernées par la fusion des formats.
UNITES_PIECE = frozenset(('u', 'u.', 'bte'))

# Libellé du format standard (nom sans qualificatif de calibre).
FORMAT_STANDARD = 'moy.'


def _nom_et_format(prod):
    """
    Sépare un nom de produit en (clé sans calibre, format).
      "Cresson en petite botte" → ("cresson en botte", "petite")
      "Cresson en botte"        → ("cresson en botte", None)
    """
    mots = re.findall(r"[\w'’]+", prod.lower())
    calibres = [m for m in mots if m in CALIBRES]
    cle = ' '.join(m for m in mots if m not in CALIBRES)
    return cle, (calibres[0] if calibres else None)


def _sans_calibre(prod):
    """Retire du nom affiché le premier qualificatif de calibre."""
    for m in re.finditer(r"[\w'’]+", prod):
        if m.group().lower() in CALIBRES:
            nom = (prod[:m.start()] + prod[m.end():]).strip()
            return re.sub(r'\s+', ' ', nom[:1].upper() + nom[1:])
    return prod


def _qte_format(q, libelle):
    """ "1 petite", "2 petites", "3 moy." (les abréviations restent invariables)."""
    if q > 1 and not libelle.endswith('.') and not libelle.endswith('s'):
        libelle += 's'
    return f"{fmt(q)} {libelle}"


def fusionner_formats(rows, paniers):
    """
    Regroupe sur une seule ligne les formats d'un même produit à la pièce
    (« X en botte » / « X en petite botte », « X à l'unité » / « X petite à
    l'unité »…) : chaque case porte alors la quantité ET le format, ex.
    « 1 petite », « 1 moy. », « 1 moy. + 1 petite ».

    Deux lignes sont fusionnées si leurs noms sont identiques au qualificatif
    de calibre près, avec la même unité à la pièce (u., bte) et des formats
    différents. Les produits au poids ne sont jamais fusionnés.
    Retourne la nouvelle liste de lignes (ordre conservé).
    """
    groupes = {}
    for r in rows:
        if r.get('unite') not in UNITES_PIECE:
            continue
        cle, _ = _nom_et_format(r['prod'])
        groupes.setdefault((cle, r['unite'].rstrip('.')), []).append(r)

    fusion = {}   # id(ligne de tête) → ligne fusionnée ; autres ids → None
    for membres in groupes.values():
        formats = [_nom_et_format(r['prod'])[1] for r in membres]
        if len(membres) < 2 or len(set(formats)) < len(formats):
            continue
        # Format standard d'abord, puis les autres dans l'ordre d'arrivée.
        ordre = sorted(zip(formats, membres, strict=True),
                       key=lambda fm: fm[0] is not None)
        libelles = [(f or FORMAT_STANDARD, r) for f, r in ordre]

        cells = {}
        for pan in paniers:
            parts = []
            for lib, r in libelles:
                val = r['cells'].get(pan['key'], '')
                if val:
                    parts.append(_qte_format(float(val.split()[0]), lib))
            cells[pan['key']] = ' + '.join(parts)

        total = ' + '.join(_qte_format(float(r['qty_num']), lib)
                           for lib, r in libelles)
        tete = ordre[0][1]
        fusion[id(membres[0])] = {
            **tete,
            'prod': tete['prod'] if ordre[0][0] is None else _sans_calibre(tete['prod']),
            'qty_num': total,
            'unite': '',
            'cells': cells,
        }
        for r in membres[1:]:
            fusion[id(r)] = None

    out = []
    for r in rows:
        if id(r) in fusion:
            if fusion[id(r)] is not None:
                out.append(fusion[id(r)])
        else:
            out.append(r)
    return out
