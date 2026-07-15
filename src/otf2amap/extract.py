"""
Extraction du texte des PDF OuvreTaFerme.

Le PDF source comporte deux pages :
  - page 1 : tableau des ventes (PRODUIT / QUANTITÉ / MONTANT / COMMANDES)
  - page 2 : récapitulatif des commandes par panier, avec la date de retrait
"""

import re
from collections import defaultdict

import pdfplumber

from .util import clean

# ── Page 2 : date et quantités par panier ─────────────────────────────────────

def extract_date_from_page2(pdf_path):
    """Retourne la date de retrait (DD/MM/YYYY) depuis la page 2, ou None."""
    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            return None
        text = pdf.pages[1].extract_text() or ''
    m = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', text)
    return m.group(1) if m else None


# Clés de paniers reconnues et leur ordre d'affichage (partagé page 1 / page 2).
PANIER_KEYS = [('petit', 'Petit'), ('moyen', 'Moyen'), ('grand', 'Grand')]
PANIER_ORDER = {'petit': 0, 'moyen': 1, 'grand': 2}

# Qualificatifs de calibre à ignorer pour regrouper les variantes d'un légume.
CALIBRES = frozenset(('demi', 'petit', 'petite', 'gros', 'grosse',
                      'grand', 'grande', 'moyen', 'moyenne'))


def group_words_by_line(words, tol):
    """Regroupe des mots en lignes visuelles selon leur position verticale.

    Un simple arrondi sur grille fixe (round(top / N) * N) sépare à tort deux
    mots d'une même ligne quand leur `top` diffère de quelques dixièmes de
    point et encadre une frontière d'arrondi (ex. le chiffre de quantité d'un
    « … en botte » rendu ~0.3pt plus bas que le nom du produit) : on tolère
    donc un écart de `tol` par rapport au premier mot de la ligne, sans
    dépendre d'une grille absolue. Retourne une liste de listes de mots,
    triées de haut en bas, chaque ligne étant elle-même triée de gauche à
    droite (x0) pour préserver l'ordre de lecture.
    """
    lignes, courante, ref_top = [], [], None
    for w in sorted(words, key=lambda w: w['top']):
        if courante and w['top'] - ref_top > tol:
            lignes.append(courante)
            courante, ref_top = [], None
        courante.append(w)
        ref_top = w['top'] if ref_top is None else min(ref_top, w['top'])
    if courante:
        lignes.append(courante)
    return [sorted(ligne, key=lambda w: w['x0']) for ligne in lignes]


def cle_tri_produit(prod):
    """Clé de tri regroupant les calibres d'un même légume (noms inchangés).

    Le tri alphabétique brut d'OuvreTaFerme éloigne deux calibres d'un même
    légume (« Chou … » vs « Demi chou … », « Concombre … » vs « Concombre
    petit … »). On trie sur le nom de base (partie avant « / »), débarrassé des
    qualificatifs de calibre, ce qui rend les variantes consécutives. Au sein
    d'un groupe : d'abord le calibre standard (sans qualificatif), puis le nom
    complet pour un ordre stable.
    """
    base = prod.split('/', 1)[0].lower()
    tokens = re.findall(r"[\w'’]+", base)
    mots = [m for m in tokens if m not in CALIBRES]
    return (' '.join(mots), len(tokens) - len(mots), prod.lower())


def extract_paniers_from_page2(pdf_path):
    """
    Extrait la liste des paniers depuis la page 2, dont l'en-tête porte
    toujours le nombre de paniers, quel que soit le format d'export :
        'Panier de la semaine - moyen 7 u. 107,10 €'  → {key:'moyen', n:7}
    Sert de repli quand la page 1 (format détaillé) ne porte pas ce nombre.
    Retourne la liste {key, label, n} triée petit/moyen/grand.
    """
    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            return []
        text = pdf.pages[1].extract_text() or ''

    paniers = []
    for line in text.splitlines():
        m = re.search(r'Panier de la semaine\s*-\s*(\w+)\s+(\d+(?:\.\d+)?)',
                      line, re.IGNORECASE)
        if not m:
            continue
        key = m.group(1).lower()
        label = next((lbl for k, lbl in PANIER_KEYS if k == key), None)
        if label:
            paniers.append({'key': key, 'label': label, 'n': int(float(m.group(2)))})

    paniers.sort(key=lambda p: PANIER_ORDER.get(p['key'], 99))
    return paniers


def extract_page2_quantities(pdf_path, paniers):
    """
    Extrait depuis la page 2 les quantités par panier pour chaque produit.
    Utilise les positions x pour ignorer les parasites (numéros de commande).
    Retourne {nom_produit: {clé_panier: (qty, unit)}}
    """
    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            return {}
        words = pdf.pages[1].extract_words(x_tolerance=3, y_tolerance=3)

    # Les noms de produits commencent vers x=168, les quantités vers x=290
    X_PROD = 140.0
    X_QTY  = 290.0

    panier_map = {p['label'].lower(): p['key'] for p in paniers}
    result = defaultdict(dict)
    current_key = None

    for ws in group_words_by_line(words, tol=3):
        line_full = ' '.join(w['text'] for w in ws)

        m = re.search(r'Panier de la semaine\s*-\s*(\w+)', line_full, re.IGNORECASE)
        if m:
            current_key = panier_map.get(m.group(1).lower())
            continue

        if current_key is None:
            continue

        prod_words = [w for w in ws if X_PROD <= w['x0'] < X_QTY]
        qty_words  = [w for w in ws if w['x0'] >= X_QTY]

        if not prod_words or not qty_words:
            continue

        nom = ' '.join(w['text'] for w in prod_words).strip()
        if 'panier de la semaine' in nom.lower():
            continue

        qty_str = ' '.join(w['text'] for w in qty_words)
        m = re.match(r'([\d.]+)\s+(kg|bte|u\.?)', qty_str)
        if m:
            result[nom][current_key] = (float(m.group(1)), m.group(2).rstrip('.'))

    return dict(result)


# ── Page 1 : tableau des ventes ───────────────────────────────────────────────

def parse_raw_cmd(raw):
    """
    Layout B : qty et montant fusionnés dans la colonne cmd.
    Extrait (qty, montant, commandes) depuis une chaîne du type :
      '5.84 kg 34,47 € 1 x 2.1 kg 1 x 3.74 kg'
    """
    m = re.match(r'^([\d.]+)\s+(\S+)\s+([\d,]+\s*€)(.*)', raw.strip())
    if m:
        return m.group(1) + ' ' + m.group(2), m.group(3).strip(), m.group(4).strip()
    return '', '', raw


def parse_page1(pdf_path):
    """
    Extrait produits et paniers depuis la page 1 du PDF.
    Gère deux layouts :
      Layout A : colonnes QUANTITÉ et MONTANT distinctes (x 198–297)
      Layout B : tout fusionné après la colonne PRODUIT
    Retourne (rows, paniers) où `rows` contient les champs bruts
    (prod, qty, mon, cmd) et `paniers` la liste {key, label, n} triée.
    """
    X_PROD_END  = 198.0
    X_QTY_START = 198.0
    X_QTY_END   = 252.0
    X_MON_START = 252.0
    X_MON_END   = 297.0
    X_CMD_START = 297.0
    # Format détaillé (OuvreTaFerme récent) : sous chaque en-tête "Panier de la
    # semaine - X" (nom à x≈33), la composition du panier est listée en retrait
    # (x≈53). Les vrais produits restent à x≈19. On ignore donc toute ligne dont
    # le nom de produit commence au-delà de ce seuil : c'est du détail de panier.
    X_INDENT = 45.0

    with pdfplumber.open(pdf_path) as pdf:
        words = pdf.pages[0].extract_words(x_tolerance=3, y_tolerance=3)

    lignes = group_words_by_line(words, tol=4)

    def col(ws, x0, x1):
        return [w for w in ws if x0 < w['x0'] < x1]

    def txt(ws):
        return clean(' '.join(w['text'] for w in ws))

    data_lignes = [ws for ws in lignes if min(w['top'] for w in ws) > 70]
    has_qty_col = any(col(ws, X_QTY_START, X_QTY_END) for ws in data_lignes[:5])

    segs = []
    for ws in lignes:
        prod_ws = [w for w in ws if w['x0'] < X_PROD_END]
        if prod_ws and min(w['x0'] for w in prod_ws) > X_INDENT:
            continue  # ligne de composition d'un panier (format détaillé) : ignorée
        if has_qty_col:
            segs.append({'prod': txt(prod_ws),
                         'qty':  txt(col(ws, X_QTY_START, X_QTY_END)),
                         'mon':  txt(col(ws, X_MON_START, X_MON_END)),
                         'cmd':  txt(col(ws, X_CMD_START, 999))})
        else:
            segs.append({'prod': txt(prod_ws),
                         'qty':  '', 'mon': '',
                         'cmd':  txt([w for w in ws if w['x0'] >= X_PROD_END])})

    def is_header(s):
        return any(kw in s for kw in ('PRODUIT', 'QUANTITÉ', 'MONTANT', 'COMMANDES'))

    def is_titre(s):
        return bool(re.match(r'^\d+\s+vente', s))

    def is_name_only(seg):
        return (seg['prod'] and not seg['qty'] and not seg['mon'] and not seg['cmd']
                and len(seg['prod'].split()) <= 2
                and not re.search(r'\d', seg['prod'])
                and not is_header(seg['prod']))

    rows_raw = []
    consumed_up_to = -1  # dernier indice absorbé par un look-ahead précédent
    i = 0
    while i < len(segs):
        s = segs[i]
        if is_titre(s['prod']) or is_header(s['prod']):
            i += 1
            continue
        if s['mon'] and re.match(r'^\d+\s+vente', s['mon']) and not s['prod']:
            i += 1
            continue
        if not s['prod']:
            i += 1
            continue

        row = {'prod': s['prod'], 'qty': s['qty'], 'mon': s['mon'], 'cmd': s['cmd']}

        # Chiffre sur la ligne juste avant le nom (layout B : quantité scindée).
        # Si segs[i-1] a déjà été absorbé par le look-ahead du produit précédent,
        # on ne l'applique QUE s'il s'agit d'un nombre nu (ex. "12" avant "u. 21,55 €…"),
        # qui est un préfixe de quantité pour le produit courant.
        # Un segment de type "1 x 3.85 kg…" est une continuation de cmd du produit
        # précédent : dans ce cas on ne le réapplique pas.
        if i > 0:
            prev = segs[i - 1]
            is_consumed = i - 1 <= consumed_up_to
            if not prev['prod'] and not prev['mon']:
                if prev['qty'] and not prev['cmd']:
                    if not is_consumed:
                        row['qty'] = (prev['qty'] + ' ' + row['qty']).strip() if row['qty'] else prev['qty']
                elif prev['cmd'] and not prev['qty']:
                    if not is_consumed:
                        row['cmd'] = (prev['cmd'] + ' ' + row['cmd']).strip() if row['cmd'] else prev['cmd']
                    elif (re.match(r'^[\d.]+\s*$', prev['cmd'].strip())
                          and not re.match(r'^[\d.]+\s+\S', row['cmd'].strip())):
                        # Nombre nu consommé mais encore nécessaire comme préfixe de quantité
                        row['cmd'] = prev['cmd'].strip() + ' ' + row['cmd']

        # Compléter les champs depuis les lignes suivantes
        j = i + 1
        while j < len(segs):
            nxt = segs[j]
            if nxt['prod'] and not is_header(nxt['prod']) and not is_titre(nxt['prod']):
                if is_name_only(nxt):
                    row['prod'] += ' ' + nxt['prod']
                    j += 1
                    continue
                break
            if nxt['qty']:
                row['qty'] = (row['qty'] + ' ' + nxt['qty']).strip() if row['qty'] else nxt['qty']
            if nxt['mon'] and not row['mon']:
                row['mon'] = nxt['mon']
            if nxt['cmd']:
                row['cmd'] = (row['cmd'] + ' ' + nxt['cmd']).strip() if row['cmd'] else nxt['cmd']
            j += 1
            if row['qty'] and row['mon']:
                if j < len(segs) and is_name_only(segs[j]):
                    row['prod'] += ' ' + segs[j]['prod']
                    j += 1
                break

        if j > i + 1:
            consumed_up_to = j - 1
            i = j
        else:
            i = i + 1

        # Layout B : extraire qty et mon depuis cmd
        if not row['qty'] and row['cmd']:
            row['qty'], row['mon'], row['cmd'] = parse_raw_cmd(row['cmd'])

        # Dédoublonner les unités (ex: "3 u. u." → "3 u.")
        row['qty'] = re.sub(r'\b(\w+\.?)\s+\1\b', r'\1', row['qty'])

        if row['qty'] and re.search(r'\d', row['qty']):
            rows_raw.append(row)

    # Séparer paniers et produits
    paniers, rows = [], []

    for r in rows_raw:
        low = r['prod'].lower()
        matched = next(((k, lbl) for k, lbl in PANIER_KEYS
                        if 'panier de la semaine' in low and k in low), None)
        if matched:
            key, lbl = matched
            nums = re.findall(r'\d+(?:\.\d+)?', r['qty'])
            paniers.append({'key': key, 'label': lbl, 'n': int(float(nums[0])) if nums else 1})
        else:
            rows.append(r)

    paniers.sort(key=lambda p: PANIER_ORDER.get(p['key'], 99))
    rows.sort(key=lambda r: cle_tri_produit(r['prod']))

    return rows, paniers
