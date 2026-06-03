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

    by_y = defaultdict(list)
    for w in words:
        by_y[round(w['top'] / 3) * 3].append(w)

    panier_map = {p['label'].lower(): p['key'] for p in paniers}
    result = defaultdict(dict)
    current_key = None

    for y in sorted(by_y.keys()):
        ws = sorted(by_y[y], key=lambda w: w['x0'])
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

    with pdfplumber.open(pdf_path) as pdf:
        words = pdf.pages[0].extract_words(x_tolerance=3, y_tolerance=3)

    by_y = defaultdict(list)
    for w in words:
        by_y[round(w['top'] / 4) * 4].append(w)

    def col(ws, x0, x1):
        return [w for w in ws if x0 < w['x0'] < x1]

    def txt(ws):
        return clean(' '.join(w['text'] for w in ws))

    data_ys = [y for y in sorted(by_y.keys()) if y > 70]
    has_qty_col = any(col(by_y[y], X_QTY_START, X_QTY_END) for y in data_ys[:5])

    segs = []
    for y in sorted(by_y.keys()):
        ws = by_y[y]
        prod_ws = [w for w in ws if w['x0'] < X_PROD_END]
        if has_qty_col:
            segs.append({'y': y,
                         'prod': txt(prod_ws),
                         'qty':  txt(col(ws, X_QTY_START, X_QTY_END)),
                         'mon':  txt(col(ws, X_MON_START, X_MON_END)),
                         'cmd':  txt(col(ws, X_CMD_START, 999))})
        else:
            segs.append({'y': y,
                         'prod': txt(prod_ws),
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

        # Chiffre sur la ligne juste avant le nom (layout A ou B)
        if i > 0:
            prev = segs[i - 1]
            if not prev['prod'] and not prev['mon']:
                if prev['qty'] and not prev['cmd']:
                    row['qty'] = (prev['qty'] + ' ' + row['qty']).strip() if row['qty'] else prev['qty']
                elif prev['cmd'] and not prev['qty']:
                    row['cmd'] = (prev['cmd'] + ' ' + row['cmd']).strip() if row['cmd'] else prev['cmd']

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

        i = j if j > i + 1 else i + 1

        # Layout B : extraire qty et mon depuis cmd
        if not row['qty'] and row['cmd']:
            row['qty'], row['mon'], row['cmd'] = parse_raw_cmd(row['cmd'])

        # Dédoublonner les unités (ex: "3 u. u." → "3 u.")
        row['qty'] = re.sub(r'\b(\w+\.?)\s+\1\b', r'\1', row['qty'])

        if row['qty'] and re.search(r'\d', row['qty']):
            rows_raw.append(row)

    # Séparer paniers et produits
    PANIER_KEYS = [('petit', 'Petit'), ('moyen', 'Moyen'), ('grand', 'Grand')]
    ORDER = {'petit': 0, 'moyen': 1, 'grand': 2}
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

    paniers.sort(key=lambda p: ORDER.get(p['key'], 99))

    return rows, paniers
