#!/usr/bin/env python3
"""
Script pour transformer la première page d'un PDF de ventes de légumes.

- Titre = date de retrait extraite de la page 2
- En-tête en gras, sans couleur de fond (noir et blanc)
- Colonnes : PRODUIT | QUANTITÉ | MONTANT | N PETIT | N MOYEN | N GRAND
- Lignes "Panier de la semaine" supprimées
- Page unique au format A5

Usage : python3 otf2amap.py entree.pdf [sortie.pdf]
"""

import sys
import re
import io
from pathlib import Path
from collections import defaultdict

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5
from reportlab.lib.colors import Color, black, white

# ── Couleurs ──────────────────────────────────────────────────────────────────
BLACK     = Color(0, 0, 0)
WHITE     = Color(1, 1, 1)
ROW_LIGHT = Color(0.94, 0.94, 0.94)   # gris très clair pour lignes paires
ROW_DARK  = Color(0.87, 0.87, 0.87)   # gris un peu plus foncé pour lignes impaires
SEP       = Color(0.65, 0.65, 0.65)

# ── Format A5 (points) ────────────────────────────────────────────────────────
PAGE_W, PAGE_H = A5          # 419.53 x 595.28
MARGIN     = 10.0
PAGE_RIGHT = PAGE_W - MARGIN

FONT      = "Helvetica"
FONT_BOLD = "Helvetica-Bold"


def clean(s):
    if not s: return ''
    return re.sub(r'\s+', ' ', s.replace('\xa0', ' ').replace('\n', ' ')).strip()


def fmt(val):
    if val == int(val): return str(int(val))
    return f"{val:.2f}".rstrip('0').rstrip('.')


def extract_date_from_page2(pdf_path):
    """Extrait la date de retrait depuis la page 2."""
    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            return None
        text = pdf.pages[1].extract_text() or ''
    # Cherche un motif DD/MM/YYYY
    m = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', text)
    return m.group(1) if m else None


def parse_raw_cmd(raw):
    """
    Dans le cas où qty et montant sont fusionnés dans la colonne cmd,
    extrait (qty, montant, commandes) depuis une chaîne du type :
    '5.84 kg 34,47 € 1 x 2.1 kg 1 x 3.74 kg'
    ou '3 bte 8,70 € 1 x 3 bte'
    ou '3 u. 45,90 € 1 x 3 u.'
    """
    # Motif : <nombre> <unité> <prix> € <reste commandes>
    m = re.match(
        r'^([\d.]+)\s+(\S+)\s+([\d,]+\s*€)(.*)',
        raw.strip()
    )
    if m:
        qty  = m.group(1) + ' ' + m.group(2)
        mon  = m.group(3).strip()
        cmd  = m.group(4).strip()
        return qty, mon, cmd
    return '', '', raw


def extract_table_data(pdf_path):
    """
    Extrait les données de la page 1.
    Retourne (rows, paniers).
    Gère deux mises en page :
      - Layout A : colonnes QUANTITÉ et MONTANT distinctes (x 198-297)
      - Layout B : tout dans la colonne après PRODUIT (x > ~198), sans colonnes séparées
    """
    X_PROD_END  = 198.0
    X_QTY_START = 198.0
    X_QTY_END   = 252.0
    X_MON_START = 252.0
    X_MON_END   = 297.0
    X_CMD_START = 297.0

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words(x_tolerance=3, y_tolerance=3)

    by_y = defaultdict(list)
    for w in words:
        by_y[round(w['top'] / 4) * 4].append(w)

    def col(ws, x0, x1): return [w for w in ws if x0 < w['x0'] < x1]
    def txt(ws): return clean(' '.join(w['text'] for w in ws))

    # Détecter le layout : si aucune ligne de données n'a de mots dans
    # la zone QUANTITÉ (198-252), on est en layout B (tout dans cmd).
    data_ys = [y for y in sorted(by_y.keys()) if y > 70]
    has_qty_col = any(
        col(by_y[y], X_QTY_START, X_QTY_END)
        for y in data_ys[:5]
    )

    segs = []
    for y in sorted(by_y.keys()):
        ws = by_y[y]
        prod_ws = [w for w in ws if w['x0'] < X_PROD_END]
        if has_qty_col:
            segs.append({
                'y':    y,
                'prod': txt(prod_ws),
                'qty':  txt(col(ws, X_QTY_START, X_QTY_END)),
                'mon':  txt(col(ws, X_MON_START, X_MON_END)),
                'cmd':  txt(col(ws, X_CMD_START, 999)),
            })
        else:
            # Layout B : tout ce qui est après PRODUIT va dans cmd brut
            # (qty/mon seront extraits apres fusion des lignes fragmentees)
            raw_cmd = txt([w for w in ws if w['x0'] >= X_PROD_END])
            segs.append({
                'y':    y,
                'prod': txt(prod_ws),
                'qty':  '',
                'mon':  '',
                'cmd':  raw_cmd,
            })

    def is_header(s):
        return any(kw in s for kw in ('PRODUIT', 'QUANTITÉ', 'MONTANT', 'COMMANDES'))

    def is_titre(s):
        return bool(re.match(r'^\d+\s+vente', s))

    def is_name_only(seg):
        """Segment qui n'est qu'une suite de nom produit (pas de données)."""
        return (seg['prod']
                and not seg['qty'] and not seg['mon'] and not seg['cmd']
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

        # Grand chiffre flottant sur la ligne juste AVANT le nom
        # Layout A : chiffre dans qty ; Layout B : chiffre dans cmd
        if i > 0:
            prev = segs[i - 1]
            if not prev['prod'] and not prev['mon']:
                if prev['qty'] and not prev['cmd']:
                    row['qty'] = (prev['qty'] + ' ' + row['qty']).strip() if row['qty'] else prev['qty']
                elif prev['cmd'] and not prev['qty']:
                    row['cmd'] = (prev['cmd'] + ' ' + row['cmd']).strip() if row['cmd'] else prev['cmd']

        # Parcourir les lignes suivantes pour compléter les champs manquants
        j = i + 1
        while j < len(segs):
            nxt = segs[j]

            # Nouveau produit réel → stop (sauf si c'est une suite de nom)
            if nxt['prod'] and not is_header(nxt['prod']) and not is_titre(nxt['prod']):
                if is_name_only(nxt):
                    row['prod'] += ' ' + nxt['prod']
                    j += 1
                    continue
                break

            # Compléter qty
            if nxt['qty']:
                row['qty'] = (row['qty'] + ' ' + nxt['qty']).strip() if row['qty'] else nxt['qty']
            # Compléter mon
            if nxt['mon'] and not row['mon']:
                row['mon'] = nxt['mon']
            # Compléter cmd
            if nxt['cmd']:
                row['cmd'] = (row['cmd'] + ' ' + nxt['cmd']).strip() if row['cmd'] else nxt['cmd']

            j += 1

            # Dès qu'on a qty et montant, vérifier encore une ligne pour suite de nom
            if row['qty'] and row['mon']:
                if j < len(segs) and is_name_only(segs[j]):
                    row['prod'] += ' ' + segs[j]['prod']
                    j += 1
                break

        i = j if j > i + 1 else i + 1

        # En layout B, qty et mon sont vides : les extraire depuis cmd fusionne
        if not row['qty'] and row['cmd']:
            row['qty'], row['mon'], row['cmd'] = parse_raw_cmd(row['cmd'])

        # Dédoublonner les unités dans qty (ex: "3 u. u." → "3 u.")
        row['qty'] = re.sub(r'\b(\w+\.?)\s+\1\b', r'\1', row['qty'])

        if row['qty'] and re.search(r'\d', row['qty']):
            rows_raw.append(row)

    # Séparer paniers / produits
    PANIER_KEYS = [('petit', 'Petit'), ('moyen', 'Moyen'), ('grand', 'Grand')]
    ORDER = {'petit': 0, 'moyen': 1, 'grand': 2}
    paniers, rows = [], []

    for r in rows_raw:
        low = r['prod'].lower()
        matched = None
        for key, lbl in PANIER_KEYS:
            if 'panier de la semaine' in low and key in low:
                matched = (key, lbl)
                break
        if matched:
            key, lbl = matched
            nums = re.findall(r'\d+(?:\.\d+)?', r['qty'])
            n = int(float(nums[0])) if nums else 1
            paniers.append({'key': key, 'label': lbl, 'n': n})
        else:
            rows.append(r)

    paniers.sort(key=lambda p: ORDER.get(p['key'], 99))

    # Calculer les cellules par panier
    for r in rows:
        parts = r['qty'].split()
        qty_total = float(parts[0]) if parts else 0
        unite = parts[1] if len(parts) > 1 else ''
        r['qty_num'] = fmt(qty_total)
        r['unite']   = unite

        tokens = re.findall(r'1\s*x\s*([\d.]+)\s*(\S+)', r['cmd'])
        cells  = {p['key']: '' for p in paniers}

        # Attribuer chaque token au bon panier.
        #
        # Règle : qté/panier est croissant du petit au grand panier.
        # 'n' = nombre de paniers de ce type dans la commande :
        #   petit panier → beaucoup d'exemplaires → grand n
        #   grand panier → peu d'exemplaires → petit n
        # Donc qté_totale_token / n donne la qté par panier,
        # et la contrainte qté/petit <= qté/moyen <= qté/grand
        # correspond à : token trié croissant ↔ panier trié par n croissant.
        #
        # Cas à 1 token : on choisit le panier dont n divise exactement
        # la quantité totale (ratio entier), ce qui identifie le panier sans ambiguïté.

        import itertools

        qtys  = [float(q) for q, u in tokens]
        units = [u for q, u in tokens]

        if len(qtys) == 1:
            # Chercher le panier dont n divise exactement la quantité
            q, u = float(tokens[0][0]), tokens[0][1]
            best = min(paniers, key=lambda p: abs(q / p['n'] - round(q / p['n'])))
            cells[best['key']] = f"{fmt(q / best['n'])} {u}"
        else:
            # Trier les tokens croissant et les paniers par n croissant,
            # puis les associer dans cet ordre.
            token_pairs_sorted = sorted(zip(qtys, units), key=lambda x: x[0])
            pans_sorted_by_n   = sorted(paniers, key=lambda p: p['n'])
            # On ne sélectionne que len(qtys) paniers (les plus petits n en premier)
            for (q, u), pan in zip(token_pairs_sorted, pans_sorted_by_n):
                cells[pan['key']] = f"{fmt(q / pan['n'])} {u}"

        r['cells'] = cells

    return rows, paniers


def build_new_page(rows, paniers, titre, avec_montant=False):
    """Construit la nouvelle page au format A5, noir et blanc."""
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))

    # Largeurs de colonnes (proportionnées pour A5)
    W_PROD = 108.0
    W_QTY  = 38.0
    W_MON  = 34.0 if avec_montant else 0.0
    n_p    = len(paniers)
    W_PAN  = (PAGE_RIGHT - MARGIN - W_PROD - W_QTY - W_MON) / max(n_p, 1)

    xP    = MARGIN
    xQ    = xP + W_PROD
    xM    = xQ + W_QTY
    xPans = [xM + W_MON + i * W_PAN for i in range(n_p)]

    def cx(x0, w): return x0 + w / 2

    # Fond blanc
    c.setFillColor(WHITE)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    # ── Titre (date de retrait) ───────────────────────────────────────────────
    c.setFillColor(BLACK)
    c.setFont(FONT_BOLD, 16)
    c.drawCentredString(PAGE_W / 2, PAGE_H - 28, titre)

    # ── En-tête tableau : fond blanc, texte noir gras, bordure noire ──────────
    HDR_Y = PAGE_H - 52.0
    HDR_H = 16.0

    # Bordure de l'en-tête
    c.setStrokeColor(BLACK)
    c.setLineWidth(0.6)
    c.rect(MARGIN, HDR_Y, PAGE_RIGHT - MARGIN, HDR_H, fill=0, stroke=1)

    hy = HDR_Y + 5
    c.setFillColor(BLACK)
    c.setFont(FONT_BOLD, 6.5)
    c.drawString(xP + 3, hy, "PRODUIT")
    c.drawCentredString(cx(xQ, W_QTY), hy, "QUANTITÉ")
    if avec_montant:
        c.drawCentredString(cx(xM, W_MON), hy, "MONTANT")
    for i, pan in enumerate(paniers):
        c.drawCentredString(cx(xPans[i], W_PAN), hy, f"{pan['n']} {pan['label'].upper()}")

    # Séparateurs verticaux en-tête
    c.setStrokeColor(BLACK)
    c.setLineWidth(0.5)
    sep_xs = [xQ] + ([xM] if avec_montant else []) + xPans
    for xs in sep_xs:
        c.line(xs, HDR_Y, xs, HDR_Y + HDR_H)

    # ── Lignes de données ─────────────────────────────────────────────────────
    ROW_H = 16.5
    S_ROW = 6.5
    S_QTY = 9.0
    cur_y = HDR_Y

    def draw_prod(text, ymid, max_w):
        c.setFont(FONT, S_ROW)
        if c.stringWidth(text, FONT, S_ROW) <= max_w:
            c.drawString(xP + 3, ymid - S_ROW / 2 + 1, text)
            return
        # Coupure au ' / '
        if ' / ' in text:
            p1, p2 = text.split(' / ', 1)
            c.drawString(xP + 3, ymid + 1,        p1 + ' /')
            c.drawString(xP + 3, ymid - S_ROW - 1, p2)
        else:
            wds = text.split()
            l1 = ''
            for w in wds:
                t = (l1 + ' ' + w).strip()
                if c.stringWidth(t, FONT, S_ROW) <= max_w:
                    l1 = t
                else:
                    break
            l2 = text[len(l1):].strip()
            c.drawString(xP + 3, ymid + 1,        l1)
            c.drawString(xP + 3, ymid - S_ROW - 1, l2)

    for idx, row in enumerate(rows):
        bg = ROW_LIGHT if idx % 2 == 0 else ROW_DARK
        rb = cur_y - ROW_H
        ym = rb + ROW_H / 2

        c.setFillColor(bg)
        c.rect(MARGIN, rb, PAGE_RIGHT - MARGIN, ROW_H, fill=1, stroke=0)

        # Bordure basse de la ligne
        c.setStrokeColor(SEP)
        c.setLineWidth(0.3)
        c.line(MARGIN, rb, PAGE_RIGHT, rb)

        c.setFillColor(BLACK)

        draw_prod(row['prod'], ym, W_PROD - 6)

        # Quantité
        c.setFont(FONT_BOLD, S_QTY)
        qw = c.stringWidth(row['qty_num'], FONT_BOLD, S_QTY)
        uw = c.stringWidth(row['unite'],   FONT,      S_ROW)
        qx = cx(xQ, W_QTY) - (qw + 2 + uw) / 2
        c.drawString(qx, ym - S_QTY / 2 + 1, row['qty_num'])
        c.setFont(FONT, S_ROW)
        c.drawString(qx + qw + 2, ym - S_ROW / 2 + 1, row['unite'])

        # Montant
        if avec_montant:
            c.setFont(FONT, S_ROW)
            c.drawCentredString(cx(xM, W_MON), ym - S_ROW / 2 + 1, row['mon'])

        # Cellules paniers
        for i, pan in enumerate(paniers):
            val = row['cells'].get(pan['key'], '')
            if val:
                c.setFont(FONT, S_ROW)
                c.drawCentredString(cx(xPans[i], W_PAN), ym - S_ROW / 2 + 1, val)

        # Séparateurs verticaux
        c.setStrokeColor(SEP)
        c.setLineWidth(0.3)
        for xs in sep_xs:
            c.line(xs, rb, xs, cur_y)

        cur_y = rb

    # Bordure extérieure du tableau (bas + côtés)
    table_top    = HDR_Y + HDR_H
    table_bottom = cur_y
    c.setStrokeColor(BLACK)
    c.setLineWidth(0.6)
    c.rect(MARGIN, table_bottom, PAGE_RIGHT - MARGIN, table_top - table_bottom,
           fill=0, stroke=1)

    c.save()
    packet.seek(0)
    return packet


def transformer_pdf(input_path, output_path=None, avec_montant=False):
    input_path  = Path(input_path)
    output_path = Path(output_path) if output_path else \
                  input_path.parent / (input_path.stem + "_modifie.pdf")

    print(f"Lecture de : {input_path}")

    date = extract_date_from_page2(input_path)
    titre = f"Retrait du {date}" if date else "Ventes"
    print(f"  Titre : {titre}")

    rows, paniers = extract_table_data(input_path)

    if not paniers:
        print("ERREUR : aucun panier 'Panier de la semaine' trouvé.")
        sys.exit(1)

    for p in paniers:
        print(f"  Panier {p['label']} : {p['n']} unité(s)")
    print(f"  Produits : {len(rows)}")
    for r in rows:
        print(f"    {r['prod']} | {r['qty_num']} {r['unite']} | {r['mon']} | {r['cells']}")

    new_page = build_new_page(rows, paniers, titre, avec_montant=avec_montant)

    # Page unique
    writer = PdfWriter()
    writer.add_page(PdfReader(new_page).pages[0])

    with open(output_path, "wb") as f:
        writer.write(f)
    print(f"\nPDF enregistré : {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)
    args = sys.argv[1:]
    avec_montant = "--montant" in args
    args = [a for a in args if a != "--montant"]
    transformer_pdf(args[0], args[1] if len(args) > 1 else None, avec_montant=avec_montant)
