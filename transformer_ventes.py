#!/usr/bin/env python3
"""
Script pour transformer la première page d'un PDF de ventes de légumes.

Remplace la colonne COMMANDES par des colonnes par type de panier
(Petit, Moyen, Grand) avec la quantité par panier dans chaque cellule.
Les lignes "Panier de la semaine" sont supprimées car l'info figure en en-tête.
Supporte 1 à 3 types de paniers selon la semaine.

Usage : python3 transformer_ventes.py Ventes-6.pdf [sortie.pdf]
"""

import sys
import re
import io
from pathlib import Path
from collections import defaultdict

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color

# ── Couleurs ─────────────────────────────────────────────────────────────────
DARK      = Color(0.1294, 0.1451, 0.1608)
WHITE     = Color(1.0, 1.0, 1.0)
HEADER_BG = Color(0.2667, 0.2667, 0.4157)
ROW_LIGHT = Color(0.9882, 0.9882, 0.9961)
ROW_DARK  = Color(0.9608, 0.9608, 0.9686)
SEP       = Color(0.80, 0.80, 0.84)
SEP_HDR   = Color(0.45, 0.45, 0.60)

PAGE_W     = 594.96
PAGE_H     = 841.92
MARGIN     = 14.25
PAGE_RIGHT = 581.25
FONT       = "Helvetica"
FONT_BOLD  = "Helvetica-Bold"


def clean(s):
    if not s: return ''
    return re.sub(r'\s+', ' ', s.replace('\xa0', ' ').replace('\n', ' ')).strip()


def fmt(val):
    if val == int(val): return str(int(val))
    return f"{val:.2f}".rstrip('0').rstrip('.')


def extract_table_data(pdf_path):
    """
    Extrait les données de la page 1 en reconstruisant les lignes du tableau.
    Retourne (rows, paniers).
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

    # Construire la liste brute de segments par ligne y
    segs = []
    for y in sorted(by_y.keys()):
        ws = by_y[y]
        segs.append({
            'y':    y,
            'prod': txt(col(ws, 0,           X_PROD_END)),
            'qty':  txt(col(ws, X_QTY_START, X_QTY_END)),
            'mon':  txt(col(ws, X_MON_START, X_MON_END)),
            'cmd':  txt(col(ws, X_CMD_START, 999)),
        })

    # Fusionner en lignes produit cohérentes
    # Règles observées dans le PDF :
    #  - Le grand chiffre (quantité) peut être sur la ligne y-4 au-dessus
    #    du nom de produit (Blette : qty='6.71' y=104 avant prod='Blette' y=108)
    #  - L'unité peut être sur la même ligne que le produit (qty='bte' côté 'Navet')
    #    ou sur une ligne séparée ('3' y=148 puis 'bte' sur 'Navet' y=152)
    #  - Pour "Panier moyen" : '3' y=192, 'u.' sur la ligne y=196 avec le nom
    #  - La fin du nom "rouge" peut apparaître APRÈS la ligne de données (y=268)

    rows_raw = []

    def is_header(s):
        return any(kw in s for kw in ('PRODUIT', 'QUANTITÉ', 'MONTANT', 'COMMANDES'))

    def is_titre(s):
        return bool(re.match(r'^\d+\s+vente', s))

    i = 0
    while i < len(segs):
        s = segs[i]

        # Sauter titre et en-tête
        if is_titre(s['prod']) or is_header(s['prod']):
            i += 1
            continue
        # Lignes totalement vides en colonne produit ET qty (orphelins de grand chiffre)
        if not s['prod'] and not s['mon'] and not s['cmd']:
            # Ce segment est un chiffre flottant (grand chiffre de quantité)
            # rattaché à la prochaine ligne produit -> on le conserve comme préfixe
            i += 1
            # On le récupère ci-dessous dans la logique de fusion
            continue
        # Ligne "1 vente" déplacée dans MON
        if s['mon'] and re.match(r'^\d+\s+vente', s['mon']) and not s['prod']:
            i += 1
            continue

        if not s['prod']:
            i += 1
            continue

        # On a un nom de produit
        row = {'prod': s['prod'], 'qty': s['qty'], 'mon': s['mon'], 'cmd': s['cmd']}

        # Chercher le grand chiffre flottant juste AVANT cette ligne (y - 4)
        if i > 0:
            prev = segs[i-1]
            if (not prev['prod'] and prev['qty'] and not prev['mon'] and not prev['cmd']):
                # C'est le grand chiffre séparé
                if row['qty']:
                    # Fusionner : grand chiffre + unité
                    row['qty'] = prev['qty'] + ' ' + row['qty']
                else:
                    row['qty'] = prev['qty']

        # Si la qty est encore incomplète (contient seulement l'unité sans chiffre),
        # chercher la ligne suivante
        if row['qty'] and not re.search(r'\d', row['qty']):
            # qty ne contient que des lettres (ex: "bte", "kg") → chercher le chiffre avant
            # (déjà traité ci-dessus normalement)
            pass

        # Chercher les lignes suivantes pour compléter qty/mon/cmd et suite du nom
        j = i + 1
        while j < len(segs):
            nxt = segs[j]
            # Si c'est un prod non vide : soit suite du nom, soit nouveau produit
            if nxt['prod'] and not is_header(nxt['prod']) and not is_titre(nxt['prod']):
                # Suite du nom si : une seule "petite" partie de nom (1-2 mots sans chiffres)
                # et pas de données numériques sur cette ligne
                is_name_continuation = (
                    not nxt['qty'] and not nxt['mon'] and not nxt['cmd']
                    and len(nxt['prod'].split()) <= 2
                    and not re.search(r'\d', nxt['prod'])
                )
                if is_name_continuation:
                    row['prod'] = row['prod'] + ' ' + nxt['prod']
                    j += 1
                    continue
                break
            if nxt['qty'] and not row['qty']:
                row['qty'] = nxt['qty']
            elif nxt['qty']:
                row['qty'] = row['qty'] + ' ' + nxt['qty']
            if nxt['mon'] and not row['mon']:
                row['mon'] = nxt['mon']
            if nxt['cmd'] and not row['cmd']:
                row['cmd'] = nxt['cmd']
            elif nxt['cmd']:
                row['cmd'] = row['cmd'] + ' ' + nxt['cmd']
            j += 1
            # Après avoir les données, continuer encore une ligne pour attraper
            # une éventuelle suite de nom (ex: "rouge" après les données de la salade)
            if row['qty'] and row['mon']:
                # Regarder si la prochaine ligne est une suite de nom
                if j < len(segs):
                    nxt2 = segs[j]
                    if (nxt2['prod'] and not nxt2['qty'] and not nxt2['mon'] and not nxt2['cmd']
                            and len(nxt2['prod'].split()) <= 2
                            and not re.search(r'\d', nxt2['prod'])
                            and not is_header(nxt2['prod'])):
                        row['prod'] = row['prod'] + ' ' + nxt2['prod']
                        j += 1
                break

        i = j if j > i + 1 else i + 1

        # Nettoyer la qty (parfois "3 u." accumule "3 u. u." si l'unité est répétée)
        row['qty'] = re.sub(r'\b(\w+\.?)\s+\1\b', r'\1', row['qty'])

        if row['qty'] and re.search(r'\d', row['qty']):
            rows_raw.append(row)

    # Identifier paniers vs produits
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

    # Calculer cells
    for r in rows:
        parts = r['qty'].split()
        qty_total = float(parts[0]) if parts else 0
        unite = parts[1] if len(parts) > 1 else ''
        r['qty_num'] = fmt(qty_total)
        r['unite']   = unite

        tokens = re.findall(r'1\s*x\s*([\d.]+)\s*(\S+)', r['cmd'])
        cells  = {p['key']: '' for p in paniers}

        # Identifier le panier de chaque token par correspondance de division exacte
        used_keys = set()
        for q, u in tokens:
            qty_cmd = float(q)
            best, best_err = None, float('inf')
            for pan in paniers:
                if pan['key'] in used_keys:
                    continue
                err = abs(qty_cmd / pan['n'] - round(qty_cmd / pan['n'], 4))
                if err < best_err:
                    best_err, best = err, pan
            if best:
                cells[best['key']] = f"{fmt(qty_cmd / best['n'])} {u}"
                used_keys.add(best['key'])

        r['cells'] = cells

    return rows, paniers


def build_new_page(rows, paniers):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))

    W_PROD = 155.0
    W_QTY  = 52.0
    W_MON  = 46.0
    n_p    = len(paniers)
    W_PAN  = (PAGE_RIGHT - MARGIN - W_PROD - W_QTY - W_MON) / max(n_p, 1)

    xP    = MARGIN
    xQ    = xP + W_PROD
    xM    = xQ + W_QTY
    xPans = [xM + W_MON + i * W_PAN for i in range(n_p)]

    def cx(x0, w): return x0 + w / 2

    # Fond blanc
    c.setFillColor(WHITE)
    c.rect(MARGIN, MARGIN, PAGE_RIGHT - MARGIN, PAGE_H - 2*MARGIN, fill=1, stroke=0)

    # Titre
    c.setFillColor(DARK)
    c.setFont(FONT, 22.5)
    c.drawCentredString(PAGE_W/2, PAGE_H - 42.5, "1 vente")

    # En-tête
    HDR_Y = PAGE_H - 76.5
    HDR_H = 21.0
    c.setFillColor(HEADER_BG)
    c.rect(MARGIN, HDR_Y, PAGE_RIGHT - MARGIN, HDR_H, fill=1, stroke=0)

    hy = HDR_Y + 6
    c.setFillColor(WHITE)
    c.setFont(FONT, 7.6)
    c.drawString(xP + 4, hy, "PRODUIT")
    c.drawCentredString(cx(xQ, W_QTY), hy, "QUANTITÉ")
    c.drawCentredString(cx(xM, W_MON), hy, "MONTANT")
    for i, pan in enumerate(paniers):
        c.drawCentredString(cx(xPans[i], W_PAN), hy, f"{pan['n']} {pan['label'].upper()}")

    # Séparateurs en-tête
    c.setStrokeColor(SEP_HDR)
    c.setLineWidth(0.4)
    for xs in [xQ, xM] + xPans:
        c.line(xs, HDR_Y, xs, HDR_Y + HDR_H)

    # Lignes de données
    ROW_H = 21.7
    S_ROW = 8.1
    S_QTY = 11.7
    cur_y = HDR_Y

    def draw_prod(text, ymid, max_w):
        c.setFont(FONT, S_ROW)
        if c.stringWidth(text, FONT, S_ROW) <= max_w:
            c.drawString(xP + 4, ymid - S_ROW/2 + 1, text)
            return
        # Coupure au ' / '
        if ' / ' in text:
            p1, p2 = text.split(' / ', 1)
            c.drawString(xP + 4, ymid + 1.5,        p1 + ' /')
            c.drawString(xP + 4, ymid - S_ROW - 0.5, p2)
        else:
            words = text.split()
            l1 = ''
            for w in words:
                t = (l1 + ' ' + w).strip()
                if c.stringWidth(t, FONT, S_ROW) <= max_w:
                    l1 = t
                else:
                    break
            l2 = text[len(l1):].strip()
            c.drawString(xP + 4, ymid + 1.5,        l1)
            c.drawString(xP + 4, ymid - S_ROW - 0.5, l2)

    for idx, row in enumerate(rows):
        bg = ROW_LIGHT if idx % 2 == 0 else ROW_DARK
        rb = cur_y - ROW_H
        ym = rb + ROW_H / 2

        c.setFillColor(bg)
        c.rect(MARGIN, rb, PAGE_RIGHT - MARGIN, ROW_H, fill=1, stroke=0)
        c.setFillColor(DARK)

        draw_prod(row['prod'], ym, W_PROD - 8)

        # Quantité : grand chiffre + petite unité
        c.setFont(FONT_BOLD, S_QTY)
        qw = c.stringWidth(row['qty_num'], FONT_BOLD, S_QTY)
        uw = c.stringWidth(row['unite'],   FONT,      S_ROW)
        qx = cx(xQ, W_QTY) - (qw + 3 + uw)/2
        c.drawString(qx, ym - S_QTY/2 + 1, row['qty_num'])
        c.setFont(FONT, S_ROW)
        c.drawString(qx + qw + 3, ym - S_ROW/2 + 1, row['unite'])

        # Montant
        c.setFont(FONT, S_ROW)
        c.drawCentredString(cx(xM, W_MON), ym - S_ROW/2 + 1, row['mon'])

        # Cellules paniers
        for i, pan in enumerate(paniers):
            val = row['cells'].get(pan['key'], '')
            if val:
                c.setFont(FONT, S_ROW)
                c.drawCentredString(cx(xPans[i], W_PAN), ym - S_ROW/2 + 1, val)

        # Séparateurs verticaux
        c.setStrokeColor(SEP)
        c.setLineWidth(0.4)
        for xs in [xQ, xM] + xPans:
            c.line(xs, rb, xs, cur_y)

        cur_y = rb

    c.save()
    packet.seek(0)
    return packet


def transformer_pdf(input_path, output_path=None):
    input_path  = Path(input_path)
    output_path = Path(output_path) if output_path else \
                  input_path.parent / (input_path.stem + "_modifie.pdf")

    print(f"Lecture de : {input_path}")
    rows, paniers = extract_table_data(input_path)

    if not paniers:
        print("ERREUR : aucun panier 'Panier de la semaine' trouvé.")
        sys.exit(1)

    for p in paniers:
        print(f"  Panier {p['label']} : {p['n']} unité(s)")
    print(f"  Produits : {len(rows)}")
    for r in rows:
        print(f"    {r['prod']} | {r['qty_num']} {r['unite']} | {r['mon']} | {r['cells']}")

    new_page = build_new_page(rows, paniers)

    original = PdfReader(str(input_path))
    writer   = PdfWriter()
    writer.add_page(PdfReader(new_page).pages[0])
    for page in original.pages[1:]:
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)
    print(f"\nPDF enregistré : {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)
    transformer_pdf(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
