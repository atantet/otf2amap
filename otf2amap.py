#!/usr/bin/env python3
"""
Transforme un PDF de ventes OTF en feuille de préparation des paniers AMAP.

Colonnes générées : DATE | TOTAL | N PETIT | N MOYEN | [N GRAND]
- Date extraite de la page 2 du PDF source
- Lignes "Panier de la semaine" filtrées
- Format A5 paysage, noir et blanc

Usage : python3 otf2amap.py entree.pdf [sortie.pdf] [--montant] [--scale 1.0]
"""

import sys
import re
import io
from pathlib import Path
from collections import defaultdict
import itertools

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5
from reportlab.lib.colors import Color

SORTIE = "feuille_paniers_amap"

BLACK = Color(0, 0, 0)
WHITE = Color(1, 1, 1)

PAGE_H, PAGE_W = A5   # paysage : 595.28 x 419.53 pt
MARGIN     = 10.0
PAGE_RIGHT = PAGE_W - MARGIN

FONT      = "Helvetica"
FONT_BOLD = "Helvetica-Bold"


# ── Utilitaires ───────────────────────────────────────────────────────────────

def clean(s):
    if not s: return ''
    return re.sub(r'\s+', ' ', s.replace('\xa0', ' ').replace('\n', ' ')).strip()


def fmt(val):
    if val == int(val): return str(int(val))
    return f"{val:.2f}".rstrip('0').rstrip('.')


# ── Extraction PDF ────────────────────────────────────────────────────────────

def extract_date_from_page2(pdf_path):
    """Retourne la date de retrait (DD/MM/YYYY) depuis la page 2, ou None."""
    with pdfplumber.open(pdf_path) as pdf:
        if len(pdf.pages) < 2:
            return None
        text = pdf.pages[1].extract_text() or ''
    m = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', text)
    return m.group(1) if m else None


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


def extract_table_data(pdf_path):
    """
    Extrait produits et paniers depuis la page 1 du PDF.
    Gère deux layouts :
      Layout A : colonnes QUANTITÉ et MONTANT distinctes (x 198–297)
      Layout B : tout fusionné après la colonne PRODUIT
    Retourne (rows, paniers).
    """
    X_PROD_END  = 198.0
    X_QTY_START = 198.0; X_QTY_END = 252.0
    X_MON_START = 252.0; X_MON_END = 297.0
    X_CMD_START = 297.0

    with pdfplumber.open(pdf_path) as pdf:
        words = pdf.pages[0].extract_words(x_tolerance=3, y_tolerance=3)

    by_y = defaultdict(list)
    for w in words:
        by_y[round(w['top'] / 4) * 4].append(w)

    def col(ws, x0, x1): return [w for w in ws if x0 < w['x0'] < x1]
    def txt(ws): return clean(' '.join(w['text'] for w in ws))

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
            i += 1; continue
        if s['mon'] and re.match(r'^\d+\s+vente', s['mon']) and not s['prod']:
            i += 1; continue
        if not s['prod']:
            i += 1; continue

        row = {'prod': s['prod'], 'qty': s['qty'], 'mon': s['mon'], 'cmd': s['cmd']}

        if i > 0:
            prev = segs[i - 1]
            if not prev['prod'] and not prev['mon']:
                if prev['qty'] and not prev['cmd']:
                    row['qty'] = (prev['qty'] + ' ' + row['qty']).strip() if row['qty'] else prev['qty']
                elif prev['cmd'] and not prev['qty']:
                    row['cmd'] = (prev['cmd'] + ' ' + row['cmd']).strip() if row['cmd'] else prev['cmd']

        j = i + 1
        while j < len(segs):
            nxt = segs[j]
            if nxt['prod'] and not is_header(nxt['prod']) and not is_titre(nxt['prod']):
                if is_name_only(nxt):
                    row['prod'] += ' ' + nxt['prod']
                    j += 1; continue
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

        if not row['qty'] and row['cmd']:
            row['qty'], row['mon'], row['cmd'] = parse_raw_cmd(row['cmd'])

        row['qty'] = re.sub(r'\b(\w+\.?)\s+\1\b', r'\1', row['qty'])

        if row['qty'] and re.search(r'\d', row['qty']):
            rows_raw.append(row)

    PANIER_KEYS = [('petit', 'Petit'), ('moyen', 'Moyen'), ('grand', 'Grand')]
    ORDER = {'petit': 0, 'moyen': 1, 'grand': 2}
    paniers, rows = [], []

    for r in rows_raw:
        low = r['prod'].lower()
        matched = next(((k, l) for k, l in PANIER_KEYS if 'panier de la semaine' in low and k in low), None)
        if matched:
            key, lbl = matched
            nums = re.findall(r'\d+(?:\.\d+)?', r['qty'])
            paniers.append({'key': key, 'label': lbl, 'n': int(float(nums[0])) if nums else 1})
        else:
            rows.append(r)

    paniers.sort(key=lambda p: ORDER.get(p['key'], 99))

    for r in rows:
        parts = r['qty'].split()
        qty_total = float(parts[0]) if parts else 0
        r['qty_num'] = fmt(qty_total)
        r['unite']   = parts[1] if len(parts) > 1 else ''

        tokens = re.findall(r'1\s*x\s*([\d.]+)\s*(\S+)', r['cmd'])
        cells  = {p['key']: '' for p in paniers}
        qtys   = [float(q) for q, u in tokens]
        units  = [u for q, u in tokens]

        if qtys:
            best_score  = None
            best_assign = list(range(len(qtys)))
            for perm in itertools.permutations(range(len(paniers)), len(qtys)):
                ratios = [qtys[i] / paniers[perm[i]]['n'] for i in range(len(qtys))]
                score  = tuple(sum(abs(r - round(r, dp)) for r in ratios) for dp in [2, 1, 0])
                if best_score is None or score < best_score:
                    best_score  = score
                    best_assign = perm
            for i, (q, u) in enumerate(zip(qtys, units)):
                pan = paniers[best_assign[i]]
                cells[pan['key']] = f"{fmt(q / pan['n'])} {u}"

        r['cells'] = cells

    return rows, paniers


# ── Génération PDF ────────────────────────────────────────────────────────────

def build_new_page(rows, paniers, titre, avec_montant=False, scale=1.0):
    """
    Construit la feuille paniers au format A5 paysage.
    scale : multiplicateur global des tailles de police (défaut 1.0).
    """
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))

    S      = 16.0 * scale
    S_UNIT = S * 0.65

    W_QTY  = 64.0
    W_MON  = 40.0 if avec_montant else 0.0
    W_PROD = 180.0
    n_p    = len(paniers)
    W_PAN  = (PAGE_RIGHT - MARGIN - W_PROD - W_QTY - W_MON) / max(n_p, 1)

    xP     = MARGIN
    xQ     = xP + W_PROD
    xM     = xQ + W_QTY
    xPans  = [xM + W_MON + i * W_PAN for i in range(n_p)]
    sep_xs = [xQ] + ([xM] if avec_montant else []) + xPans

    def cx(x0, w): return x0 + w / 2

    MARGIN_TOP = 10.0
    HDR_H      = S + 18
    ROW_H_MIN  = S + 12

    HDR_Y       = PAGE_H - MARGIN_TOP - HDR_H
    available_h = HDR_Y - MARGIN
    ROW_H       = max(ROW_H_MIN, available_h / max(len(rows), 1))

    c.setFillColor(WHITE)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    hy = HDR_Y + (HDR_H - S) / 2
    c.setFillColor(BLACK)
    c.setFont(FONT_BOLD, S)
    c.drawString(xP + 4, hy, titre)
    c.drawCentredString(cx(xQ, W_QTY), hy, "TOTAL")
    if avec_montant:
        c.drawCentredString(cx(xM, W_MON), hy, "MONTANT")
    for i, pan in enumerate(paniers):
        c.drawCentredString(cx(xPans[i], W_PAN), hy, f"{pan['n']} {pan['label'].upper()}")

    c.setStrokeColor(BLACK)
    c.setLineWidth(0.4)
    c.line(MARGIN, HDR_Y, PAGE_RIGHT, HDR_Y)
    for xs in sep_xs:
        c.line(xs, HDR_Y, xs, HDR_Y + HDR_H)

    def wrap_prod(text, max_w):
        words = text.split()
        lines, cur = [], ''
        for w in words:
            t = (cur + ' ' + w).strip()
            if c.stringWidth(t, FONT, S) <= max_w:
                cur = t
            else:
                if cur: lines.append(cur)
                cur = w
        if cur: lines.append(cur)
        return lines

    def draw_prod(text, rb, row_h, max_w):
        c.setFont(FONT, S)
        lines  = wrap_prod(text, max_w)
        line_h = S * 1.3
        y_top  = rb + row_h / 2 + len(lines) * line_h / 2 - S
        for i, line in enumerate(lines):
            c.drawString(xP + 4, y_top - i * line_h, line)

    def draw_qty(num, unit, xc, wc, ym):
        c.setFont(FONT_BOLD, S)
        qw = c.stringWidth(num,  FONT_BOLD, S)
        uw = c.stringWidth(unit, FONT,      S_UNIT)
        gap = 3.0
        qx  = cx(xc, wc) - (qw + gap + uw) / 2
        c.drawString(qx, ym - S / 2, num)
        c.setFont(FONT, S_UNIT)
        c.drawString(qx + qw + gap, ym - S_UNIT / 2, unit)

    cur_y = HDR_Y
    for row in rows:
        rb = cur_y - ROW_H
        ym = rb + ROW_H / 2

        c.setFillColor(BLACK)
        draw_prod(row['prod'], rb, ROW_H, W_PROD - 8)
        draw_qty(row['qty_num'], row['unite'], xQ, W_QTY, ym)

        if avec_montant:
            c.setFont(FONT, S)
            c.drawCentredString(cx(xM, W_MON), ym - S / 2, row['mon'])

        for i, pan in enumerate(paniers):
            val = row['cells'].get(pan['key'], '')
            if val:
                parts = val.rsplit(' ', 1)
                if len(parts) == 2:
                    draw_qty(parts[0], parts[1], xPans[i], W_PAN, ym)
                else:
                    c.setFont(FONT_BOLD, S)
                    c.drawCentredString(cx(xPans[i], W_PAN), ym - S / 2, val)

        c.setStrokeColor(BLACK)
        c.setLineWidth(0.4)
        c.line(MARGIN, cur_y, PAGE_RIGHT, cur_y)
        for xs in sep_xs:
            c.line(xs, rb, xs, cur_y)

        cur_y = rb

    c.save()
    packet.seek(0)
    return packet


# ── Point d'entrée ────────────────────────────────────────────────────────────

def transformer_pdf(input_path, output_path=None, avec_montant=False, scale=1.0):
    input_path  = Path(input_path)
    output_path = Path(output_path) if output_path else input_path.parent / (SORTIE + ".pdf")

    print(f"Lecture de : {input_path}")
    titre = extract_date_from_page2(input_path) or "Ventes"
    print(f"  Date : {titre}")

    rows, paniers = extract_table_data(input_path)
    if not paniers:
        print("ERREUR : aucun panier 'Panier de la semaine' trouvé.")
        sys.exit(1)

    for p in paniers:
        print(f"  Panier {p['label']} : {p['n']} unité(s)")
    print(f"  Produits : {len(rows)}")
    for r in rows:
        print(f"    {r['prod']} | {r['qty_num']} {r['unite']} | {r['cells']}")

    page = build_new_page(rows, paniers, titre, avec_montant=avec_montant, scale=scale)
    writer = PdfWriter()
    writer.add_page(PdfReader(page).pages[0])
    with open(output_path, "wb") as f:
        writer.write(f)
    print(f"PDF enregistré : {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    args = sys.argv[1:]
    avec_montant = "--montant" in args
    args = [a for a in args if a != "--montant"]

    scale = 1.0
    for i, a in enumerate(args):
        if a.startswith("--scale="):
            scale = float(a.split("=")[1])
        elif a == "--scale" and i + 1 < len(args):
            scale = float(args[i + 1])
    args = [a for a in args if not a.startswith("--scale")]

    transformer_pdf(args[0], args[1] if len(args) > 1 else None,
                    avec_montant=avec_montant, scale=scale)
