#!/usr/bin/env python3
"""
Transforme un PDF de ventes OTF en feuille de préparation des paniers AMAP.

Colonnes générées : DATE | TOTAL | N PETIT | N MOYEN | [N GRAND]
- Date extraite de la page 2 du PDF source
- Lignes "Panier de la semaine" filtrées
- Format A5 paysage, noir et blanc

Usage : python3 otf2amap.py entree.pdf [sortie] [--pdf] [--montant] [--scale 1.0]

Formats de sortie (défaut : PNG) :
  --format=png   image PNG haute résolution (défaut)
  --format=pdf   PDF
  --format=md    tableau Markdown
  --format=txt   tableau texte brut
  (l'extension du fichier de sortie prime sur --format)

Configuration (config.toml dans le même dossier que le script) :
  [output]
  dossier = "/chemin/vers/dossier"   # dossier de sortie
  format  = "png"                    # format de sortie
  nom     = "mon_fichier"            # nom sans extension (remplace le nom par défaut)

Les options ligne de commande prévalent sur config.toml.
config.toml doit exister ; ses variables sont toutes optionnelles.
"""

import sys
import re
import io
from pathlib import Path
from collections import defaultdict
import itertools
from datetime import datetime, timedelta
try:
    import tomllib
except ImportError:
    try:
        import tomli as tomllib
    except ImportError:
        tomllib = None

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


def prefixe_semaine(date_str):
    """
    Calcule le préfixe YYYY_SNN à partir d'une date de panier DD/MM/YYYY.
    La distribution a lieu le mercredi suivant (ou le jour même si c'est mercredi).
    """
    try:
        d = datetime.strptime(date_str, '%d/%m/%Y')
        jours = (2 - d.weekday()) % 7   # 0 si mercredi, sinon jours jusqu'au prochain
        mercredi = d + timedelta(days=jours)
        annee, semaine, _ = mercredi.isocalendar()
        return f"{annee}_S{semaine:02d}"
    except Exception:
        return None


def load_config():
    """
    Charge config.toml depuis le dossier du script.
    Erreur si le fichier n'existe pas.
    Retourne un dict avec les clés trouvées dans [output] : dossier, format, nom.
    """
    config_path = Path(__file__).parent / 'config.toml'
    if not config_path.exists():
        print(f"ERREUR : config.toml introuvable ({config_path})")
        sys.exit(1)
    if tomllib is None:
        print("ERREUR : pip install tomli pour lire config.toml (Python < 3.11)")
        sys.exit(1)
    with open(config_path, 'rb') as f:
        data = tomllib.load(f)
    return data.get('output', {})


# ── Extraction PDF ─────────────────────────────────────────────────────────────

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
        matched = next(((k, l) for k, l in PANIER_KEYS if 'panier de la semaine' in low and k in low), None)
        if matched:
            key, lbl = matched
            nums = re.findall(r'\d+(?:\.\d+)?', r['qty'])
            paniers.append({'key': key, 'label': lbl, 'n': int(float(nums[0])) if nums else 1})
        else:
            rows.append(r)

    paniers.sort(key=lambda p: ORDER.get(p['key'], 99))

    # Calculer les quantités par panier pour chaque produit
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
            # Permutation optimale : score lexicographique sur l'erreur d'arrondi
            # à 2, 1 et 0 décimales — favorise les ratios qté/n les plus ronds.
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

    # Tailles de police : une seule taille S pour tout le contenu ;
    # les unités (kg, bte, u.) sont affichées en S × 0.65.
    S      = 16.0 * scale
    S_UNIT = S * 0.65

    # Largeurs de colonnes (indépendantes de scale)
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

    # Hauteurs
    MARGIN_TOP = 10.0
    HDR_H      = S + 18
    ROW_H_MIN  = S + 12

    HDR_Y       = PAGE_H - MARGIN_TOP - HDR_H
    available_h = HDR_Y - MARGIN

    # Hauteur de ligne : on prend le max entre :
    #  - la hauteur minimale (S + 12)
    #  - la hauteur disponible divisée par le nombre de lignes
    # Puis on s'assure que chaque ligne est assez haute pour son contenu (wrap).
    # Le centrage vertical dans draw_prod est borné à la hauteur de la case.
    ROW_H       = max(ROW_H_MIN, available_h / max(len(rows), 1))

    # Fond blanc
    c.setFillColor(WHITE)
    c.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    # En-tête : trait bas + séparateurs verticaux uniquement (pas de cadre)
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
    c.line(MARGIN, HDR_Y, PAGE_RIGHT, HDR_Y)   # trait bas de l'en-tête
    for xs in sep_xs:
        c.line(xs, HDR_Y, xs, HDR_Y + HDR_H)   # séparateurs verticaux

    # Lignes de données
    def wrap_prod(text, max_w, s=None):
        """Découpe le nom produit en lignes sans couper les mots."""
        if s is None: s = S
        words = text.split()
        lines, cur = [], ''
        for w in words:
            t = (cur + ' ' + w).strip()
            if c.stringWidth(t, FONT, s) <= max_w:
                cur = t
            else:
                if cur: lines.append(cur)
                cur = w
        if cur: lines.append(cur)
        return lines

    def draw_prod(text, rb, row_h, max_w):
        # Réduire la police si le bloc wrappé dépasse la hauteur de la case
        s = S
        while s >= 7:
            lines  = wrap_prod(text, max_w, s)
            line_h = s * 1.3
            if line_h * len(lines) <= row_h - 4:
                break
            s -= 0.5
        c.setFont(FONT, s)
        line_h = s * 1.3
        y_top  = rb + row_h / 2 + len(lines) * line_h / 2 - s
        for i, line in enumerate(lines):
            c.drawString(xP + 4, y_top - i * line_h, line)

    def draw_qty(num, unit, xc, wc, ym):
        """Chiffre gras + unité petite, centrés dans la colonne."""
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

        # Trait de séparation (haut de ligne, après le fond)
        c.setStrokeColor(BLACK)
        c.setLineWidth(0.4)
        c.line(MARGIN, cur_y, PAGE_RIGHT, cur_y)
        for xs in sep_xs:
            c.line(xs, rb, xs, cur_y)

        cur_y = rb

    c.save()
    packet.seek(0)
    return packet


# ── Sorties texte ─────────────────────────────────────────────────────────────

def build_text_table(rows, paniers, titre, mode='md'):
    """
    Génère le tableau en Markdown (mode='md') ou texte ASCII (mode='txt').
    """
    pan_hdrs = [f"{p['n']} {p['label'].upper()}" for p in paniers]
    hdrs  = [titre, 'TOTAL'] + pan_hdrs
    col_w = [max(len(hdrs[0]), max((len(r['prod']) for r in rows), default=0))]
    col_w += [max(len(hdrs[1]), max((len(f"{r['qty_num']} {r['unite']}") for r in rows), default=0))]
    for i, pan in enumerate(paniers):
        col_w.append(max(len(pan_hdrs[i]),
                     max((len(r['cells'].get(pan['key'], '')) for r in rows), default=0)))
    col_w = [w + 2 for w in col_w]

    if mode == 'md':
        def row_str(cells):
            parts = [f" {c:<{col_w[i]-1}}" for i, c in enumerate(cells)]
            return '|' + '|'.join(parts) + '|'
        def sep_line():
            return '|' + '|'.join('-' * col_w[i] for i in range(len(hdrs))) + '|'
    else:  # txt : bordures + et séparateurs | et -
        def row_str(cells):
            parts = [f" {c:<{col_w[i]-1}}" for i, c in enumerate(cells)]
            return '+' + '|'.join(parts) + '+'
        def sep_line():
            return '+' + '+'.join('-' * col_w[i] for i in range(len(hdrs))) + '+'

    lines = [row_str(hdrs), sep_line()]
    for r in rows:
        cells = [r['prod'], f"{r['qty_num']} {r['unite']}"]
        cells += [r['cells'].get(p['key'], '') for p in paniers]
        lines.append(row_str(cells))
    if mode == 'txt':
        lines.append(sep_line())
    return '\n'.join(lines)


# ── Point d'entrée ────────────────────────────────────────────────────────────

def transformer(input_path, output_path=None, fmt_out='png',
                avec_montant=False, scale=1.0):
    """
    fmt_out : 'png' (défaut), 'pdf', 'md', 'txt'
    """
    input_path = Path(input_path)

    # Déterminer le format depuis l'extension du fichier de sortie si fourni
    if output_path:
        ext = Path(output_path).suffix.lower().lstrip('.')
        if ext in ('pdf', 'png', 'md', 'txt'):
            fmt_out = ext

    print(f"Lecture de : {input_path}")
    titre = extract_date_from_page2(input_path) or "Ventes"
    print(f"  Date : {titre}")

    # Nom de sortie par défaut (nécessite le titre pour le préfixe semaine)
    if not output_path:
        prefix = prefixe_semaine(titre) or 'inconnu'
        output_path = input_path.parent / f"{prefix}_{SORTIE}.{fmt_out}"
    output_path = Path(output_path)

    rows, paniers = extract_table_data(input_path)
    if not paniers:
        print("ERREUR : aucun panier 'Panier de la semaine' trouvé.")
        sys.exit(1)

    for p in paniers:
        print(f"  Panier {p['label']} : {p['n']} unité(s)")
    print(f"  Produits : {len(rows)}")
    for r in rows:
        print(f"    {r['prod']} | {r['qty_num']} {r['unite']} | {r['cells']}")

    if fmt_out == 'pdf':
        page = build_new_page(rows, paniers, titre, avec_montant=avec_montant, scale=scale)
        writer = PdfWriter()
        writer.add_page(PdfReader(page).pages[0])
        with open(output_path, "wb") as f:
            writer.write(f)

    elif fmt_out == 'png':
        # Générer le PDF en mémoire puis convertir en PNG via pdf2image
        try:
            from pdf2image import convert_from_bytes
        except ImportError:
            print("ERREUR : pip install pdf2image poppler pour la sortie PNG.")
            sys.exit(1)
        page = build_new_page(rows, paniers, titre, avec_montant=avec_montant, scale=scale)
        writer = PdfWriter()
        writer.add_page(PdfReader(page).pages[0])
        buf = io.BytesIO()
        writer.write(buf)
        images = convert_from_bytes(buf.getvalue(), dpi=200)
        images[0].save(str(output_path))

    elif fmt_out == 'md':
        text = build_text_table(rows, paniers, titre, mode='md')
        output_path.write_text(text, encoding='utf-8')

    elif fmt_out == 'txt':
        text = build_text_table(rows, paniers, titre, mode='txt')
        output_path.write_text(text, encoding='utf-8')

    print(f"Fichier enregistré : {output_path}")


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

    # Format ligne de commande
    fmt_cli = None
    for a in args:
        if a.startswith("--format="):
            fmt_cli = a.split("=")[1].lower()
    args = [a for a in args if not a.startswith("--format=")]

    # Fichier de sortie ligne de commande
    input_path = str(Path(args[0]).expanduser().resolve())
    output_cli = str(Path(args[1]).expanduser().resolve()) if len(args) > 1 else None

    # Charger config.toml
    cfg = load_config()

    # Résolution du format (CLI > config > défaut)
    fmt_out = fmt_cli or cfg.get('format', 'png')

    # Résolution du fichier de sortie
    # Si output_cli fourni : il prime (l'extension peut overrider fmt_out, géré dans transformer)
    # Sinon : construire depuis config (dossier + nom) ou laisser transformer gérer le défaut
    if output_cli:
        output_path = output_cli
    else:
        dossier = cfg.get('dossier')
        nom     = cfg.get('nom')          # nom sans extension
        if dossier or nom:
            # Extraire la date pour le préfixe semaine si pas de nom explicite
            titre_tmp = extract_date_from_page2(input_path)
            prefix = prefixe_semaine(titre_tmp) if titre_tmp else 'inconnu'
            base = nom if nom else f"{prefix}_{SORTIE}"
            dossier_path = Path(dossier).expanduser().resolve() if dossier else Path(input_path).expanduser().resolve().parent
            output_path = str(dossier_path / f"{base}.{fmt_out}")
        else:
            output_path = None  # transformer calculera le nom par défaut

    # Créer le dossier de sortie si nécessaire (avec confirmation)
    if output_path:
        out_dir = Path(output_path).expanduser().resolve().parent
    else:
        out_dir = Path(input_path).expanduser().resolve().parent
    if not out_dir.exists():
        print(f"Le dossier de sortie n'existe pas : {out_dir}")
        reponse = input("Créer ce dossier et ses parents ? [o/N] ").strip().lower()
        if reponse in ('o', 'oui', 'y', 'yes'):
            out_dir.mkdir(parents=True, exist_ok=True)
            print(f"Dossier créé : {out_dir}")
        else:
            print("Annulé.")
            sys.exit(0)

    transformer(input_path, output_path,
                fmt_out=fmt_out, avec_montant=avec_montant, scale=scale)
