"""
Rendu de la feuille paniers au format A5 paysage (PDF puis, au besoin, PNG).

Une seule taille de police (S = 16 pt × scale) pour tout le contenu ; seules
les unités (kg, bte, u.) sont affichées en S × 0.65, en police normale.
"""

import io
import sys
from pathlib import Path

from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import Color
from reportlab.lib.pagesizes import A5
from reportlab.pdfgen import canvas

BLACK = Color(0, 0, 0)
WHITE = Color(1, 1, 1)

PAGE_H, PAGE_W = A5   # paysage : 595.28 x 419.53 pt
MARGIN     = 10.0
PAGE_RIGHT = PAGE_W - MARGIN

FONT      = "Helvetica"
FONT_BOLD = "Helvetica-Bold"


def build_new_page(rows, paniers, titre, avec_montant=False, scale=1.0):
    """
    Construit la feuille paniers au format A5 paysage.
    scale : multiplicateur global des tailles de police (défaut 1.0).
    Retourne un BytesIO contenant le PDF d'une page.
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

    def cx(x0, w):
        return x0 + w / 2

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
        c.drawCentredString(cx(xPans[i], W_PAN), hy, f"{pan['n']} {pan['label'].upper()}S")

    c.setStrokeColor(BLACK)
    c.setLineWidth(0.4)
    c.line(MARGIN, HDR_Y, PAGE_RIGHT, HDR_Y)   # trait bas de l'en-tête
    for xs in sep_xs:
        c.line(xs, HDR_Y, xs, HDR_Y + HDR_H)   # séparateurs verticaux

    # Lignes de données
    def wrap_prod(text, max_w, s=None):
        """Découpe le nom produit en lignes sans couper les mots."""
        if s is None:
            s = S
        words = text.split()
        lines, cur = [], ''
        for w in words:
            t = (cur + ' ' + w).strip()
            if c.stringWidth(t, FONT, s) <= max_w:
                cur = t
            else:
                if cur:
                    lines.append(cur)
                cur = w
        if cur:
            lines.append(cur)
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


def _page_pdf_bytes(rows, paniers, titre, avec_montant, scale):
    """Construit la page et renvoie les octets du PDF final (une page)."""
    page = build_new_page(rows, paniers, titre, avec_montant=avec_montant, scale=scale)
    writer = PdfWriter()
    writer.add_page(PdfReader(page).pages[0])
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def write_pdf(output_path, rows, paniers, titre, avec_montant=False, scale=1.0):
    """Écrit la feuille au format PDF."""
    data = _page_pdf_bytes(rows, paniers, titre, avec_montant, scale)
    Path(output_path).write_bytes(data)


def write_png(output_path, rows, paniers, titre, avec_montant=False, scale=1.0, dpi=200):
    """Écrit la feuille au format PNG (PDF en mémoire converti via pdf2image)."""
    try:
        from pdf2image import convert_from_bytes
    except ImportError:
        print("ERREUR : pip install pdf2image poppler pour la sortie PNG.")
        sys.exit(1)
    data = _page_pdf_bytes(rows, paniers, titre, avec_montant, scale)
    images = convert_from_bytes(data, dpi=dpi)
    images[0].save(str(output_path))
