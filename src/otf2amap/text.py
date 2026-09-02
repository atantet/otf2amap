"""Rendu du tableau en Markdown ou en texte ASCII."""


def build_text_table(rows, paniers, titre, mode='md'):
    """
    Génère le tableau en Markdown (mode='md') ou texte ASCII (mode='txt').
    """
    pan_hdrs = [f"{p['n']} {p['label'].upper()}S" for p in paniers]
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
