"""Petits utilitaires de nettoyage et de formatage de texte."""

import re


def clean(s):
    """Normalise les espaces (insécables, retours ligne) et coupe les bords."""
    if not s:
        return ''
    return re.sub(r'\s+', ' ', s.replace('\xa0', ' ').replace('\n', ' ')).strip()


def fmt(val):
    """Formate un nombre sans décimales superflues (5.0 → "5", 0.40 → "0.4")."""
    if val == int(val):
        return str(int(val))
    return f"{val:.2f}".rstrip('0').rstrip('.')
