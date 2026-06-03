"""
otf2amap — convertit un PDF de ventes OuvreTaFerme en feuille de préparation
des paniers AMAP (A5 paysage, noir et blanc).

API publique :
    build_sheet(input_path)              -> (rows, paniers, titre)
    build_text_table(rows, paniers, titre, mode)
    write_pdf / write_png(output_path, rows, paniers, titre, ...)
"""

from .core import build_sheet
from .render import write_pdf, write_png
from .text import build_text_table

__all__ = [
    "build_sheet",
    "build_text_table",
    "write_pdf",
    "write_png",
]

__version__ = "1.0.0"
