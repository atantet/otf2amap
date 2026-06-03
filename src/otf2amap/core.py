"""Orchestration : du PDF source aux lignes prêtes à rendre."""

from .allocate import allocate
from .extract import (
    extract_date_from_page2,
    extract_page2_quantities,
    parse_page1,
)


def build_sheet(input_path):
    """
    Lit un PDF OuvreTaFerme et retourne (rows, paniers, titre) :
      - rows    : produits avec qty_num, unite et cells (quantité par panier)
      - paniers : liste {key, label, n} triée petit/moyen/grand
      - titre   : date de retrait (DD/MM/YYYY) ou "Ventes"

    Lève ValueError si aucun panier "Panier de la semaine" n'est trouvé.
    """
    titre = extract_date_from_page2(input_path) or "Ventes"
    rows, paniers = parse_page1(input_path)
    if not paniers:
        raise ValueError("aucun panier 'Panier de la semaine' trouvé")
    page2_data = extract_page2_quantities(input_path, paniers)
    allocate(rows, paniers, page2_data)
    return rows, paniers, titre
