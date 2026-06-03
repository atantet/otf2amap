"""
Tests d'intégration sur les vrais PDFs d'exemple.

Le dossier exemples/ n'est pas versionné (cf. .gitignore) : ces tests sont
automatiquement ignorés s'il est absent. Ils vérifient des invariants de
structure plutôt qu'une sortie figée, pour rester robustes à de nouveaux PDFs.
"""

import re
from pathlib import Path

import pytest

from otf2amap import build_sheet, build_text_table

EXEMPLES = Path(__file__).resolve().parent.parent / "exemples"
PDFS = sorted(EXEMPLES.glob("*.pdf")) if EXEMPLES.exists() else []

pytestmark = pytest.mark.skipif(not PDFS, reason="aucun PDF dans exemples/")


@pytest.fixture(params=PDFS, ids=lambda p: p.stem)
def feuille(request):
    rows, paniers, titre = build_sheet(request.param)
    return rows, paniers, titre


def test_paniers_detectes(feuille):
    _, paniers, _ = feuille
    assert paniers, "au moins un panier doit être détecté"
    keys = [p["key"] for p in paniers]
    ordre = ["petit", "moyen", "grand"]
    assert keys == sorted(keys, key=ordre.index)
    for p in paniers:
        assert p["n"] >= 1


def test_chaque_ligne_est_complete(feuille):
    rows, paniers, _ = feuille
    panier_keys = {p["key"] for p in paniers}
    assert rows, "au moins un produit attendu"
    for r in rows:
        assert re.fullmatch(r"\d+(\.\d+)?", r["qty_num"]), r
        assert r["unite"] in ("kg", "bte", "u.", "u"), r["unite"]
        assert set(r["cells"]) == panier_keys
        # Aucune ligne 'Panier de la semaine' ne doit subsister dans les produits
        assert "panier de la semaine" not in r["prod"].lower()


def test_titre_est_une_date_ou_ventes(feuille):
    _, _, titre = feuille
    assert titre == "Ventes" or re.fullmatch(r"\d{2}/\d{2}/\d{4}", titre)


@pytest.mark.parametrize("mode", ["md", "txt"])
def test_rendu_texte_coherent(feuille, mode):
    rows, paniers, titre = feuille
    out = build_text_table(rows, paniers, titre, mode=mode)
    lines = out.splitlines()
    # en-tête + séparateur + une ligne par produit (+ bordure finale en txt)
    expected = 2 + len(rows) + (1 if mode == "txt" else 0)
    assert len(lines) == expected
    assert titre in lines[0]
