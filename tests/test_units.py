"""Tests unitaires des briques pures (sans PDF)."""

from otf2amap.allocate import allocate, match_produit
from otf2amap.extract import parse_raw_cmd
from otf2amap.naming import prefixe_semaine
from otf2amap.text import build_text_table
from otf2amap.util import clean, fmt

# ── util ──────────────────────────────────────────────────────────────────────

def test_clean_normalise_les_espaces():
    assert clean("  a\xa0 b\n c  ") == "a b c"
    assert clean(None) == ""
    assert clean("") == ""


def test_fmt_supprime_les_decimales_superflues():
    assert fmt(5.0) == "5"
    assert fmt(0.40) == "0.4"
    assert fmt(0.65) == "0.65"
    assert fmt(12.15) == "12.15"


# ── naming ────────────────────────────────────────────────────────────────────

def test_prefixe_semaine():
    # 03/06/2026 (mercredi) → semaine ISO 23 de 2026
    assert prefixe_semaine("03/06/2026") == "2026_S23"


def test_prefixe_semaine_invalide():
    assert prefixe_semaine("pas une date") is None
    assert prefixe_semaine(None) is None


# ── extract.parse_raw_cmd ─────────────────────────────────────────────────────

def test_parse_raw_cmd_layout_b():
    raw = "5.84 kg 34,47 € 1 x 2.1 kg 1 x 3.74 kg"
    qty, mon, cmd = parse_raw_cmd(raw)
    assert qty == "5.84 kg"
    assert mon == "34,47 €"
    assert cmd == "1 x 2.1 kg 1 x 3.74 kg"


def test_parse_raw_cmd_sans_correspondance():
    assert parse_raw_cmd("texte libre") == ("", "", "texte libre")


# ── allocate ──────────────────────────────────────────────────────────────────

def test_match_produit_exact_et_partiel():
    page2 = {"Fève": {"petit": (1.0, "kg")}}
    assert match_produit("Fève", page2) == {"petit": (1.0, "kg")}
    # partiel : le nom page 2 est inclus dans le nom page 1
    assert match_produit("Fève des marais", page2) == {"petit": (1.0, "kg")}
    assert match_produit("Carotte", page2) is None


def test_allocate_permutation_optimale_betterave():
    # Cas du fichier de notes : tokens 2.8 et 2.6, paniers n=7 et n=13.
    # 2.8/7 = 0.4 (rond) doit aller au petit, 2.6/13 = 0.2 au moyen.
    rows = [{"prod": "Betterave", "qty": "5.4 kg", "cmd": "1 x 2.8 kg 1 x 2.6 kg"}]
    paniers = [{"key": "petit", "label": "Petit", "n": 7},
               {"key": "moyen", "label": "Moyen", "n": 13}]
    allocate(rows, paniers, page2_data={})
    r = rows[0]
    assert r["qty_num"] == "5.4"
    assert r["unite"] == "kg"
    assert r["cells"]["petit"] == "0.4 kg"
    assert r["cells"]["moyen"] == "0.2 kg"


def test_allocate_lecture_directe_page2():
    rows = [{"prod": "Fève", "qty": "12.15 kg", "cmd": ""}]
    paniers = [{"key": "petit", "label": "Petit", "n": 11},
               {"key": "moyen", "label": "Moyen", "n": 8}]
    page2 = {"Fève": {"petit": (4.95, "kg"), "moyen": (7.2, "kg")}}
    allocate(rows, paniers, page2)
    cells = rows[0]["cells"]
    assert cells["petit"] == "0.45 kg"   # 4.95 / 11
    assert cells["moyen"] == "0.9 kg"    # 7.2 / 8


# ── text ──────────────────────────────────────────────────────────────────────

def test_build_text_table_markdown():
    rows = [{"prod": "Fève", "qty_num": "12.15", "unite": "kg",
             "cells": {"petit": "0.45 kg", "moyen": "0.9 kg"}}]
    paniers = [{"key": "petit", "label": "Petit", "n": 11},
               {"key": "moyen", "label": "Moyen", "n": 8}]
    out = build_text_table(rows, paniers, "03/06/2026", mode="md")
    lines = out.splitlines()
    assert lines[0].startswith("| 03/06/2026")
    assert "11 PETIT" in lines[0] and "8 MOYEN" in lines[0]
    assert set(lines[1]) <= {"|", "-"}      # ligne de séparation
    assert "Fève" in lines[2] and "0.45 kg" in lines[2]
    assert len(lines) == 3                  # en-tête + séparateur + 1 produit


def test_build_text_table_txt_a_des_bordures():
    rows = [{"prod": "Aillet", "qty_num": "27", "unite": "u.",
             "cells": {"petit": "1 u"}}]
    paniers = [{"key": "petit", "label": "Petit", "n": 11}]
    out = build_text_table(rows, paniers, "03/06/2026", mode="txt")
    lines = out.splitlines()
    assert lines[0].startswith("+") and lines[0].endswith("+")
    assert lines[-1].startswith("+")        # bordure de fermeture en mode txt
