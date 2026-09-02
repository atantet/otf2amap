"""Tests unitaires des briques pures (sans PDF)."""

from datetime import datetime
from email.message import EmailMessage

from otf2amap.allocate import allocate, match_produit
from otf2amap.extract import cle_tri_produit, parse_raw_cmd
from otf2amap.legumes import renommer_avec_paniers
from otf2amap.mailbox import extract_contrat, parse_date_livraison
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


# ── mailbox.parse_date_livraison ──────────────────────────────────────────────

def test_parse_date_livraison_objet_amapj():
    # objet réel AmapJ, avec jour de semaine
    s = "AMAP du Bocage - Feuille de livraison du mercredi 24 juin 2026"
    assert parse_date_livraison(s) == datetime(2026, 6, 24)


def test_parse_date_livraison_sans_jour_de_semaine_et_accents():
    assert parse_date_livraison("… du 17 juin 2026") == datetime(2026, 6, 17)
    assert parse_date_livraison("… du mercredi 5 août 2026") == datetime(2026, 8, 5)


def test_parse_date_livraison_invalide():
    assert parse_date_livraison("pas de date") is None
    assert parse_date_livraison("le 32 juin 2026") is None
    assert parse_date_livraison(None) is None


# ── mailbox.extract_contrat ───────────────────────────────────────────────────

def _mail(corps):
    msg = EmailMessage()
    msg["From"] = "saint-malo@m.amapj.fr"
    msg["Subject"] = "AMAP du Bocage - Feuille de livraison du mercredi 25 février 2026"
    msg.set_content(corps)
    return msg


_CORPS_LEGUMES = """AMAP du Bocage-Saint Malo

Bonjour ,

Vous trouverez ci joint la feuille de livraison pour le mercredi 25 février 2026

Nom du contrat : Légumes (octobre 2025 - mars 2026)
Nom du producteur : EARL La Petite Claye des champs (légumes)
"""

_CORPS_POMMES = _CORPS_LEGUMES.replace(
    "Légumes (octobre 2025 - mars 2026)", "Pommes (octobre 2025 à mars 2026)"
)


def test_extract_contrat_distingue_les_contrats():
    # objet + expéditeur identiques : seul le corps distingue les contrats
    assert extract_contrat(_mail(_CORPS_LEGUMES)) == "Légumes (octobre 2025 - mars 2026)"
    assert extract_contrat(_mail(_CORPS_POMMES)) == "Pommes (octobre 2025 à mars 2026)"


def test_extract_contrat_absent():
    assert extract_contrat(_mail("Bonjour, pas de ligne contrat ici.")) is None


def test_mail_transfere_manuellement():
    # Transfert manuel : objet préfixé « Tr: », corps original cité.
    sujet = "Tr: AMAP du Bocage - Feuille de livraison du mercredi 25 février 2026"
    corps = (
        "---------- Message transféré ----------\n"
        "De : AMAP du Bocage <saint-malo@m.amapj.fr>\n"
        "Objet : Feuille de livraison du mercredi 25 février 2026\n\n"
        "Vous trouverez ci joint la feuille de livraison pour le mercredi 25 février 2026\n"
        "Nom du contrat : Légumes (octobre 2025 - mars 2026)\n"
    )
    assert parse_date_livraison(sujet) == datetime(2026, 2, 25)
    assert extract_contrat(_mail(corps)) == "Légumes (octobre 2025 - mars 2026)"


# ── legumes.renommer_avec_paniers ─────────────────────────────────────────────

def test_renommer_avec_paniers_ajoute_le_postfixe(tmp_path):
    src = tmp_path / "distri-Légumes-24-06-2026.xls"
    src.write_bytes(b"contenu")
    cible = renommer_avec_paniers(src, {"petit": 14, "moyen": 8, "grand": 0})
    assert cible.name == "distri-Légumes-24-06-2026_14_8_0.xls"
    assert cible.exists() and not src.exists()
    assert cible.read_bytes() == b"contenu"


# ── extract.parse_raw_cmd ─────────────────────────────────────────────────────

def test_parse_raw_cmd_layout_b():
    raw = "5.84 kg 34,47 € 1 x 2.1 kg 1 x 3.74 kg"
    qty, mon, cmd = parse_raw_cmd(raw)
    assert qty == "5.84 kg"
    assert mon == "34,47 €"
    assert cmd == "1 x 2.1 kg 1 x 3.74 kg"


def test_parse_raw_cmd_sans_correspondance():
    assert parse_raw_cmd("texte libre") == ("", "", "texte libre")


# ── extract.cle_tri_produit ───────────────────────────────────────────────────

def test_cle_tri_produit_regroupe_les_calibres():
    # Cas réel de la feuille : le tri alphabétique brut sépare « Chou » de
    # « Demi chou » ; la clé doit les rendre consécutifs, calibre standard avant.
    produits = [
        "Aubergine",
        "Chou nouveau à l'unité / Pointu",
        "Concombre à l'unité / Court lisse",
        "Concombre petit à l'unité / Court lisse",
        "Courgette petite / Verte",
        "Demi chou nouveau à l'unité / Pointu",
        "Oignon en botte",
        "Oignon en petite botte",
        "Salade à l'unité / Oreille du diable",
        "Salade petite à l'unité / Oreille du diable",
    ]
    ordonne = sorted(produits, key=cle_tri_produit)
    assert ordonne == [
        "Aubergine",
        "Chou nouveau à l'unité / Pointu",
        "Demi chou nouveau à l'unité / Pointu",
        "Concombre à l'unité / Court lisse",
        "Concombre petit à l'unité / Court lisse",
        "Courgette petite / Verte",
        "Oignon en botte",
        "Oignon en petite botte",
        "Salade à l'unité / Oreille du diable",
        "Salade petite à l'unité / Oreille du diable",
    ]


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
    assert "11 PETITS" in lines[0] and "8 MOYENS" in lines[0]
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
