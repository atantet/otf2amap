"""
Interface ligne de commande d'otf2amap.

Usage : otf2amap [entree.pdf] [sortie] [--montant] [--scale 1.0] [--format=png]

  entree.pdf : PDF OuvreTaFerme (optionnel si [otf].dossier_base est configuré)
               → cherche <dossier_base>/<année>/S<semaine>/Ventes.pdf selon la date du jour

Formats de sortie (défaut : PNG) :
  --format=png   image PNG haute résolution (défaut)
  --format=pdf   PDF
  --format=md    tableau Markdown
  --format=txt   tableau texte brut
  (l'extension du fichier de sortie prime sur --format)

Configuration (config.toml cherché depuis le dossier courant) :
  [otf]
  dossier_base = "~/chemin/vers/AMAP"    # racine pour <année>/S<semaine>/Ventes.pdf
  # nom_pdf    = "Ventes.pdf"            # nom du PDF à trouver (défaut : Ventes.pdf)

  [output]
  dossier      = "/chemin/vers/dossier"  # dossier de sortie
  format       = "png"                   # format de sortie
  nom          = "mon_fichier"           # nom sans extension (remplace le défaut)
  drive_remote = "gdrive:AMAP/feuilles"  # envoi via rclone (optionnel)

Les options ligne de commande prévalent sur config.toml.
config.toml doit exister ; ses variables sont toutes optionnelles.
"""

import sys
from datetime import datetime
from pathlib import Path

from .config import load_config, load_otf_config
from .core import build_sheet
from .extract import extract_date_from_page2
from .naming import SORTIE, dossier_semaine, prefixe_semaine
from .render import write_pdf, write_png
from .text import build_text_table
from .upload import upload_to_remote

FORMATS = ('png', 'pdf', 'md', 'txt')


def transformer(input_path, output_path=None, fmt_out='png',
                avec_montant=False, scale=1.0):
    """
    Lit le PDF source et écrit la feuille paniers.
    fmt_out : 'png' (défaut), 'pdf', 'md', 'txt'.
    """
    input_path = Path(input_path)

    # Déterminer le format depuis l'extension du fichier de sortie si fourni
    if output_path:
        ext = Path(output_path).suffix.lower().lstrip('.')
        if ext in FORMATS:
            fmt_out = ext

    print(f"Lecture de : {input_path}")
    try:
        rows, paniers, titre = build_sheet(input_path)
    except ValueError as e:
        print(f"ERREUR : {e}.")
        sys.exit(1)
    print(f"  Date : {titre}")

    # Nom de sortie par défaut (nécessite le titre pour le préfixe semaine)
    if not output_path:
        prefix = prefixe_semaine(titre) or 'inconnu'
        output_path = input_path.parent / f"{prefix}_{SORTIE}.{fmt_out}"
    output_path = Path(output_path)

    for p in paniers:
        print(f"  Panier {p['label']} : {p['n']} unité(s)")
    print(f"  Produits : {len(rows)}")
    for r in rows:
        print(f"    {r['prod']} | {r['qty_num']} {r['unite']} | {r['cells']}")

    if fmt_out == 'pdf':
        write_pdf(output_path, rows, paniers, titre, avec_montant=avec_montant, scale=scale)
    elif fmt_out == 'png':
        write_png(output_path, rows, paniers, titre, avec_montant=avec_montant, scale=scale)
    elif fmt_out == 'md':
        output_path.write_text(build_text_table(rows, paniers, titre, mode='md'), encoding='utf-8')
    elif fmt_out == 'txt':
        output_path.write_text(build_text_table(rows, paniers, titre, mode='txt'), encoding='utf-8')

    print(f"Fichier enregistré : {output_path}")
    return output_path


def _resolve_output_path(input_path, output_cli, cfg, fmt_out):
    """Détermine le chemin de sortie (CLI > config dossier/nom > défaut)."""
    if output_cli:
        return output_cli
    dossier = cfg.get('dossier')
    nom     = cfg.get('nom')          # nom sans extension
    if not (dossier or nom):
        return None                   # transformer calculera le nom par défaut
    # Extraire la date pour le préfixe semaine si pas de nom explicite
    titre_tmp = extract_date_from_page2(input_path)
    prefix = prefixe_semaine(titre_tmp) if titre_tmp else 'inconnu'
    base = nom if nom else f"{prefix}_{SORTIE}"
    dossier_path = (Path(dossier).expanduser().resolve() if dossier
                    else Path(input_path).expanduser().resolve().parent)
    return str(dossier_path / f"{base}.{fmt_out}")


def _trouver_pdf_semaine(base, date_ref, nom_pdf='Ventes.pdf'):
    """Cherche <base>/<année>/S<semaine>/<nom_pdf> pour la date de référence."""
    dossier = dossier_semaine(date_ref, base)
    candidat = dossier / nom_pdf
    return candidat if candidat.exists() else None


def _ensure_out_dir(output_path, input_path):
    """Crée le dossier de sortie si besoin, après confirmation."""
    if output_path:
        out_dir = Path(output_path).expanduser().resolve().parent
    else:
        out_dir = Path(input_path).expanduser().resolve().parent
    if out_dir.exists():
        return True
    print(f"Le dossier de sortie n'existe pas : {out_dir}")
    reponse = input("Créer ce dossier et ses parents ? [o/N] ").strip().lower()
    if reponse in ('o', 'oui', 'y', 'yes'):
        out_dir.mkdir(parents=True, exist_ok=True)
        print(f"Dossier créé : {out_dir}")
        return True
    print("Annulé.")
    return False


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--help" in argv or "-h" in argv:
        print(__doc__)
        return 0

    avec_montant = "--montant" in argv
    argv = [a for a in argv if a != "--montant"]

    scale = 1.0
    for i, a in enumerate(argv):
        if a.startswith("--scale="):
            scale = float(a.split("=")[1])
        elif a == "--scale" and i + 1 < len(argv):
            scale = float(argv[i + 1])
    argv = [a for a in argv if not a.startswith("--scale")]

    fmt_cli = None
    for a in argv:
        if a.startswith("--format="):
            fmt_cli = a.split("=")[1].lower()
    argv = [a for a in argv if not a.startswith("--format=")]

    # Configuration (CLI > config > défaut)
    cfg = load_config()
    cfg_otf = load_otf_config()
    fmt_out = fmt_cli or cfg.get('format', 'png')

    # Fichier d'entrée : argument explicite ou auto-détection via [otf].dossier_base
    if argv:
        input_path = str(Path(argv[0]).expanduser().resolve())
        output_cli = str(Path(argv[1]).expanduser().resolve()) if len(argv) > 1 else None
    else:
        base = cfg_otf.get('dossier_base')
        if not base:
            print(__doc__)
            return 0
        nom_pdf = cfg_otf.get('nom_pdf', 'Ventes.pdf')
        date_ref = datetime.today()
        pdf = _trouver_pdf_semaine(base, date_ref, nom_pdf)
        if pdf is None:
            dossier = dossier_semaine(date_ref, base)
            print(f"ERREUR : {nom_pdf!r} introuvable dans {dossier}")
            return 1
        input_path = str(pdf)
        output_cli = None

    output_path = _resolve_output_path(input_path, output_cli, cfg, fmt_out)

    if not _ensure_out_dir(output_path, input_path):
        return 0

    final_path = transformer(input_path, output_path,
                             fmt_out=fmt_out, avec_montant=avec_montant, scale=scale)

    # Envoi optionnel vers Google Drive (ou tout remote rclone)
    remote = cfg.get('drive_remote')
    if remote:
        upload_to_remote(final_path, remote)
    return 0


if __name__ == "__main__":
    sys.exit(main())
