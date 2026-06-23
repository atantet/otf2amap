"""Commande `mail2amap` : récupère les tableurs de distribution reçus par mail.

En amont d'otf2amap. Télécharge les pièces jointes des mails de distribution
AmapJ dans <base>/<année>/S<semaine>/, puis affiche le nombre de paniers de
légumes par type (petit / moyen / grand) — ce qui sert à créer la vente sur
OuvreTaFerme.

Usage :
    mail2amap                 # dernière livraison reçue (par défaut)
    mail2amap --tout          # toutes les livraisons trouvées dans la boîte
    mail2amap --date 24/06/2026          # une livraison précise
    mail2amap --expediteur addr@x.fr     # surcharge l'expéditeur à filtrer
                                         #   (ex. mail transféré : From = qui transfère)
    mail2amap --fichier                  # compte les paniers d'un tableur déjà
                                         #   téléchargé, SANS se connecter au mail
    mail2amap --fichier chemin/legumes.xls   # … d'un fichier précis
    mail2amap --fichier --date 24/06/2026    # … du tableur Légumes de cette semaine

Configuration :
  - config.toml [mail] : expediteur, objet_motif, dossier_base (non secrets)
  - .env               : AMAP_IMAP_HOST / AMAP_IMAP_USER / AMAP_IMAP_PASSWORD
                         (voir .env.example ; jamais versionné)
"""

import sys
from datetime import datetime
from pathlib import Path

from .config import load_mail_config
from .legumes import compter_paniers, renommer_avec_paniers
from .mailbox import connect, find_distributions, save_attachments
from .mailenv import REPO_ROOT, ConfigError, load_imap_secrets
from .naming import prefixe_semaine


def _dossier_semaine(date_livraison, base):
    """<base>/<année>/S<NN> à partir de la date de livraison (semaine ISO)."""
    prefixe = prefixe_semaine(date_livraison.strftime("%d/%m/%Y"))   # ex. "2026_S26"
    annee, semaine = prefixe.split("_")                              # "2026", "S26"
    return Path(base).expanduser() / annee / semaine


def _est_legumes(texte):
    """Vrai si `texte` (nom de contrat ou de fichier) désigne le contrat Légumes."""
    t = (texte or "").lower()
    return "légume" in t or "legume" in t


def _trouver_legumes_local(base, date_cible):
    """Cherche un tableur Légumes déjà téléchargé. Renvoie un Path ou None.

    Avec --date : dans le dossier de la semaine correspondante ; sinon dans le
    dossier de semaine le plus récent présent sous `base`.
    """
    base = Path(base).expanduser()
    if date_cible is not None:
        dossiers = [_dossier_semaine(date_cible, base)]
    else:
        dossiers = sorted(
            (d for annee in base.glob("[0-9][0-9][0-9][0-9]") if annee.is_dir()
               for d in annee.glob("S[0-9]*") if d.is_dir()),
            reverse=True,
        )
    for dossier in dossiers:
        for motif in ("*[Ll]égume*.xls", "*[Ll]egume*.xls"):
            trouves = sorted(dossier.glob(motif))
            if trouves:
                return trouves[0]
    return None


def _selectionner(distributions, tout, date_cible):
    """Choisit les mails à traiter selon les options.

    --tout : tous ; --date : ceux de cette date ; défaut : la dernière livraison.
    """
    datees = [d for d in distributions if d["date"] is not None]
    if tout:
        return distributions
    if date_cible is not None:
        return [d for d in datees if d["date"].date() == date_cible.date()]
    if not datees:
        return []
    derniere = max(d["date"] for d in datees)
    return [d for d in datees if d["date"] == derniere]


def _parse_args(argv):
    """Retourne (tout, date_cible, expediteur, fichier_mode, fichier_path).

    fichier_mode : True si --fichier est présent ; fichier_path : chemin explicite
    donné après --fichier, ou None (→ chemin par défaut). Quitte sur erreur.
    """
    tout = "--tout" in argv
    date_cible = None
    expediteur = None
    fichier_mode = False
    fichier_path = None
    for i, a in enumerate(argv):
        if a == "--date" and i + 1 < len(argv):
            try:
                date_cible = datetime.strptime(argv[i + 1], "%d/%m/%Y")
            except ValueError:
                print("ERREUR : --date attend le format JJ/MM/AAAA (ex. 24/06/2026)")
                sys.exit(1)
        elif a == "--expediteur" and i + 1 < len(argv):
            expediteur = argv[i + 1]
        elif a == "--fichier":
            fichier_mode = True
            if i + 1 < len(argv) and not argv[i + 1].startswith("-"):
                fichier_path = argv[i + 1]
    return tout, date_cible, expediteur, fichier_mode, fichier_path


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--help" in argv or "-h" in argv:
        print(__doc__)
        return 0

    tout, date_cible, expediteur_cli, fichier_mode, fichier_path = _parse_args(argv)

    cfg = load_mail_config()
    # --expediteur surcharge la config (utile pour un mail transféré manuellement,
    # dont le From devient l'adresse qui transfère et non saint-malo@m.amapj.fr).
    expediteur = expediteur_cli or cfg.get("expediteur", "saint-malo@m.amapj.fr")
    objet_motif = cfg.get("objet_motif", "Feuille de livraison")
    base = cfg.get("dossier_base") or str(REPO_ROOT.parent)

    # --fichier : compte les paniers d'un tableur déjà téléchargé, sans IMAP.
    if fichier_mode:
        return _mode_local(fichier_path, date_cible, base)

    try:
        secrets = load_imap_secrets()
    except ConfigError as e:
        print(f"ERREUR : {e}")
        return 1

    print(f"Connexion IMAP à {secrets['host']} ({secrets['user']}) …")
    imap = connect(secrets["host"], secrets["user"], secrets["password"])
    try:
        distributions = find_distributions(imap, expediteur, objet_motif)
    finally:
        imap.logout()

    if not distributions:
        print(f"Aucun mail de {expediteur} avec « {objet_motif} » dans l'objet.")
        return 0

    a_traiter = _selectionner(distributions, tout, date_cible)
    if not a_traiter:
        cible = date_cible.strftime("%d/%m/%Y") if date_cible else "?"
        print(f"Aucune livraison à traiter (date demandée : {cible}).")
        return 0

    for d in a_traiter:
        sujet, date_livraison, contrat = d["sujet"], d["date"], d["contrat"]
        if date_livraison is None:
            print(f"AVERTISSEMENT : date illisible dans l'objet, mail ignoré : {sujet!r}")
            continue
        dossier = _dossier_semaine(date_livraison, base)
        chemins = save_attachments(d["message"], dossier)
        libelle = contrat or "contrat inconnu"
        print(f"\nLivraison du {date_livraison:%d/%m/%Y} — {libelle} → {dossier}")
        if not chemins:
            print("  (aucune pièce jointe)")
            continue
        for c in chemins:
            print(f"  téléchargé : {c.name}")
        # Identification Légumes : contrat (corps du mail) prioritaire, repli nom de PJ.
        if _est_legumes(contrat) or any(_est_legumes(c.name) for c in chemins):
            for c in chemins:
                if c.suffix.lower() == ".xls" and (_est_legumes(contrat) or _est_legumes(c.name)):
                    _traiter_legumes(c, renommer=True)

    return 0


def _mode_local(fichier_path, date_cible, base):
    """Compte les paniers d'un tableur déjà téléchargé, sans connexion IMAP."""
    if fichier_path:
        chemin = Path(fichier_path).expanduser()
    else:
        chemin = _trouver_legumes_local(base, date_cible)
        if chemin is None:
            ou = (f"la semaine du {date_cible:%d/%m/%Y}" if date_cible
                  else f"sous {Path(base).expanduser()}")
            print(f"Aucun tableur Légumes trouvé localement ({ou}).")
            return 1
    if not chemin.exists():
        print(f"ERREUR : fichier introuvable : {chemin}")
        return 1
    print(f"Lecture du tableur Légumes : {chemin}")
    _traiter_legumes(chemin)
    return 0


def _traiter_legumes(chemin, renommer=False):
    """Lit le tableur Légumes, imprime les paniers ; renomme le fichier si demandé."""
    try:
        paniers = compter_paniers(chemin)
    except Exception as e:                                    # robustesse lecture xls
        print(f"  ERREUR lecture Légumes ({chemin.name}) : {e}")
        return
    if renommer:
        nouveau = renommer_avec_paniers(chemin, paniers)
        if nouveau.name != chemin.name:
            print(f"  renommé : {nouveau.name}")
    print("  Légumes — paniers à livrer :")
    print(f"    petit : {paniers['petit']}")
    print(f"    moyen : {paniers['moyen']}")
    print(f"    grand : {paniers['grand']}")


if __name__ == "__main__":
    sys.exit(main())
