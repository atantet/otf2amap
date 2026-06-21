"""Chargement des secrets IMAP depuis l'environnement (.env ou os.environ).

Même principe que le projet meteo : aucun secret n'est versionné. Les
identifiants viennent de variables d'environnement, alimentées en local par
un fichier `.env` (gitignoré ; voir `.env.example`).
"""

import os
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:               # python-dotenv absent
    load_dotenv = None

# Racine du dépôt = parent de src/ (mailenv.py est dans src/otf2amap/).
REPO_ROOT = Path(__file__).resolve().parents[2]


class ConfigError(Exception):
    """Identifiant manquant ou configuration invalide."""


def load_dotenv_if_present(dotenv_path=None):
    """Charge `.env` du dépôt dans os.environ, idempotent.

    No-op si python-dotenv est absent ou si le fichier n'existe pas (les
    variables peuvent déjà être présentes dans l'environnement).
    """
    if load_dotenv is None:
        return
    p = Path(dotenv_path) if dotenv_path else REPO_ROOT / ".env"
    if p.exists():
        load_dotenv(p)


def get_secret(name, default=None):
    """Lit une variable d'environnement, lève ConfigError si requise et absente."""
    value = os.environ.get(name, default)
    if value is None or value == "":
        raise ConfigError(
            f"Variable d'environnement requise non définie : {name}. "
            f"Voir .env.example pour le setup."
        )
    return value


def load_imap_secrets():
    """Charge les identifiants IMAP depuis l'environnement.

    Retourne un dict {host, user, password}. Lève ConfigError si l'un manque.
    """
    load_dotenv_if_present()
    return {
        "host": get_secret("AMAP_IMAP_HOST", "imap.gmail.com"),
        "user": get_secret("AMAP_IMAP_USER"),
        "password": get_secret("AMAP_IMAP_PASSWORD"),
    }
