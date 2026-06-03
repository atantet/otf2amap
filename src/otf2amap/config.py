"""Lecture de config.toml ([output] dossier / format / nom)."""

import sys
from pathlib import Path

try:
    import tomllib
except ImportError:               # Python < 3.11
    try:
        import tomli as tomllib
    except ImportError:
        tomllib = None


def find_config(start=None):
    """
    Cherche config.toml dans le dossier courant puis ses parents.
    Retourne le premier Path trouvé, ou None.
    """
    start = Path(start or Path.cwd()).resolve()
    for directory in (start, *start.parents):
        candidate = directory / 'config.toml'
        if candidate.exists():
            return candidate
    return None


def load_config(config_path=None):
    """
    Charge la section [output] de config.toml.

    Si `config_path` n'est pas fourni, cherche config.toml depuis le dossier
    courant. Erreur si aucun fichier n'est trouvé (toutes les clés restent
    optionnelles, mais le fichier doit exister).

    Retourne un dict avec les clés présentes parmi : dossier, format, nom.
    """
    config_path = Path(config_path) if config_path else find_config()
    if not config_path or not config_path.exists():
        print("ERREUR : config.toml introuvable (cherché depuis le dossier courant)")
        sys.exit(1)
    if tomllib is None:
        print("ERREUR : pip install tomli pour lire config.toml (Python < 3.11)")
        sys.exit(1)
    with open(config_path, 'rb') as f:
        data = tomllib.load(f)
    return data.get('output', {})
