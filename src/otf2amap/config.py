"""Lecture de config.toml ([output], [mail], [otf])."""

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
    """Cherche config.toml dans le dossier courant puis ses parents. Retourne le premier Path trouvé, ou None."""
    start = Path(start or Path.cwd()).resolve()
    for directory in (start, *start.parents):
        candidate = directory / 'config.toml'
        if candidate.exists():
            return candidate
    return None


def _load_section(section, config_path=None):
    """Charge une section de config.toml. Le fichier doit exister ; la section est optionnelle."""
    config_path = Path(config_path) if config_path else find_config()
    if not config_path or not config_path.exists():
        print("ERREUR : config.toml introuvable (cherché depuis le dossier courant)")
        sys.exit(1)
    if tomllib is None:
        print("ERREUR : pip install tomli pour lire config.toml (Python < 3.11)")
        sys.exit(1)
    with open(config_path, 'rb') as f:
        data = tomllib.load(f)
    return data.get(section, {})


def load_config(config_path=None):
    """Charge la section [output] de config.toml (dossier, format, nom, drive_remote)."""
    return _load_section('output', config_path)


def load_mail_config(config_path=None):
    """Charge la section [mail] de config.toml (expediteur, objet_motif, dossier_base)."""
    return _load_section('mail', config_path)


def load_otf_config(config_path=None):
    """Charge la section [otf] de config.toml (dossier_base, nom_pdf)."""
    return _load_section('otf', config_path)
