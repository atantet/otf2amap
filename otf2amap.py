#!/usr/bin/env python3
"""
Lanceur de compatibilité : `python3 otf2amap.py …` reste valable.

Le code vit désormais dans le package src/otf2amap/. Après `pip install -e .`,
la commande `otf2amap …` fait la même chose. On insère src/ en tête de
sys.path pour que le package l'emporte sur ce fichier homonyme.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from otf2amap.cli import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())
