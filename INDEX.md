# Index du projet otf2amap

## Arborescence

| Chemin | Description |
|--------|-------------|
| `src/otf2amap/` | Package principal (voir détail ci-dessous) |
| `otf2amap.py` | Lanceur de compatibilité — `python3 otf2amap.py …` |
| `pyproject.toml` | Packaging, commande `otf2amap`, config ruff et pytest |
| `environment.yml` | Environnement conda (dépendances + outils de dev) |
| `config.toml` | Configuration de sortie (dossier, format, nom) |
| `README.md` | Documentation d'usage et d'installation |
| `INDEX.md` | Ce fichier |
| `tests/` | Batterie de tests (unitaires + intégration) |
| `notes/note_otf2amap.md` | Note de travail — algorithme, structure du PDF, typographie |
| `exemples/` | PDFs source OuvreTaFerme (non versionné, cf. `.gitignore`) |

## Package `src/otf2amap/`

| Module | Rôle |
|--------|------|
| `extract.py` | Lecture du PDF : date (page 2), quantités par panier (page 2), tableau des ventes (page 1) |
| `allocate.py` | Attribution des quantités aux paniers (page 2 ou permutation optimale) |
| `render.py` | Rendu PDF / PNG (A5 paysage) |
| `text.py` | Rendu Markdown / texte ASCII |
| `config.py` | Lecture de `config.toml` |
| `naming.py` | Nommage par défaut des fichiers de sortie (préfixe semaine ISO) |
| `core.py` | Orchestration : `build_sheet(pdf) → (rows, paniers, titre)` |
| `cli.py` | Analyse des arguments, orchestration, point d'entrée `main()` |

## Usage rapide

```bash
conda env create -f environment.yml && conda activate otf2amap
pip install -e .
otf2amap exemples/Ventes.pdf
```

Voir `README.md` pour le détail des options et de la configuration.
