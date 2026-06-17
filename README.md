# otf2amap

Convertit un PDF de synthèse de vente [OuvreTaFerme](https://ouvretaferme.org) en feuille de préparation des paniers AMAP, au format A5 paysage, prête à imprimer en noir et blanc.

## Installation

Avec conda (recrée l'environnement complet, poppler inclus) :

```bash
conda env create -f environment.yml
conda activate otf2amap
pip install -e .            # installe la commande `otf2amap`
```

Ou avec pip seul (poppler doit être installé par ailleurs pour la sortie PNG) :

```bash
pip install -e .
```

## Usage

```bash
otf2amap Ventes.pdf [sortie] [--montant] [--scale 1.0] [--format=png]
```

Équivalents :

```bash
python3 otf2amap.py Ventes.pdf      # lanceur de compatibilité (sans installation)
python3 -m otf2amap Ventes.pdf
```

- `Ventes.pdf` : PDF exporté depuis OuvreTaFerme (2 pages)
- `sortie` : optionnel ; par défaut `YYYY_SNN_feuille_paniers_amap.<format>`
- `--montant` : ajoute une colonne MONTANT
- `--scale` : ajuste la taille globale du texte (ex. `--scale 0.8` pour 8+ produits, `--scale 1.2` pour 5 produits ou moins)
- `--format` : `png` (défaut), `pdf`, `md`, `txt` — l'extension du fichier de sortie prime

## Configuration (`config.toml`)

Cherché depuis le dossier courant. Toutes les clés sont optionnelles, mais le fichier doit exister.

```toml
[output]
dossier = "~/chemin/vers/feuilles/"   # défaut : même dossier que le PDF source
format  = "png"                       # png (défaut), pdf, md, txt
# nom   = "mon_fichier"               # sans extension ; défaut : YYYY_SNN_feuille_paniers_amap
# drive_remote = "gdrive:La Petite Claye/AMAP/feuilles"   # envoi auto sur Google Drive (rclone)
```

Les options ligne de commande prévalent sur `config.toml`.

## Envoi automatique sur Google Drive (`rclone`)

Le fichier est **toujours** produit en local dans `dossier`. Si `drive_remote`
est renseigné, une **copie** est en plus envoyée sur Drive (le local reste).

Google Drive Desktop n'existant pas sous Linux, on passe par
[rclone](https://rclone.org). Configuration en une fois :

```bash
conda activate otf2amap          # rclone est installé dans l'environnement
rclone config                    # créer un remote « gdrive » de type Google Drive
                                 #  (n > drive > laisser la plupart par défaut,
                                 #   autoriser dans le navigateur, y pour confirmer)
rclone lsd gdrive:              # vérifier que la connexion marche
```

Puis dans `config.toml`, indiquer le dossier cible (le chemin tel qu'il apparaît
dans « Mon Drive », sans le préfixe « Mon Drive ») :

```toml
drive_remote = "gdrive:La Petite Claye/AMAP/feuilles"
```

À chaque conversion, la feuille est envoyée automatiquement. Si rclone est
absent ou l'envoi échoue, l'outil prévient sans planter et conserve le fichier
local.

## Ce que ça produit

Un tableau A5 paysage avec :

- La **date de retrait** en en-tête de la colonne produit (extraite de la page 2 du PDF)
- Une colonne **TOTAL** (quantité totale commandée)
- Une colonne par type de panier : **N PETIT**, **N MOYEN**, **N GRAND** (selon ce qui est commandé)
- Les quantités par panier calculées automatiquement pour chaque produit
- Les lignes « Panier de la semaine » filtrées (elles servent uniquement à détecter les types de paniers)

## Structure du code

```
src/otf2amap/
  extract.py    lecture du PDF (date, page 2, tableau page 1)
  allocate.py   attribution des quantités aux paniers
  render.py     rendu PDF / PNG
  text.py       rendu Markdown / txt
  config.py     lecture de config.toml
  naming.py     nommage par défaut (préfixe semaine ISO)
  cli.py        arguments et orchestration
otf2amap.py     lanceur de compatibilité
tests/          batterie de tests (unitaires + intégration)
notes/          notes de travail
```

L'attribution des quantités préfère la lecture exacte de la page 2 ; à défaut, elle répartit les tokens `1 x N unité` de la page 1 par **permutation optimale** (score lexicographique sur l'erreur d'arrondi des ratios `qté / n_paniers` à 2, 1 puis 0 décimales).

## Développement

```bash
pytest          # batterie de tests
ruff check .    # lint
```

Les tests d'intégration s'appuient sur les PDFs du dossier `exemples/` (non versionné) et sont ignorés s'il est absent.
