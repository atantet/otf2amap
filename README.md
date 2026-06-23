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

## En amont : récupérer les tableurs par mail (`mail2amap`)

AmapJ envoie chaque semaine, sur la boîte Gmail de l'AMAP, un mail par contrat
(Légumes, Pain, …) avec le tableur de distribution en pièce jointe. La commande
`mail2amap` télécharge ces pièces jointes et affiche le nombre de paniers de
légumes par type — ce qui sert à créer la vente sur OuvreTaFerme, dont l'export
PDF est ensuite passé à `otf2amap`.

```bash
mail2amap                      # dernière livraison reçue (défaut)
mail2amap --tout               # toutes les livraisons trouvées dans la boîte
mail2amap --date 24/06/2026    # une livraison précise
mail2amap --fichier            # recompte un tableur déjà téléchargé, SANS mail
mail2amap --fichier chemin/legumes.xls   # … un fichier précis
mail2amap --fichier --date 24/06/2026    # … le tableur Légumes de cette semaine
```

À chaque téléchargement, le tableur Légumes est renommé avec le nombre de
paniers en suffixe (`…-24-06-2026_14_8_0.xls` = 14 petits, 8 moyens, 0 grands),
ce qui rend le compte visible directement dans le nom de fichier. L'option
`--fichier` relit un tableur déjà présent sans se connecter au mail : avec un
chemin explicite, ou (sans argument) le tableur Légumes de la semaine `--date`,
ou à défaut celui de la semaine la plus récente trouvée sous `dossier_base`.

Pour une même livraison, AmapJ envoie **un mail par contrat** (Légumes, Pain,
Pommes, Poires…) avec un objet et un expéditeur **identiques** : seul le corps
du mail (`Nom du contrat : …`) distingue les contrats. `mail2amap` lit donc le
corps pour identifier le contrat de chaque mail.

Les mails sont filtrés sur l'expéditeur **et** le motif d'objet ; la date de
livraison est lue dans l'objet (« … du mercredi 24 juin 2026 ») et les pièces
jointes sont enregistrées dans `<dossier_base>/<année>/S<semaine>/` (semaine ISO,
ex. `…/2026/S26/`). Le tableur du contrat **Légumes** est ensuite lu et le
cumul affiché :

```
Livraison du 24/06/2026 — Légumes (avril 2026 à fin septembre 2026) → …/2026/S26
  téléchargé : distri-Légumes (…)-24-06-2026.xls
  Légumes — paniers à livrer :
    petit : 12
    moyen : 7
    grand : 0

Livraison du 24/06/2026 — Pain (avril 2026 - septembre 2026) → …/2026/S26
  téléchargé : distri-Pain (…)-24-06-2026.xls
```

### Configuration

Partie non secrète dans `config.toml` :

```toml
[mail]
expediteur   = "saint-malo@m.amapj.fr"   # adresse émettrice à filtrer
objet_motif  = "Feuille de livraison"    # l'objet du mail doit le contenir
# dossier_base = "~/…/AMAP"              # défaut : dossier parent du dépôt
```

Identifiants IMAP **jamais versionnés** : ils vivent dans `.env` (gitignoré).
Copier `.env.example` en `.env` et renseigner un **mot de passe d'application**
Gmail (16 caractères, dédié et révocable — pas le mot de passe du compte) :

```bash
cp .env.example .env
# puis remplir AMAP_IMAP_USER / AMAP_IMAP_PASSWORD
```

Obtention du mot de passe d'application : activer la validation en 2 étapes du
compte Google, puis le générer sur <https://myaccount.google.com/apppasswords>.

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
  config.py     lecture de config.toml ([output] et [mail])
  naming.py     nommage par défaut (préfixe semaine ISO)
  cli.py        arguments et orchestration (otf2amap)
  mailbox.py    récupération IMAP, date (objet) et contrat (corps) des mails
  mailenv.py    identifiants IMAP depuis .env (jamais versionné)
  legumes.py    comptage des paniers dans le tableur .xls AmapJ
  mailcli.py    arguments et orchestration (mail2amap)
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
