# otf2amap

Convertit un PDF de synthèse de vente [OuvreTaFerme](https://ouvretaferme.org) en feuille de préparation des paniers AMAP, au format A5 paysage, prête à imprimer en noir et blanc.

## Usage

```bash
python3 otf2amap.py Ventes.pdf [sortie.pdf] [--montant] [--scale 1.0]
```

- `Ventes.pdf` : PDF exporté depuis OuvreTaFerme (2 pages)
- `sortie.pdf` : optionnel, par défaut `feuille_paniers_amap.pdf` dans le même dossier
- `--montant` : ajoute une colonne MONTANT
- `--scale` : ajuste la taille globale du texte (ex. `--scale 0.8` pour 8+ produits, `--scale 1.2` pour 5 produits ou moins)

## Ce que ça produit

Un tableau A5 paysage avec :

- La **date de retrait** en en-tête de la colonne produit (extraite de la page 2 du PDF)
- Une colonne **TOTAL** (quantité totale commandée)
- Une colonne par type de panier : **N PETIT**, **N MOYEN**, **N GRAND** (selon ce qui est commandé)
- Les quantités par panier calculées automatiquement pour chaque produit
- Les lignes "Panier de la semaine" filtrées (elles servent uniquement à détecter les types de paniers)

## Dépendances

```bash
pip install pdfplumber pypdf reportlab
```

## Fonctionnement interne

Le PDF source comporte deux pages :

- **Page 1** : tableau des ventes (PRODUIT / QUANTITÉ / MONTANT / COMMANDES)
- **Page 2** : récapitulatif des commandes par panier, avec la date de retrait

Le script gère deux mises en page possibles du PDF source (colonnes séparées ou fusionnées).

L'attribution des quantités aux paniers utilise une **permutation optimale** : pour chaque produit, le script teste toutes les façons d'associer les tokens `1 x N unité` aux paniers, et retient celle dont les ratios `qté / n_paniers` sont les plus ronds (score lexicographique sur l'erreur d'arrondi à 2, 1 et 0 décimales).
