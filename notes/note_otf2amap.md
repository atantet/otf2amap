# Note de travail — Script `otf2amap.py`

## Contexte

Script Python développé pour l'AMAP du Bocage. Il transforme un PDF de ventes exporté depuis la plateforme OuvreTaFerme (OTF) en une feuille de préparation des paniers, imprimable en A5 paysage, noir et blanc.

Le PDF source comporte deux pages :
- **Page 1** : tableau des ventes (PRODUIT / QUANTITÉ / MONTANT / COMMANDES)
- **Page 2** : récapitulatif des commandes par panier, avec la date de retrait

Voir `INDEX.md` pour la liste complète des fichiers du projet.

---

## Ce que fait le script

1. **Extraction de la date de retrait** depuis la page 2 (regex `DD/MM/YYYY`)
2. **Extraction du tableau page 1** via `pdfplumber`, gestion de deux layouts PDF possibles (colonnes séparées ou fusionnées)
3. **Identification des paniers** (`Panier de la semaine - petit/moyen/grand`) et de leur nombre d'unités
4. **Attribution des quantités par panier** pour chaque produit, à partir des tokens `1 x N.NN unité` de la colonne COMMANDES
5. **Génération d'un PDF A5 paysage** avec `reportlab`

---

## Algorithme d'attribution token → panier

C'est le cœur du problème et le plus délicat. Chaque produit a autant de tokens que de types de paniers. Il faut associer chaque token au bon panier.

### Approche retenue : permutation optimale par score lexicographique

Pour chaque permutation possible (token → panier), on calcule un score en trois niveaux de précision décimale :

```python
score = tuple(
    sum(abs(r - round(r, dp)) for r in ratios)
    for dp in [2, 1, 0]
)
```

La permutation avec le score lexicographiquement minimal est retenue. Cela favorise les ratios `qté_token / n_panier` qui sont des nombres ronds à 2 décimales (ex. 0.4, 0.65) plutôt que des approximations (ex. 0.2154).

### Pourquoi ce score et pas l'erreur d'arrondi entière simple ?

L'erreur entière seule échoue sur des cas comme Betterave (tokens 2.8 et 2.6, paniers n=7 et n=13) : les deux permutations ont des erreurs presque identiques (0.587 vs 0.600) et l'algorithme choisit la mauvaise.

Le score à 2 décimales discrimine correctement : 2.8/7=0.4 (erreur 0 à 2dp) bat 2.8/13=0.215 (erreur 0.035 à 2dp).

### Cas à 1 token (produit présent dans un seul panier)

Géré naturellement par la permutation optimale : le seul token est associé au panier dont n le divise exactement, l'autre case reste vide.

---

## Structure du PDF généré
DD/MM/YYYY    │  TOTAL  │   N PETIT    │   N MOYEN
───────────────┼─────────┼──────────────┼───────────
Nom produit   │  10 kg  │   0.4 kg     │   0.7 kg
(sur 2 lignes │         │              │
si besoin)   │         │              │
- Format **A5 paysage** (595 × 420 pt)
- **Noir et blanc** — pas de fond coloré
- La date de retrait remplace l'en-tête "PRODUIT"
- Pas de cadre extérieur, pas de cadre autour de l'en-tête — uniquement les séparateurs de colonnes et de lignes
- Noms de produits découpés mot à mot sur plusieurs lignes si nécessaire
- Colonnes paniers symétriques, leur nombre s'adapte (2 ou 3 paniers)

---

## Typographie

Une seule taille de police (`S = 16 pt × scale`) pour tout le contenu — date, en-têtes de colonnes, noms de produits, chiffres. Seules les **unités** (kg, bte, u.) sont en `S × 0.65` (~10 pt), en police normale à côté des chiffres en gras.

Format affiché : **`10`** `kg` — identique dans la colonne TOTAL et dans les colonnes paniers.

---

## Largeurs de colonnes (valeurs par défaut)

| Colonne | Largeur |
|---------|---------|
| Date / PRODUIT | 180 pt |
| TOTAL | 64 pt |
| PETIT / MOYEN / GRAND | `(575 - 180 - 64) / n_paniers` pt |

Avec 2 paniers : ~165 pt chacun. Avec 3 paniers : ~110 pt chacun.

---

## Options

```bash
python3 otf2amap.py Ventes.pdf                  # défaut
python3 otf2amap.py Ventes.pdf sortie.pdf       # fichier de sortie explicite
python3 otf2amap.py Ventes.pdf --scale 0.8      # texte plus petit (8+ produits)
python3 otf2amap.py Ventes.pdf --scale 1.2      # texte plus grand (5 produits ou moins)
python3 otf2amap.py Ventes.pdf --montant        # ajoute une colonne MONTANT
```

Les largeurs de colonnes et marges sont indépendantes de `scale`.

---

## Dépendances Python


pdfplumber   # extraction du texte PDF
pypdf        # lecture/écriture PDF
reportlab    # génération du PDF de sortie
