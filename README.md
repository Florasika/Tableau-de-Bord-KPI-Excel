# 📊 Jour 1 / 10 — Tableau de Bord KPI Excel

> **Série : 10 Days of Excel** · Jour 1/10  
> Niveau : Débutant · Durée estimée : 45 min  
> Concepts : SOMME.SI · NB.SI · MOYENNE.SI · Mise en forme conditionnelle · Graphiques

---

## 📁 Structure du fichier

```
dashboard_kpi_jour1.xlsx
│
├── DONNÉES     ← 48 lignes de ventes brutes (ne jamais modifier ici)
└── DASHBOARD   ← KPIs + tableaux croisés + graphiques + mise en forme
```

---

## 🛠️ Guide étape par étape

### ÉTAPE 1 — Préparer la feuille DONNÉES

1. Ouvre le fichier et va sur la feuille **DONNÉES**
2. Vérifie que les 8 colonnes sont bien présentes :
   `Date · Produit · Région · Catégorie · Vendeur · Quantité · Prix Unitaire · Montant`
3. La colonne **Montant** doit contenir la formule `=F2*G2` (Quantité × Prix Unitaire)
   — tire-la sur toutes les lignes si ce n'est pas fait
4. Sélectionne toute la plage `A1:H49`
5. **Insertion → Tableau** → coche "Mon tableau comporte des en-têtes" → OK
   _(Nommer le tableau "TblVentes" dans le champ en haut à gauche)_
6. **Affichage → Figer les volets → Figer la ligne supérieure**

---

### ÉTAPE 2 — Créer les KPI globaux sur le DASHBOARD

Va sur la feuille **DASHBOARD** et crée les 5 KPI suivants dans des cellules séparées.
Mets un label au-dessus et la formule en dessous :

| KPI | Formule | Format |
|-----|---------|--------|
| CA Total | `=SOMME(DONNÉES!H2:H49)` | `#,##0 "€"` |
| Nb Ventes | `=NB(DONNÉES!H2:H49)` | `#,##0` |
| Panier Moyen | `=MOYENNE(DONNÉES!H2:H49)` | `#,##0 "€"` |
| Qté Totale | `=SOMME(DONNÉES!F2:F49)` | `#,##0` |
| Nb Produits | `=NB.SI.ENS(DONNÉES!B2:B49,DONNÉES!B2:B49)` | `#,##0` |

**Mise en forme des KPI :**
- Fond sombre sur les labels (couleur personnalisée `#1E3A5F`)
- Valeurs en grand (taille 18-22) avec couleur vive par KPI
- Bordure colorée en haut de chaque carte KPI

---

### ÉTAPE 3 — Tableau SOMME.SI par Produit

Crée un tableau à 3 colonnes : **Produit · CA (€) · Nb Ventes**

Saisis les 5 produits manuellement en colonne B :
```
Laptop Pro
Smartphone X
Tablette Air
Écouteurs BT
Montre Smart
```

En colonne C (CA), utilise **SOMME.SI** :
```excel
=SOMME.SI(DONNÉES!B$2:B$49, B10, DONNÉES!H$2:H$49)
```
- `DONNÉES!B$2:B$49` → plage de recherche (colonne Produit)
- `B10` → le produit à chercher (cellule courante)
- `DONNÉES!H$2:H$49` → plage à additionner (colonne Montant)
- Le **$** fixe la plage pour pouvoir tirer la formule vers le bas

En colonne D (Nb Ventes), utilise **NB.SI** :
```excel
=NB.SI(DONNÉES!B$2:B$49, B10)
```

Ajoute une ligne **TOTAL** avec `=SOMME(C10:C14)` et `=SOMME(D10:D14)`

---

### ÉTAPE 4 — Tableau SOMME.SI par Vendeur

Même principe que l'étape 3, mais sur la colonne **Vendeur** (colonne E des données).

Saisis les 5 vendeurs :
```
Alice M. · Karim B. · Lucie D. · Thomas R. · Nadia K.
```

Formule CA par vendeur :
```excel
=SOMME.SI(DONNÉES!E$2:E$49, F10, DONNÉES!H$2:H$49)
```

---

### ÉTAPE 5 — Mise en forme conditionnelle

**Sur la colonne CA Produit :**
1. Sélectionne les cellules de CA (ex. `C10:C14`)
2. **Accueil → Mise en forme conditionnelle → Barres de données**
3. Choisis une couleur vive (vert ou bleu)
→ Chaque cellule affiche une mini barre proportionnelle à la valeur

**Sur la colonne CA Vendeur :**
1. Sélectionne les cellules de CA vendeur
2. **Accueil → Mise en forme conditionnelle → Nuances de couleurs**
3. Choisis le dégradé vert-blanc-rouge (du plus haut au plus bas)
→ Le meilleur vendeur est en vert, le moins bon en rouge

---

### ÉTAPE 6 — Créer les graphiques

**Graphique 1 — Bar chart CA par Produit :**
1. Sélectionne les colonnes Produit + CA (y compris en-têtes)
2. **Insertion → Graphique à barres → Barres groupées**
3. Clic droit → **Mettre en forme la série de données**
4. Change la couleur des barres → couleur vive (vert `#00C96E`)
5. Ajoute les étiquettes : **Disposition du graphique → Étiquettes de données**
6. Supprime le quadrillage inutile
7. Change le fond du graphique : **Zone de traçage → Remplissage → couleur sombre**

**Graphique 2 — Camembert CA par Vendeur :**
1. Sélectionne Vendeur + CA vendeur
2. **Insertion → Graphique → Secteurs (camembert)**
3. Style de graphique → choisir un style avec % affichés
4. Clic droit sur les secteurs → **Mettre en forme** → applique des couleurs distinctes

---

### ÉTAPE 7 — Finitions et mise en page

1. **Nommer la feuille** DASHBOARD avec une couleur d'onglet (bleu)
2. **Masquer les quadrillages** : Affichage → décocher Quadrillage
3. **Titre du dashboard** : fusionne plusieurs cellules en haut, écris le titre,
   mets un fond sombre et le texte en blanc gras taille 16
4. **Zones de couleur** : applique des fonds différents pour séparer visuellement
   les sections (KPIs / Produits / Vendeurs / Graphiques)
5. **Protéger la feuille DONNÉES** :
   Révision → Protéger la feuille → laisse uniquement "Sélectionner les cellules"
   _(pour éviter les modifications accidentelles)_

---

## 🔑 Formules utilisées

### SOMME.SI
```excel
=SOMME.SI(plage_critère, critère, plage_somme)
```
Additionne les cellules de `plage_somme` où `plage_critère` = `critère`.

### NB.SI
```excel
=NB.SI(plage, critère)
```
Compte le nombre de cellules de `plage` qui correspondent à `critère`.

### MOYENNE.SI
```excel
=MOYENNE.SI(plage_critère, critère, plage_moyenne)
```
Calcule la moyenne des cellules qui correspondent au critère.

### Référence absolue ($)
```excel
=SOMME.SI(DONNÉES!B$2:B$49, B10, DONNÉES!H$2:H$49)
```
Le `$` avant le numéro de ligne fixe la plage — indispensable pour tirer la formule
vers le bas sans que la plage de référence se décale.

---

## 💡 Bonnes pratiques

- **Ne jamais analyser directement dans la feuille DONNÉES** — toujours dans DASHBOARD
- **Toujours utiliser `$`** sur les plages de données dans SOMME.SI et NB.SI
- **Nommer les tableaux** (Tableau1 → TblVentes) pour des formules plus lisibles :
  `=SOMME.SI(TblVentes[Produit], B10, TblVentes[Montant])`
- **Actualisation automatique** : les formules se recalculent dès que tu ajoutes
  une ligne dans DONNÉES — le dashboard est toujours à jour

---

## 📸 Aperçu du dashboard final

```
┌─────────────────────────────────────────────────┐
│  📊 TABLEAU DE BORD KPI — VENTES 2024           │
├──────────┬──────────┬──────────┬────────────────┤
│ CA TOTAL │ NB VENTES│ PANIER   │ QTÉ TOTALE    │
│ 289,450€ │    48    │  6,030€  │     534        │
├──────────┴──────────┴──────────┴────────────────┤
│ CA PAR PRODUIT (SOMME.SI)   │ GRAPH BAR        │
│ Laptop Pro    ████ 85,200€  │                  │
│ Smartphone X  ███  72,800€  │     [Bar Chart]  │
│ Tablette Air  ██   58,400€  │                  │
├─────────────────────────────┤                  │
│ CA PAR VENDEUR              │ [Camembert]      │
│ Alice    ██████ 68,000€     │                  │
│ Karim    █████  55,000€     │                  │
└─────────────────────────────┴──────────────────┘
```
<img width="853" height="746" alt="image" src="https://github.com/user-attachments/assets/36e79618-4ec9-427b-a7f1-70b82daa45fc" />


- [Aide officielle SOMME.SI](https://support.microsoft.com/fr-fr/office/somme-si)
- [Aide officielle NB.SI](https://support.microsoft.com/fr-fr/office/nb-si)
- [Mise en forme conditionnelle](https://support.microsoft.com/fr-fr/office/mise-en-forme-conditionnelle)
