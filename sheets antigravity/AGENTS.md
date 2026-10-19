# FLEET OS — NORDRIVE VTC (MÉMOIRE CENTRALE ANTIGRAVITY)

> **Ce fichier est la mémoire opérationnelle du projet.**  
> Tout agent Antigravity qui ouvre ce dossier doit lire et appliquer rigoureusement les règles ci-dessous sans jamais déroger à l'architecture établie.

---

## 1. Contexte du Projet & Objectif
- **Entreprise** : Flotte VTC Nordrive (Géré par Aimen & Rayen).
- **Objectif** : Concilier chaque semaine les relevés CSV des plateformes **Uber** et **Bolt**, calculer les gains réels de chaque chauffeur, déduire les forfaits contractuels et charges/malus, et générer le fichier de virement bancaire **SEPA pour Qonto**.
- **Outil Principal** : **Google Sheets** (avec interface Apps Script intégrée dans le menu `🚗 FLEET OS`) et moteur miroir local en Node.js (`calculer_semaine.cjs`).

---

## 2. Structure et Disposition Obligatoire du Dashboard (`DASHBOARD_SEMAINE`)

| Emplacement | Rôle | Règle Stricte |
| :--- | :--- | :--- |
| **Lignes 1 à 3** | Titres & En-têtes de bordure | Ne JAMAIS écraser ou altérer |
| **Lignes 4 à 5** | 4 Cartes KPI Consolidées | `B5` (Encaissé Plateformes), `F5` (Espèces totales), `J5` (Redevance flotte), `N5` (Net à virer Qonto) |
| **Lignes 8 à 9** | En-têtes de colonnes (Colonnes A à S) | Ne pas déplacer |
| **Lignes 10 à 77** | Chauffeurs N° 01 à 68 | Flotte historique |
| **Ligne 78** | **Rayen Klai** (N° 69) | Forfait : 0,00 € (Direction) |
| **Ligne 79** | **Yahia Zemani** (N° 70) | Forfait : 50,00 € |
| **Ligne 80** | **Hasan Emin Coskun** (N° 71) | Forfait : 40,00 € |
| **Ligne 81** | **Raphaël Ernest Amié Vimond** (N° 72) | Forfait : 50,00 € |
| **LIGNE 82 (ET SUIVANTES)** | **`TOTAL CONSOLIDÉ SEMAINE (FLOTTE)`** | **DOIT TOUJOURS ÊTRE STRICTEMENT TOUT EN BAS !** |

### ⚠️ RÈGLE FONDAMENTALE D'AJOUT DE CHAUFFEUR :
Lors de l'ajout d'un nouveau chauffeur (via formulaire ou script), le code doit **TOUJOURS** faire :
```javascript
sheetDash.insertRowBefore(totalRow);
```
**Interdiction formelle d'ajouter un chauffeur sous la ligne du total.** Le total doit être repoussé automatiquement vers le bas pour rester l'ultime ligne du tableau.

---

## 3. Grille Contractuelle des Forfaits Hebdomadaires
- **Règle générale standard** : 50,00 € / semaine (prélevé uniquement si le chauffeur a roulé dans la semaine : `M > 0` ou `Espèces > 0`).
- **Anciens chauffeurs à 70 €** : Tous passés officiellement à **50,00 €** (*Hamzi, Mohamed Safir, Ben Smida, Boufalgha, Chinoune, Fitouri, Belkhir*).
- **Tarif Spécial 40 €** : Ali Kara, Stephane Kouassi, Hasan Emin Coskun.
- **Tarif Spécial 60 €** : Bouchaib Belkharras, Brandon Louissaint, Farid Adkhis, Imane Sarraj, Mehmet Ergun, Mohamed Aloui, Omar Chemkhi, Raoul Mansfeld Bemma, Ratiba Beldjoudi/Guesri, Zahia Guesri.
- **Direction / Exempté (0 €)** : Rayen Klai.

---

## 4. Règles Mathématiques Absolues & Calibrage

### A. Calibrage étalon : Siré Cyril Niakaté
- Uber Réel : **1 325,32 €**
- Bolt Réel : **160,25 €**
- Total Réel Encaissé (Col M) : **1 485,57 €**
- Forfait Flotte (Col N) : **50,00 €**
- Net à virer Qonto (Col R) : **1 435,57 €**
- IBAN Qonto : `FR7630006000010000000000189`

### B. Déduction des Frais Uber
- Toujours déduire les frais de réservation Uber (ex : 0,12 € colonne L de `DATA_UBER`).
- Formule Virement Uber (Col Z) :
  `=(D + G + K + R) - ABS(J) - ABS(L)`

### C. Déduction des Commissions et Espèces Bolt
- Formule Virement Bolt (Col Z) :
  `=(E + F + H + I + J + K + L + N + V + W) - ABS(O + P + Q + S) - ABS(G)`

### D. Règle Syntaxique Google Apps Script
- Google Apps Script `.setFormula()` requiert **obligatoirement la syntaxe anglaise standard (EN-US)** avec des virgules `,` :
  - Utiliser `OR(...)` et jamais `OU(...)`.
  - Utiliser `AND(...)` et jamais `ET(...)`.
  - Utiliser `SUM(...)` et jamais `SOMME(...)`.
  - Pour les cellules fusionnées KPI (`B5:D5`), cibler uniquement la première cellule : `sheetDash.getRange("B5").setFormula(...)`.

---

## 5. Fichiers du Projet et Rôles

1. **`GoogleAppsScript_FLEET_OS.js`** :
   Script principal à coller dans **Extensions > Apps Script** sur Google Sheets. Ajoute le menu interactif `🚗 FLEET OS` avec les boutons 1-clic :
   - `✨ 1-CLIC : Aligner 72 Chauffeurs + Total TOUT EN BAS`
   - `📥 Importer Relevés Uber & Bolt (Mac & PC)`
   - `➕ Ajouter un nouveau chauffeur (Total reste en bas)`
   - `📥 Télécharger le CSV pour Qonto`
   - `📦 Archiver la semaine en cours`

2. **`FLEET_OS_Nordrive_Reconciliation_Qonto.xlsx`** :
   Classeur Excel officiel prêt à importer dans Google Sheets (Fichier > Importer > Téléverser > Remplacer la feuille de calcul).

3. **`calculer_semaine.cjs`** :
   Script Node.js autonome qui lit `uber_current.csv` et `bolt_current.csv` et recalcule instantanément toute la flotte.

4. **`IBANS_CHAUFFEURS.csv`** :
   Tableau des coordonnées bancaires des chauffeurs pour l'export SEPA.
