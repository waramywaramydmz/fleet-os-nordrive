# GUIDE DE MIGRATION VERS UNE NOUVELLE MACHINE 💻🚀

Ce guide vous explique exactement comment déplacer votre environnement **FLEET OS** sur un autre ordinateur (PC ou Mac) et faire en sorte qu'**Antigravity** reprenne le projet sans rien perdre de son intelligence ni de ses réglages.

---

## 📦 Contenu du dossier `sheets antigravity`

Ce dossier contient **l'intégralité du moteur**, prêt à l'emploi :

1. **`AGENTS.md` & `GEMINI.md`** : **Le cerveau de l'IA**. C'est le fichier le plus important. Dès qu'Antigravity s'ouvre dans ce dossier sur votre nouvelle machine, il lit ce fichier et sait immédiatement tout sur vos 72 chauffeurs, vos calculs, le placement de la ligne Total en bas, et vos règles de paie.
2. **`GoogleAppsScript_FLEET_OS.js`** : Le code à mettre dans **Extensions > Apps Script** sur Google Sheets.
3. **`FLEET_OS_Nordrive_Reconciliation_Qonto.xlsx`** : Le fichier Excel complet à importer dans Google Sheets.
4. **`calculer_semaine.cjs`** : Le script de calcul automatique.
5. **`CALCULER_SEMAINE.bat`** : Le lanceur rapide Windows.
6. **`IBANS_CHAUFFEURS.csv`** : Vos IBANs pour Qonto.
7. **`uber_current.csv` & `bolt_current.csv`** : Vos derniers relevés pour tests.
8. **`package.json`** : Les dépendances logicielles.

---

## 🛠️ Étape 1 : Le meilleur moyen de transférer ce dossier

Le moyen le plus simple et le plus rapide dépend de vos outils :

### Option A : Clé USB ou Disque externe (Le plus direct)
1. Branchez votre clé USB sur cette machine.
2. Copiez le dossier entier `sheets antigravity` dessus.
3. Branchez la clé sur votre nouvelle machine et collez le dossier où vous voulez (par exemple dans `Documents` ou sur le `Bureau`).

### Option B : Google Drive ou OneDrive (Sans clé USB)
1. Compressez le dossier `sheets antigravity` en fichier `.zip` (Clic droit ➔ *Envoyer vers* ➔ *Dossier compressé* ou *Compresser*).
2. Déposez ce `.zip` sur votre Google Drive ou OneDrive.
3. Sur votre nouvelle machine, téléchargez le `.zip` et décompressez-le.

### Option C : WeTransfer (Rapide et gratuit)
1. Allez sur [wetransfer.com](https://wetransfer.com).
2. Glissez le dossier zippé et envoyez-vous le lien par email.
3. Ouvrez le lien sur votre nouvelle machine.

---

## 🤖 Étape 2 : Lancer Antigravity sur la nouvelle machine

1. **Installez Antigravity** sur votre nouvelle machine (connectez-vous avec n'importe quel compte, peu importe qu'il soit différent).
2. Cliquez sur **File ➔ Open Folder...** (ou *Ouvrir le dossier*).
3. Sélectionnez le dossier **`sheets antigravity`** que vous venez de copier.
4. **C'est tout !**
   - Grâce au fichier **`AGENTS.md`** présent dans le dossier, le nouvel Antigravity va automatiquement ingérer tout le contexte du projet dès votre premier message.
   - Vous n'aurez pas besoin de lui réexpliquer la flotte, les 72 chauffeurs, ni que le Total doit rester tout en bas : **il le sait déjà**.

---

## 🚗 Étape 3 : Que dire à Antigravity sur la nouvelle machine ?

Pour démarrer votre session, il vous suffira simplement de lui dire par exemple :
> *"J'ai déposé mes 2 CSV Uber et Bolt de la semaine dans le dossier, calcule-moi la semaine et prépare le virement Qonto et le Google Sheets."*

Et Antigravity saura exactement quels scripts exécuter, quelles colonnes vérifier, et comment mettre à jour votre tableau avec précision chirurgicale.
