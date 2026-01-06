<div align="center">
  <img src="public/logo.svg" alt="CaneFlow Logo" width="120" height="120">
  <h1>CaneFlow</h1>
  <p>Convertis ton carnet de câbles Caneco en fichier Excel compatible Multidoc, en un clic.</p>
</div>

Version actuelle : v0.2.0

## 📥 Installation

1. Va sur la page des [Releases](https://github.com/Matthmusic/CaneFlow/releases)
2. Télécharge le fichier `.exe` du dernier tag (ex: `CaneFlow-Setup-0.2.0.exe`)
3. Lance l'installateur
4. C'est tout ! L'application vérifie automatiquement les mises à jour.

## 🚀 Fonctionnement de l'app

CaneFlow simplifie la conversion de carnets de câbles Caneco vers le format Multidoc en 3 étapes :

### Étape 1 : Choisir le fichier source
- Importe ton fichier Excel Caneco (.xls ou .xlsx)
- Tu peux cliquer sur "Choisir un Excel" ou glisser-déposer le fichier
- Les fichiers .xls sont automatiquement convertis via Microsoft Excel (doit être installé)
- L'app charge automatiquement un aperçu des lignes détectées

### Étape 2 : Définir les prix (optionnel)
Deux modes de tarification au choix :
- **Par ligne** : Définis un prix unitaire pour chaque ligne du carnet
- **Par câble + type** : Définis un prix par catégorie de câble (colonne "Type de câble" dans Caneco)

Tu peux aussi :
- Définir un prix par défaut et l'appliquer à toutes les lignes/catégories
- Ajouter un taux de TVA global

### Étape 3 : Exporter vers Multidoc
- Clique sur "Convertir" pour générer l'Excel Multidoc
- Le fichier est enregistré avec les colonnes attendues par Multidoc
- Clique sur "Ouvrir le dossier" pour accéder directement au fichier généré
- Utilise "Nouvel export" pour recommencer une nouvelle conversion

### Colonnes Caneco requises
Ton export Caneco doit contenir ces colonnes dans cet ordre :
1. Amont
2. Descriptif
3. Longueur
4. Câble
5. Neutre
6. PE ou PEN
7. Type de câble

### Configuration Multidoc
Dans Multidoc, configure les numéros de colonnes comme suit :
- Numéros : 1
- Titres : 2
- Unités : 3
- Quantités : 4
- Prix unitaires : 5
- Colonne vide : 6
- TVA : 7
- Descriptif : 8
