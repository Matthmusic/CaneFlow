# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commandes de développement

```bash
# Lancer le serveur Vite seul (UI uniquement, sans Electron)
npm run dev

# Lancer l'app complète en mode dev (Vite + Electron simultanément)
npm run electron:dev

# Build de production (TypeScript + Vite)
npm run build

# Build installateur Windows (.exe via electron-builder)
npm run build:electron

# Linter
npm run lint
```

En mode dev, le DevTools Electron s'ouvre avec `Ctrl+Shift+I` ou `F12`.

## Architecture

L'app est une application **Electron + React** (frameless window) qui convertit des carnets de câbles Caneco (Excel) vers le format Multidoc (Excel).

### Couches

```
src/App.tsx              — UI React : état, 3 étapes, IPC vers window.api
src/global.d.ts          — Types TypeScript de window.api exposé par preload

electron/preload.cjs     — Pont contextIsolation : expose window.api via contextBridge
electron/main.cjs        — Process principal : BrowserWindow, ipcMain handlers, autoUpdater
electron/excel.cjs       — Lecture/écriture Excel (xlsx + ExcelJS) + conversion .xls via PowerShell/COM
electron/transform.cjs   — Logique métier : parsing des lignes Caneco, génération du titre, mapping Multidoc
electron/nexus.cjs       — Lookup des prix câbles depuis NEXUS (R2 → Z: → config locale)
electron/cables.json     — Coûts de pose par section (pt_pose_value), embarqué dans l'ASAR
```

### Flux de données

1. L'utilisateur choisit un fichier Excel Caneco (`.xls` ou `.xlsx`)
2. `window.api.previewRows` → IPC → `electron/main.cjs` → `readSheetRows` (excel.cjs) → `buildPreviewRows` (transform.cjs) → affichage dans la table de prix
3. L'utilisateur saisit les prix (par ligne ou par type de câble) et la TVA
4. `window.api.convertFile` → IPC → `mapSheetRows` (transform.cjs) → `writeSheetRows` (excel.cjs) → fichier Multidoc

### Conversion .xls

Les fichiers `.xls` sont d'abord tentés via la lib `xlsx`. En cas d'échec, fallback vers un script PowerShell qui utilise l'API COM de Microsoft Excel pour convertir en `.xlsx` temporaire, lu ensuite par ExcelJS.

### Colonnes Caneco attendues (transform.cjs)

L'ordre par défaut est : Amont (0), Repere (1), Longueur (2), Cable (3), Neutre (4), PE (5), TypeCable (6), Nature (7). La détection d'en-tête est automatique par `hasHeaderRow` + `buildColumnIndex` avec normalisation NFD. Aliases reconnus pour Nature : `nature`, `matiere`, `mat.`, `conducteur`.

### Output Multidoc

8 colonnes : `N°`, `Titre`, `Unité`, `Quantité`, `Prix unitaire`, `` (vide), `TVA`, `Descriptif`. Unité fixe : `ml`.

### IPC channels

| Channel | Sens | Description |
|---|---|---|
| `pick-input-file` | renderer → main | Dialog d'ouverture Excel |
| `pick-output-file` | renderer → main | Dialog de sauvegarde |
| `preview-rows` | renderer → main | Lecture + preview des lignes Caneco |
| `convert-file` | renderer → main | Conversion complète vers Multidoc |
| `reveal-path` | renderer → main | Ouvre le dossier dans l'explorateur |
| `lookup-cable-prices` | renderer → main | Lookup des prix câbles depuis NEXUS (R2 → Z: → config locale) |
| `check-updates` / `download-update` / `install-update` | renderer → main | Auto-update via electron-updater |
| `window-close` / `window-minimize` / `window-toggle-maximize` | renderer → main | Contrôles de fenêtre (frame: false) |
| `update-event` | main → renderer | Événements de mise à jour |
| `app-log` | main → renderer | Logs redirigés vers la console du renderer |

### Intégration NEXUS (nexus.cjs)

Lookup automatique des prix câbles au chargement d'un fichier Caneco.

**Priorité de source** : Cloudflare R2 (5 s timeout) → `Z:\F - UTILITAIRES\NEXUS\prices.json` → `%APPDATA%\nexus\config.json` → `dataDir\prices.json`

**Clé composite** : `buildCableKey(row)` dans `transform.cjs` produit `"cable | typeCable"` (séparateur `CABLE_KEY_SEP = ' | '`). Ces clés sont passées comme `compositeKeys` à `lookupCablePrices`, puis re-splitées dans `parseCableKey`.

**Formule prix** : `materialPrice × margin + pt_pose_value` (marge par défaut 1.33, pose depuis `cables.json`).

**Match** : lookup exact sur `norm(typeCable + ' ' + cable)` dans l'exactMap, puis scan linéaire sur `normNames` (noms pré-normalisés). Les items sont filtrés par `CABLE_CATEGORIES` (13 catégories câbles — ne pas en ajouter sans vérifier les catégories dans `prices.json`).

**Cache** : TTL 5 min par variable de module (`_nexusItemsCache`, `_nexusExactMap`, `_nexusNormNames`). Le fallback Z:/local est aussi tenté si R2 renvoie 0 article (pas seulement en cas d'erreur réseau).

**ASAR** : `nexus.cjs` est autonome — ne pas le faire `require()` depuis `transform.cjs` ni l'inverse, sous peine d'échec silencieux dans l'app packagée.

### Releases

Les releases sont publiées sur GitHub via `electron-builder` avec `provider: github`. `electron-updater` vérifie automatiquement les mises à jour au démarrage (désactivé en mode dev). La version dans `package.json` est injectée dans le renderer via `import.meta.env.VITE_APP_VERSION`.
