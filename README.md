# CaneFlow

CaneFlow est une application Electron/React pour convertir un carnet de cables Caneco en fichier Excel compatible Multidoc.

Version actuelle : v0.2.0

## Demarrer

```bash
npm install
npm run electron:dev
```

## Build Windows

```bash
npm run build:electron
```

## Conversion en ligne de commande

```bash
node scripts/convert.js --input "CARNET DE CABLES TGBT.xls" --prix 0 --tva 0
```

## Notes

- L'icone Windows doit se trouver dans `electron/caneflow.ico`.
- Les fichiers .xls sont convertis via Microsoft Excel (installe sur la machine).
- Les mises a jour sont publiees via GitHub Releases.
