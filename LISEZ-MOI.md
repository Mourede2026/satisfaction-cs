# Déploiement
1. **Apps Script** : remplacez tout le contenu de `Code.gs` par le fichier fourni, enregistrez, déployez une **nouvelle version**, exécutez `setupComptesEtQrTokens` puis `effacerMotsDePasse`.
2. **GitHub** : déposez à la racine du dépôt tous les fichiers de ce dossier sauf `Code.gs` et `LISEZ-MOI.md` (11 fichiers : 6 pages HTML + `config.js`, `sw.js`, `manifest.json`, 2 icônes). Les noms respectent la casse exacte.
3. **GitHub Pages** : Settings > Pages > branche `main`, dossier `/`.
4. Ouvrez `AdminQR.html`, connectez-vous en `admin`, imprimez le QR enquêteurs. Les enquêteurs le scannent une fois.
5. Ne publiez jamais `Code.gs` dans le dépôt public.
