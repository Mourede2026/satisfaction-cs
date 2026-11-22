# Déploiement
1. **Apps Script** : remplacez tout le contenu de `Code.gs`, enregistrez, déployez une **nouvelle version**, puis exécutez `setupComptesEtQrTokens` (crée la feuille `CodesAcces`) et `setupFeuillesHopitaux` (crée les 3 feuilles hôpital). Si une feuille `ACTEURS_COM_HOPITAL` existe, supprimez-la : elle n'est plus utilisée. Ne rééxécutez pas `effacerMotsDePasse` si déjà fait.
2. **GitHub** : déposez à la racine tous les fichiers sauf `Code.gs` et `LISEZ-MOI.md`, puis attendez 1-2 min (le cache `sat-v2` se met à jour tout seul).
3. **Utilisation** : l'accueil propose *Collecter des données* (Zone sanitaire / Hôpital) et *Administration*.
4. **Codes** : `AdminQR.html` (connexion admin, depuis téléphone ou tablette) permet de créer un code général ou un code par zone sanitaire / par hôpital, de le partager, imprimer, changer, désactiver ou supprimer.
5. Le personnel (CS et hôpital) n'a jamais besoin de code. Un enquêteur saisit son code une seule fois par niveau.
6. Ne publiez jamais `Code.gs` dans le dépôt public.

## Créer ou réinitialiser un compte directement dans Google Sheets
Feuille `Comptes` : ajoutez une ligne (ou modifiez la ligne existante) :
- A `Identifiant` = admin · C `Niveau` = admin · I `Actif` = OUI · **K `MotDePasseInitial` = le mot de passe choisi**
- Laissez B (`Hash`) vide ou inchangé.
À la première connexion, l'application vérifie ce mot de passe, le transforme en empreinte sécurisée et **efface la colonne K**.

## Formulaires hôpital (3)
`FormUsagersHopital.html`, `FormPersonnelHopital.html` (libre, sans code) et `FormCommunauteHopital.html`, générés à partir de vos fichiers XLSForm, avec `hform.js` (moteur commun). Les acteurs communautaires ont été retirés (formulaire, feuille, route serveur).
Les réponses à choix sont enregistrées avec leur **libellé** ; pour enregistrer les codes, mettre `STORE_LABELS = false` en haut de `hform.js`.
