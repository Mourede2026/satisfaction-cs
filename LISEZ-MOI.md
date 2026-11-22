# Déploiement
**URL du déploiement Apps Script actuel** (déjà intégrée dans `config.js`) :
`https://script.google.com/macros/s/AKfycbwbFgMu71sQhLWfLZSo7_UKLkCzy1sNnfju0Z6o0q9EFubPmXC101fxpY125wP8NKu18w/exec`

Après un redéploiement, les appareils qui avaient une adresse corrigée à la main depuis la page Administration reviennent automatiquement à cette adresse.

1. **Apps Script** : remplacez tout le contenu de `Code.gs`, enregistrez, déployez une **nouvelle version**, puis exécutez `setupComptesEtQrTokens` (crée la feuille `CodesAcces`) et `setupFeuillesHopitaux` (crée les 3 feuilles hôpital). Si une feuille `ACTEURS_COM_HOPITAL` existe, supprimez-la : elle n'est plus utilisée. Ne rééxécutez pas `effacerMotsDePasse` si déjà fait.
2. **GitHub** : déposez à la racine tous les fichiers sauf `Code.gs` et `LISEZ-MOI.md`, puis attendez 1-2 min (le cache `sat-v6` se met à jour tout seul).
3. **Utilisation** : l'accueil propose *Collecter des données* (Zone sanitaire / Hôpital) et *Administration*.
4. **Codes** : `AdminQR.html` (connexion admin, depuis téléphone ou tablette) permet de créer un code général ou un code par zone sanitaire / par hôpital, de le partager, imprimer, changer, désactiver ou supprimer.
5. Le personnel (CS et hôpital) n'a jamais besoin de code. Un enquêteur saisit son code une seule fois par niveau.
6. Ne publiez jamais `Code.gs` dans le dépôt public.

## Créer ou réinitialiser un compte directement dans Google Sheets
Feuille `Comptes` : ajoutez **à droite de la dernière colonne utilisée** une colonne dont le titre (ligne 1) est exactement `MotDePasseInitial`. Puis, sur la ligne du compte (ex. `admin`, Niveau `admin`, Actif `OUI`), tapez le mot de passe choisi dans cette colonne.
À la première connexion, l'application vérifie ce mot de passe, le transforme en empreinte sécurisée et **efface la cellule**. Seule cette colonne (repérée par son titre) est lue ; aucune autre colonne n'est jamais prise pour un mot de passe.

## Formulaires hôpital (3)
`FormUsagersHopital.html`, `FormPersonnelHopital.html` (libre, sans code) et `FormCommunauteHopital.html`, générés à partir de vos fichiers XLSForm, avec `hform.js` (moteur commun). Les acteurs communautaires ont été retirés (formulaire, feuille, route serveur).
Les réponses à choix sont enregistrées avec leur **libellé** ; pour enregistrer les codes, mettre `STORE_LABELS = false` en haut de `hform.js`.

## Liens directs pour le personnel
Dans `AdminQR.html` (connecté en administrateur), la section **Liens directs pour le personnel** donne deux liens à copier ou partager dans les forums : `FormPersonnels.html?solo=1` (personnel du centre de santé) et `FormPersonnelHopital.html?solo=1` (personnel de l'hôpital). Chaque lien ouvre uniquement son formulaire, sans lien de retour vers l'accueil, et les réponses vont dans la base correspondante.

## Enregistrement et synchronisation automatique
Dans chaque formulaire, « Enregistrer la réponse » envoie directement les données au serveur. Hors connexion (ou si le serveur ne répond pas), la réponse est conservée sur l'appareil, puis envoyée automatiquement : dès le retour de la connexion et par nouvelle tentative toutes les 30 secondes tant qu'il reste des réponses en attente. Chaque réponse porte un identifiant unique (anti-doublon côté serveur) et une réponse ajoutée pendant une synchronisation ne peut plus être perdue.
