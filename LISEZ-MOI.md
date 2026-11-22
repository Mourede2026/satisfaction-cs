# Déploiement
**URL du déploiement Apps Script actuel** (déjà intégrée dans `config.js`) :
`https://script.google.com/macros/s/AKfycbwbFgMu71sQhLWfLZSo7_UKLkCzy1sNnfju0Z6o0q9EFubPmXC101fxpY125wP8NKu18w/exec`

Après un redéploiement, les appareils qui avaient une adresse corrigée à la main depuis la page Administration reviennent automatiquement à cette adresse.

1. **Apps Script** : remplacez tout le contenu de `Code.gs`, enregistrez, déployez une **nouvelle version**. Si c'est une première installation, exécutez aussi `setupComptesEtQrTokens` (crée la feuille `CodesAcces`) et `setupFeuillesHopitaux` (crée les 3 feuilles hôpital). Rien à exécuter pour les activités : la feuille `Activites` se crée toute seule. Si une feuille `ACTEURS_COM_HOPITAL` existe, supprimez-la : elle n'est plus utilisée. Ne rééxécutez pas `effacerMotsDePasse` si déjà fait.
2. **GitHub** : déposez à la racine tous les fichiers sauf `Code.gs` et `LISEZ-MOI.md`, puis attendez 1-2 min (le cache `sat-v14` se met à jour tout seul).
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

## Tableau de bord par compte, hôpitaux et périodes (`AdminDashboard.html`)
**Qui voit quoi** (compte dans la feuille `Comptes`) :
- `admin`, `national`, `superviseur` : toutes les données (centres de santé et hôpitaux), sur toutes les années.
- `departement` / `zone` : les données de leur département ou de leur zone sanitaire, y compris les hôpitaux de cette zone.
- `commune` : les données de leur commune (centres de santé).
- `hopital` (manager d'hôpital) : uniquement son hôpital. Dans la colonne **Commune** de la feuille `Comptes`, écrire le nom de l'hôpital tel qu'il apparaît dans le formulaire hôpital (ex. `HZ Kandi`). Les accents et majuscules sont ignorés.

**Périodes** : l'administrateur définit les périodes dans l'onglet **⚙️ Périodes** du tableau de bord (année, 2 semestres, 4 trimestres en un clic, ou dates libres, sur plusieurs années). La feuille `Periodes` est créée automatiquement. Chaque utilisateur choisit la période en haut de son tableau de bord ou « Toutes les périodes ». À l'ouverture, la période en cours est choisie automatiquement. L'onglet **📈 Évolution** compare toutes les périodes pour revenir en arrière.

**Rattachement d'une réponse à une période** : par sa date d'envoi. Une colonne `DateEnvoi` est ajoutée automatiquement à l'extrême droite de chaque feuille de réponses à la première réponse reçue après redéploiement. Les anciennes réponses utilisent leur date `today` (centres de santé) ; celles des feuilles hôpital sans date apparaissent dans « Sans date » (visibles en cumul).

**Satisfaction des hôpitaux** : calculée directement à partir des réponses (pourcentage de 0 à 100 par rubrique, moyenne des réponses notées). « Oui » = 1, « Non » = 0 ; Excellent, Bon, Passable, Mauvais ; Très satisfait, etc. répartis de 1 à 0 ; « Ne sait pas » et « Non concerné » sont ignorés. Pour les questions où « Oui » est négatif (sollicitation de cadeaux, difficultés, etc.), le score est inversé. Le barème est dans `HOSP_SCALES` et `HOSP_CFG` de `Code.gs`.

## Activités : un code QR par activité (nouveau)
**Pour l'enquêteur** : il ouvre l'application (ou simplement l'appareil photo du téléphone), **scanne le QR de son activité**, et c'est tout : ses formulaires s'ouvrent directement, sans code à saisir. Sur l'accueil, la carte **📷 Scanner le code QR de mon activité** ouvre la caméra intégrée à l'application (utile quand l'application est déjà installée). La saisie manuelle du code reste possible sur la même page (« Saisir le code d'accès »). Le scan n'est demandé qu'une fois ; hors connexion, les réponses sont gardées sur l'appareil puis envoyées.

**Pour l'administrateur** (`AdminQR.html` → section **📅 Activités**) : créez une activité avec son nom, son niveau (zone sanitaire / hôpital), la zone ou l'hôpital concerné, les formulaires ouverts (aucun coché = tous les formulaires à code), éventuellement des dates de début / fin (pas de choix de commune : les activités ne sont pas limitées à des communes). Le QR s'affiche seul (le code texte est replié sous « Code de secours ») : **⬇️ Image du QR** (PNG haute qualité avec le titre de l'activité), **📤 Partager l'image** (WhatsApp, e-mail…), **📋 Copier le lien**, **Imprimer** (une page par activité, QR seul), **Modifier** (le QR ne change pas), **Nouveau code**, **Désactiver**, **Supprimer**.

**Ce que fait le serveur** : chaque réponse reçue avec le QR d'une activité reçoit automatiquement, dans Google Sheets, les colonnes **Activité** et **IdActivité** (ajoutées à droite à la première réponse). Elles sont posées par le serveur et ne peuvent pas être falsifiées par l'application. Une activité n'ouvre que les formulaires prévus ; avant la date de début ou après la date de fin, le scan est refusé (les réponses d'un enquêteur hors connexion restent acceptées 1 jour après la fin, puis il faut prolonger la date de fin dans l'administration). Désactiver ou supprimer une activité coupe l'accès dans la minute qui suit ; les réponses en attente sur l'appareil sont conservées.

**Gestionnaires et administrateur** : rien ne change, ils continuent d'utiliser le lien de la plateforme (ou l'application installée) avec leur compte. Les codes généraux et les codes par zone / hôpital restent disponibles pour les enquêteurs qui n'ont pas d'activité.

**Limites actuelles** : le scan exige une connexion Internet (le serveur vérifie le code) ; le scanner intégré demande l'autorisation de la caméra et un site en HTTPS (GitHub Pages l'est). Dans les formulaires hôpital, la formation sanitaire reste à choisir dans le formulaire : seule l'étiquette d'activité est automatique.

## Personnel et gestionnaires : QR et liens (v9)
**Personnel** (`AdminQR.html` → « Personnel — un QR par formulaire ») : deux QR séparés, un pour le personnel des **centres de santé (zone sanitaire)** et un pour le personnel des **hôpitaux**. Chacun se télécharge en image PNG haute qualité, se partage ou s'imprime ; le lien correspondant se copie aussi. Le scan (ou le lien) ouvre directement le formulaire, sans code ni identifiant.

**Gestionnaires** (« Gestionnaires — un lien par gestionnaire ») : un lien par compte actif (hors administrateur), de la forme `AdminDashboard.html?u=identifiant`. Le gestionnaire ouvre le lien, son identifiant est déjà rempli, il saisit son mot de passe et accède à son tableau de bord. Le mot de passe n'apparaît jamais dans le lien. Un QR du lien est disponible (à afficher, télécharger en image ou imprimer).

## Factures des enquêteurs et attestations de bonne fin d'exécution (v10)
**Installation (une seule fois)** : dans Apps Script, remplacez `Code.gs`, puis exécutez la fonction **`setupEnqueteursEtFactures`**. **Aucune autorisation Google Drive n'est nécessaire** : les PDF sont enregistrés directement dans Google Sheets (feuille `FacturesFichiers`, découpés en morceaux). Déployez ensuite une **nouvelle version**. Quatre feuilles se créent : `Enqueteurs`, `Factures`, `FacturesFichiers`, `AttestParams`. Ne modifiez ni ne triez la feuille `FacturesFichiers` à la main. Les anciennes factures déjà stockées sur Drive (s'il y en a) restent lisibles. Dans GitHub, ajoutez aussi le dossier `att/` (polices, logo, tampon) et les fichiers `Facture.html`, `AdminAttestations.html`, `attestation.js`.

**Pour l'enquêteur** (`Facture.html`, carte **🧾 Facture et attestation** de l'accueil) : il identifie son activité (QR ou code ; l'activité peut être terminée depuis moins de 60 jours), saisit **nom** et **prénom**, ajoute une ou plusieurs factures normalisées en **PDF** (5 Mo maximum chacune) et envoie. Il peut revenir ajouter d'autres factures : elles se rattachent à la même personne (même nom et prénom, même activité).

**Pour l'administrateur principal seulement** (`AdminAttestations.html`, lien « 🎓 Attestations » dans Administration ; compte de niveau `admin`) :
1. Choisir l'activité (période de collecte) : la liste de tous les enquêteurs s'affiche avec leurs factures (boutons *Ouvrir* et ⬇️ *Télécharger*, et **🧾 Télécharger toutes les factures (ZIP)** pour l'activité), modification possible du nom/prénom, suppression.
2. **Paramètres de l'attestation** (par activité) : zone sanitaire ou hôpital affiché, dates de collecte, année, lieu et **date de signature**. Les champs sont préremplis depuis l'activité ; la date de signature est obligatoire.
3. Pour chaque enquêteur, un **bouton bascule « Attestation validée »** (ou *Valider tous ceux qui ont déposé une facture*).
4. Pour les seuls enquêteurs validés : **⬇️ Attestation PDF** (une personne) ou **ZIP** (toutes les attestations validées de l'activité).

**Le PDF** : une page A4 paysage fidèle au canevas (cadre, logo, tampon et signature), type *Zone Sanitaire* ou *Hôpital* selon l'activité. Le contenu est aplati en image et le fichier est protégé (impression autorisée, modification et copie interdites). La génération se fait dans le navigateur (connexion Internet requise pour charger jsPDF et JSZip).

**Attestation côté enquêteur (v12)** : dans `Facture.html`, la carte **🎓 Mon attestation** vérifie automatiquement l'état du dossier (nom + prénom + code de l'activité). Tant que l'administrateur n'a pas validé, l'enquêteur voit « en attente ». Dès la validation (bouton bascule dans `AdminAttestations.html`, avec les paramètres de l'attestation enregistrés dont la date de signature), le bouton **⬇️ Télécharger mon attestation (PDF)** apparaît. Côté administrateur, **Ouvrir** affiche un aperçu intégré de la facture, avec bouton Télécharger.

## Suivi de la collecte, superviseurs d'activité, objectifs et effectif du personnel (v13)
**Installation** : remplacez `Code.gs`, exécutez une fois **`setupSuivi`** (crée les feuilles `Affectations`, `Objectifs`, `Effectifs`), déployez une nouvelle version, puis déposez les fichiers sur GitHub (nouvelle page `AdminSuivi.html`, liens ajoutés dans l'accueil, Administration, Tableau de bord et Attestations).

**Administrateur** (`AdminSuivi.html` → onglet **Configuration**, pour l'activité choisie) :
1. **Superviseurs** : créer un compte (identifiant, nom, mot de passe), cocher ceux qui suivent l'activité, copier leur lien de connexion, changer le mot de passe, désactiver ou supprimer. Un superviseur d'activité (niveau `superviseur_act`) ne voit que le suivi des activités qui lui sont affectées.
2. **Objectifs** : nombre de questionnaires attendu par enquêteur et par cible (Usagers, COGECS ou Communauté selon le niveau), avec « même objectif pour tous ». Les enquêteurs sont détectés à partir des réponses de l'activité ; on peut en ajouter à l'avance.
3. **Effectif du personnel** : total et/ou détail par structure.
Penser à **Enregistrer la configuration**.

**Suivi (administrateur et superviseurs)** : évolution cumulée par enquêteur jour après jour avec le **niveau moyen attendu** (objectif moyen étalé linéairement sur la période), tableau par enquêteur (réalisé, objectif, attendu à ce jour, écart, avancement, statut, rythme réel et requis, dernière saisie) et **point du personnel** (effectif, réponses reçues, reste à enquêter, taux de participation, détail par structure). Actualisation automatique toutes les minutes. **📷 Exporter en image** produit un PNG (partage direct possible sur téléphone).

**Règles de calcul** : le réalisé est compté par le code QR de l'activité (colonne `IdActivité`), l'enquêteur par la colonne « Identification de l'enquêteur » (ou « Enqueteur » en hôpital), sans tenir compte des accents ni des majuscules. Le personnel n'utilisant pas de code, ses réponses sont comptées sur la période de l'activité (et, pour un hôpital, sur la formation sanitaire correspondant au nom de l'activité).


## Inscription des enquêteurs, objectif unique et montant à payer (v14)
**Installation** : remplacez `Code.gs`, exécutez une fois **`setupSuivi`** (crée la feuille `SuiviParams` et ajoute la colonne `Numéro` à la feuille `Enqueteurs`), déployez une **nouvelle version**, puis déposez les fichiers sur GitHub (nouvelle page `Enqueteur.html`).

**1. Inscription à la première connexion** : après le scan du code QR de l'activité, l'enquêteur renseigne **nom, prénom et numéro d'enquêteur** (`Enqueteur.html`). Le numéro est unique dans l'activité (« 07 » et « 7 » sont le même numéro) ; un numéro déjà pris par une autre personne est refusé. Ensuite, dans tous les formulaires d'enquête, le champ « Identification de l'enquêteur » n'est plus une liste « Enquêteur 1, 2… » : il est rempli **automatiquement avec le prénom**, et le serveur ajoute les colonnes `IdEnquêteur` et `N° enquêteur` (non falsifiables). Le lien « Ce n'est pas moi » du bandeau permet de changer d'enquêteur sur un appareil. Sans activité (simple code d'accès), l'ancienne liste déroulante reste en place.

**2. Facture et attestation** : la page `Facture.html` reprend l'identité de l'inscription (nom, prénom, numéro). L'enquêteur y dépose sa facture normalisée à la fin du travail, puis télécharge son attestation après validation. Les enquêteurs déjà enregistrés avant la v14 (sans numéro) sont retrouvés par nom et prénom, et leur numéro est complété à leur prochaine inscription ou dépôt.

**3. Suivi** (`AdminSuivi.html`) : la liste des enquêteurs est celle des **inscrits** (même sans aucune réponse). Les réponses sont rattachées à chaque enquêteur par son identifiant ; les réponses sans inscription (anciennes, ou code sans activité) apparaissent à part, marquées « non inscrit ».

**4. Objectif unique** (onglet Configuration) : un seul nombre de questionnaires **par cible**, identique pour tous les enquêteurs ; il s'applique automatiquement à chaque inscrit, y compris à ceux qui s'inscrivent après. D'anciens objectifs individuels (v13) sont repris comme objectif commun (valeur la plus fréquente) puis remplacés à l'enregistrement.

**5. Montant total à payer** (onglet Configuration, en FCFA) : réparti à parts égales entre les enquêteurs inscrits. Tarif par questionnaire = montant ÷ (nombre d'inscrits × objectif). Le per diem de chacun = tarif × questionnaires réalisés, **plafonnés à l'objectif de chaque cible** ; un enquêteur qui atteint son objectif perçoit sa part complète, sinon elle est réduite au prorata. La page de suivi affiche la section **💰 Paiement des enquêteurs** (montant total, part, tarif, total dû, reste non versé, montant par enquêteur) et l'inclut dans l'export en image.

**Droits** : l'administrateur configure tout. Un superviseur d'activité peut désormais régler l'objectif et le montant des activités qui lui sont affectées ; seuls l'administrateur gère les comptes superviseurs, leurs affectations et l'effectif du personnel.
