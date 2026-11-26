// URL du déploiement Apps Script (nouveau déploiement) : https://script.google.com/macros/s/AKfycbwbFgMu71sQhLWfLZSo7_UKLkCzy1sNnfju0Z6o0q9EFubPmXC101fxpY125wP8NKu18w/exec
// Elle doit être identique à API_URL_DEFAULT dans config.js.
/**
 * ENQUÊTES DE SATISFACTION - ZONES SANITAIRES DU BÉNIN — Backend Apps Script (API JSON)
 * Version complète et corrigée : accès enquêteurs par QR (code changeable), comptes par niveau,
 * filtrage serveur des données, anti-doublons, limite d'essais de connexion.
 * Sécurité : AUCUN mot de passe n'est écrit dans ce code ni dans la feuille.
 */
var SHEET_USAGERS = "👩‍👦 USAGERS", SHEET_PERSONNELS = "👩‍⚕️ PERSONNELS", SHEET_COGECS = "👨‍👩‍👧‍👦 COGECS";
var T_USAGERS = "👩‍👦 T_USAGERS", T_PERSONNELS = "👩‍⚕️ T_PERSONNELS", T_COGECS = "👨‍👩‍👧‍👦 T_COGECS";
var KEY_COLUMN = "Commune";
var ACCOUNTS_SHEET_NAME = 'Comptes', SESSIONS_SHEET_NAME = 'Sessions', QRTOKENS_SHEET_NAME = 'QRTokens', ENVOIS_SHEET_NAME = 'Envois';
var SESSION_DURATION_MS = 12 * 60 * 60 * 1000;
var QR_ROLES = ['enqueteur', 'personnel', 'enqueteur_hopital', 'personnel_hopital'];
var QR_GUARD = {
  submitUsagers: ['enqueteur'], submitCogecs: ['enqueteur'], // submitPersonnels : libre d'accès (pas de code)
  submitUsagersHopital: ['enqueteur_hopital'], // submitPersonnelHopital : libre d'accès (pas de code)
  submitCommunauteHopital: ['enqueteur_hopital']
};

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(e) { return json_({ ok: true, message: "API Satisfaction Zones Sanitaires en ligne. Utilisez POST." }); }

function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents), action = body.action, d = body.data || {};
    var need = QR_GUARD[action], cid = d._cid, isSubmit = /^submit/.test(action);
    if (need && !qrTokenValid_(body.qr, need)) return json_({ ok: false, error: 'QR_INVALID' });
    if (isSubmit) {
      // Colonnes d'activité : posées uniquement par le serveur à partir du code scanné (jamais par le client)
      delete d[ACT_COL]; delete d[ACT_ID_COL]; delete d[ENQ_ID_COL]; delete d[ENQ_NUM_COL];
      var enqId = d._enq; delete d._enq;
      var act = need ? activityByCode_(body.qr) : null;
      if (act) {
        var fk = FORM_KEY_[action];
        if (act.forms.length && fk && act.forms.indexOf(fk) === -1) return json_({ ok: false, error: 'FORM_NOT_ALLOWED' });
        d[ACT_COL] = act.name; d[ACT_ID_COL] = act.id;
        // Enquêteur inscrit : identité posée par le serveur (prénom + numéro), jamais par le client
        var en = enqId ? enqById_(act.id, enqId) : null;
        if (en && d._dev) { var dst = devCheck_(act.id, en.id, d._dev); if (dst !== 'OK') return json_({ ok: false, error: 'APPAREIL_' + dst }); }
        if (en) { d[ENQ_ID_COL] = en.id; d[ENQ_NUM_COL] = en.numero; d["Identification de l'enquêteur"] = en.prenom; d['Enqueteur'] = en.prenom; }
      }
    }
    if (isSubmit && dejaRecu_(cid)) return json_({ ok: true, duplicate: true });
    // Enregistrement audio obligatoire (Usagers, Usagers Hôpital, COGECS, Communauté) : un fichier par enquête sur Google Drive,
    // le lien (URL) est écrit dans la colonne « Audio ». Si Drive échoue, repli : stockage par morceaux dans Google Sheets.
    var audioB64 = d._audio_b64, audioMime = d._audio_mime, audioDur = d._audio_dur; delete d._audio_b64; delete d._audio_mime; delete d._audio_dur;
    if (isSubmit && audioB64 && FORM_KEY_[action]) {
      try { d[AUDIO_COL] = saveAudioToDrive_(cid, FORM_KEY_[action], audioMime, audioB64, audioDur); }
      catch (driveErr) {
        try { saveAudioChunks_(cid, FORM_KEY_[action], audioMime, audioB64); d[AUDIO_COL] = cid || 'Oui'; }
        catch (audioErr) { d[AUDIO_COL] = 'Erreur enregistrement audio : ' + audioErr.message; }
      }
    }
    if (action === "submitUsagers") out = appendRecord_(SHEET_USAGERS, d);
    else if (action === "submitPersonnels") out = appendRecord_(SHEET_PERSONNELS, d);
    else if (action === "submitCogecs") out = appendRecord_(SHEET_COGECS, d);
    else if (action === "submitUsagersHopital") out = appendRecord_(SHEET_USAGERS_H, d, 'dept');
    else if (action === "submitPersonnelHopital") out = appendRecord_(SHEET_PERSONNELS_H, d, 'dept');
    else if (action === "submitCommunauteHopital") out = appendRecord_(SHEET_COMMUNAUTE_H, d, 'dept');
    else if (action === "checkQr") out = { ok: true, valid: qrTokenValid_(d.token, d.level === 'hopital' ? ['enqueteur_hopital'] : d.level === 'zone' ? ['enqueteur'] : QR_ROLES) };
    else if (action === "registerEnqueteur") out = handleRegisterEnqueteur(d);
    else if (action === "myCounts") out = handleMyCounts(d);
    else if (action === "setMomo") out = handleSetMomo(d);
    else if (action === "comptaGet") out = handleComptaGet(d.token, d.actId);
    else if (action === "comptaPay") out = handleComptaPay(d.token, d.actId, d.enqId, d.paid);
    else if (action === "createAccountant") out = handleCreateAccountant(d.token, d.username, d.label, body.password);
    else if (action === "submitFacture") out = handleSubmitFacture(d);
    else if (action === "matchFactureCode") out = handleMatchFactureCode(d.text);
    else if (action === "getAttestation") out = handleGetAttestation(d);
    else if (action === "suiviList") out = handleSuiviList(d.token);
    else if (action === "suiviGet") out = handleSuiviGet(d.token, d.actId);
    else if (action === "suiviDevices") out = handleSuiviDevices(d.token, d.actId);
    else if (action === "suiviDeviceDecision") out = handleSuiviDeviceDecision(d.token, d.actId, d.enqId, d.dev, d.decision);
    else if (action === "suiviConfig") out = handleSuiviConfig(d.token, d.actId);
    else if (action === "suiviSave") out = handleSuiviSave(d.token, d.actId, d.cfg);
    else if (action === "createSupervisor") out = handleCreateSupervisor(d.token, d.username, d.label, body.password);
    else if (action === "setSupervisorPassword") out = handleSetSupervisorPassword(d.token, d.username, body.password);
    else if (action === "listEnqueteurs") out = handleListEnqueteurs(d.token);
    else if (action === "getFacture") out = handleGetFacture(d.token, d.id);
    else if (action === "setAttestationValide") out = handleSetAttestationValide(d.token, d.ids, d.valide);
    else if (action === "updateEnqueteur") out = handleUpdateEnqueteur(d.token, d.id, d.nom, d.prenom);
    else if (action === "deleteEnqueteur") out = handleDeleteEnqueteur(d.token, d.id);
    else if (action === "deleteFacture") out = handleDeleteFacture(d.token, d.id);
    else if (action === "saveAttestParams") out = handleSaveAttestParams(d.token, d.params);
    else if (action === "listAccessCodes") out = handleListAccessCodes(d.token);
    else if (action === "saveAccessCode") out = handleSaveAccessCode(d.token, d.level, d.name, d.code);
    else if (action === "setAccessCodeActive") out = handleSetAccessCodeActive(d.token, d.level, d.name, d.active);
    else if (action === "deleteAccessCode") out = handleDeleteAccessCode(d.token, d.level, d.name);
    else if (action === "listActivities") out = handleListActivities(d.token);
    else if (action === "saveActivity") out = handleSaveActivity(d.token, d.activity);
    else if (action === "setActivityActive") out = handleSetActivityActive(d.token, d.id, d.active);
    else if (action === "deleteActivity") out = handleDeleteActivity(d.token, d.id);
    else if (action === "getDashboardData") out = getDashboardData(d.token, d.filter);
    else if (action === "listPeriods") out = handleListPeriods(d.token);
    else if (action === "savePeriod") out = handleSavePeriod(d.token, d.period);
    else if (action === "deletePeriod") out = handleDeletePeriod(d.token, d.id);
    else if (action === "accountLogin") out = handleAccountLogin(d.username, body.password);
    else if (action === "changePassword") out = handleChangePassword(d.token, body.password, d.newPassword);
    else if (action === "changeUsername") out = handleChangeUsername(d.token, body.password, d.newUsername);
    else if (action === "getQrToken") out = handleGetQrToken(d.token, d.role);
    else if (action === "refreshQrToken") out = handleRefreshQrToken(d.token, d.role);
    else if (action === "setQrCode") out = handleSetQrCode(d.token, d.role, d.code);
    else if (action === "matchQrCode") out = handleMatchQrCode(d.text);
    else if (action === "listAccounts") out = handleListAccounts(d.token);
    else if (action === "setAccountActive") out = handleSetAccountActive(d.token, d.username, d.active);
    else if (action === "updateAccount") out = handleUpdateAccount(d.token, d.username, d.fields);
    else if (action === "deleteAccount") out = handleDeleteAccount(d.token, d.username);
    else out = { ok: false, error: "Action inconnue : " + action };
    if (isSubmit && out && out.ok) noterRecu_(cid);
  } catch (err) { out = { ok: false, error: err.message }; }
  return json_(out);
}

/* ---------- Enregistrement des réponses ---------- */
function buildHeaderMap_(sheet) {
  var lastCol = sheet.getLastColumn(), headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0], map = {};
  for (var i = 0; i < headers.length; i++) { var h = (headers[i] || "").toString().trim(); if (h) map[h] = i + 1; }
  return { map: map, lastCol: lastCol };
}
function getNextRow_(sheet, keyCol) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  var lastVal = sheet.getRange(lastRow, keyCol, 1, 1).getValue();
  if (lastVal !== "" && lastVal !== null) return lastRow + 1;
  var values = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) if (values[i][0] !== "" && values[i][0] !== null) return i + 3;
  return 2;
}
function appendRecord_(sheetName, data, keyColumnName) {
  keyColumnName = keyColumnName || KEY_COLUMN;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) throw new Error("Feuille introuvable: " + sheetName);
    var hm = buildHeaderMap_(sheet), keyCol = hm.map[keyColumnName];
    var cols_ = [DATE_COLUMN]; if (data[ACT_COL]) cols_.push(ACT_COL, ACT_ID_COL); if (data[ENQ_ID_COL]) cols_.push(ENQ_ID_COL, ENQ_NUM_COL);
    if (data._t_start || data.start || data.end) cols_.push(START_COL, END_COL);
    if (data._dev) cols_.push(DEV_COL);
    if (data[AUDIO_COL]) cols_.push(AUDIO_COL);
    cols_.forEach(function (nm) {
      if (hm.map[nm]) return;
      sheet.getRange(1, hm.lastCol + 1).setValue(nm);
      try { sheet.getRange(2, hm.lastCol + 1, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@'); } catch (e) {}
      hm = buildHeaderMap_(sheet);
    });
    data[DATE_COLUMN] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (data._t_start || data.start) data[START_COL] = String(data._t_start || data.start); if (data._t_end || data.end) data[END_COL] = String(data._t_end || data.end);
    if (data._dev) data[DEV_COL] = String(data._dev).slice(0, 60);
    if (!keyCol) throw new Error("Colonne clé '" + keyColumnName + "' introuvable dans " + sheetName);
    var row = getNextRow_(sheet, keyCol), rowValues = new Array(hm.lastCol).fill("");
    for (var key in data) { if (data.hasOwnProperty(key) && hm.map[key]) rowValues[hm.map[key] - 1] = data[key]; }
    sheet.getRange(row, 1, 1, hm.lastCol).setValues([rowValues]);
    SpreadsheetApp.flush();
    return { ok: true, row: row, sheet: sheetName };
  } finally { lock.releaseLock(); }
}
function dejaRecu_(cid) {
  if (!cid) return false;
  var sh = getOrCreateSheet_(ENVOIS_SHEET_NAME, ['cid', 'date']);
  if (sh.getLastRow() < 2) return false;
  return !!sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(cid)).matchEntireCell(true).findNext();
}
function noterRecu_(cid) { if (cid) getOrCreateSheet_(ENVOIS_SHEET_NAME, ['cid', 'date']).appendRow([String(cid), new Date().toISOString()]); }

/* ---------- Audio d'entretien sur Google Drive (lien dans la colonne « Audio ») ----------
   Dossier « Audios enquêtes satisfaction » > sous-dossier par formulaire. Les fichiers restent PRIVÉS
   (visibles seulement par le compte propriétaire du script / l'administrateur). */
var AUDIO_DRIVE_ROOT = 'Audios enquêtes satisfaction';
function audioFolder_(formKey) {
  var props = PropertiesService.getScriptProperties(), pk = 'AUDIO_FOLDER_' + (formKey || 'autre');
  var id = props.getProperty(pk);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var rootId = props.getProperty('AUDIO_FOLDER_ROOT'), root = null;
  if (rootId) { try { root = DriveApp.getFolderById(rootId); } catch (e) {} }
  if (!root) {
    var it = DriveApp.getFoldersByName(AUDIO_DRIVE_ROOT);
    root = it.hasNext() ? it.next() : DriveApp.createFolder(AUDIO_DRIVE_ROOT);
    props.setProperty('AUDIO_FOLDER_ROOT', root.getId());
  }
  var sub = root.getFoldersByName(String(formKey || 'autre'));
  var f = sub.hasNext() ? sub.next() : root.createFolder(String(formKey || 'autre'));
  props.setProperty(pk, f.getId());
  return f;
}
function saveAudioToDrive_(cid, formKey, mime, b64, dur) {
  if (!b64) throw new Error('audio vide');
  var m = String(mime || 'audio/webm'), ext = /ogg/.test(m) ? 'ogg' : /mp4|aac|m4a/.test(m) ? 'm4a' : 'webm';
  var baseMime = m.split(';')[0];
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  var name = (formKey || 'audio') + '_' + stamp + '_' + String(cid || Utilities.getUuid()).slice(0, 8) + '.' + ext;
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), baseMime, name);
  var file = audioFolder_(formKey).createFile(blob);
  if (dur) { try { file.setDescription('Durée : ' + dur + ' s — cid ' + (cid || '')); } catch (e) {} }
  return file.getUrl();
}

/* ---------- Audio d'entretien (obligatoire, enregistré avec mention visible côté enquêteur) ----------
   Stocké en morceaux de texte dans Google Sheets (pas de Drive) : une feuille dédiée, plusieurs lignes
   par enregistrement. Reconstitution : trier par (cid, partie) et concaténer la colonne "contenu". */
function saveAudioChunks_(cid, formKey, mime, b64) {
  if (!cid || !b64) return;
  var CH = 45000; // marge sous la limite ~50000 caractères par cellule Google Sheets
  var total = Math.ceil(b64.length / CH);
  var date = new Date().toISOString(), rows = [];
  for (var i = 0; i < total; i++) rows.push([String(cid), formKey || '', date, mime || '', i + 1, total, b64.substr(i * CH, CH)]);
  var sh = getOrCreateSheet_(AUDIO_SHEET_NAME, ['cid', 'formulaire', 'date', 'mime', 'partie', 'total', 'contenu']);
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var startRow = sh.getLastRow() + 1;
    sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}

/* ---------- Enquêtes hôpitaux (collecte uniquement) ---------- */
var SHEET_USAGERS_H = "🏥 USAGERS_HOPITAL", SHEET_PERSONNELS_H = "🏥 PERSONNEL_HOPITAL", SHEET_COMMUNAUTE_H = "🏥 COMMUNAUTE_HOPITAL";
var HEADERS_USAGERS_H = ["dept", "zone", "fs", "Veuillez_enregistrer_sition_au_moins_90", "Enqueteur", "Sujet_enqu_t", "Hospitalisation", "A1a", "Comment_avez_vous_tr_acc_der_l_h_pital_", "Avez_vous_rencontr_re_v_hicule_ou_moto_", "L_acc_s_aux_b_timent_mobilit_r_duite_", "Une_fois_l_int_rie_neaux_d_orientation_", "La_propret_et_le_co_ils_au_rendez_vous_", "Avez_vous_des_suggestions_part", "Si_oui_lesquelles", "Avez_vous_des_diffic_au_sein_de_l_h_pital", "A1", "_3_Qu_est_ce_que_vous_n_avez_pas_aim", "A2", "A3", "_6_Qu_est_ce_que_vous_n_avez_pas_aim", "A3a", "A4", "_9_Qu_est_ce_que_vous_n_avez_pas_aim", "A4a", "A5", "_12_Qu_est_ce_que_vous_n_avez_pas_aim", "A5a", "A6", "_15_Qu_est_ce_que_vous_n_avez_pas_aim", "A6a", "C1", "_2_Dites_nous_un_peu_comment_a_a_t", "C2", "D1", "_5_Dites_nous_un_peu_comment_a_a_t", "D2", "E1", "B1", "B2", "B3", "G1", "G", "G2", "F1", "F2", "F3", "C3", "E2", "E3", "E4", "A7", "A8", "D2_Precise_001", "couleur_tenue", "C4", "C5", "H1", "H2", "Avant_qu_une_d_cisio_vous_t_consult_e_", "Avez_vous_pu_exprime_les_soins_propos_s_", "Vous_est_il_d_j_arriv_de_ref", "Si_oui_votre_d_cisi_ression_ni_jugement_", "Lors_de_votre_prise_en_charge_", "Avez_vous_effectivement_d_sign", "Si_oui_cette_person_essionnels_de_sant_", "Lors_de_votre_prise_en_charge", "Lors_de_votre_prise_en_charge_001", "Lors_de_votre_prise_en_charge__001", "H3", "H4", "H5", "H6", "H7", "C6", "H8", "AA1", "D1_Precise", "AA2", "D2_Precise", "AA3", "D3_Precise", "AA4", "D4_Precise", "I1", "Si_non_pourquoi", "Si_oui_pourquoi", "I2", "Si_non_pourquoi_001", "Si_oui_pourquoi_001", "I3", "I4", "I5", "Age", "Sexe", "departement", "commune", "Profession", "Instruction", "Raison_Choix", "Avez_vous_autre_chose_dire", "today", "_start-geopoint_latitude", "_start-geopoint_longitude", "_start-geopoint_altitude", "_start-geopoint_precision"];
var HEADERS_PERSONNELS_H = ["dept", "zs", "fs", "Formulaire_renseign_par", "Choisir_l_identifiant", "age", "Sexe", "Q1", "Q2", "Q3", "Avez_vous_re_u_une_f_evoirs_des_patients_", "Vous_sentez_vous_sou_evoirs_des_patients_", "Avez_vous_connaissance_des_dro", "Ces_droits_et_devoir_votre_tablissement_", "Pensez_vous_que_vos_onnement_de_travail_", "Avez_vous_des_difficult_s_dans", "Quelles_sont_les_pri_cice_de_vos_devoirs_", "Q4", "Q5", "Q6", "Q7", "Q8", "Recevez_vous_les_inf_d_lais_raisonnables_", "Quels_sont_les_canau_ut_ce_qui_s_applique", "Q9", "Les_coll_gues_de_vot_protocoles_de_soins_", "Les_coll_gues_assure_esoins_des_patients_", "Les_coll_gues_collab_ontinuit_des_soins_", "Q10", "Q11", "Q12", "Q13", "Q14", "Q15", "Q16", "Q17", "Q18", "Q19", "Si_oui_de_quelle_mani_re", "Quelle_autre_mani_re", "Quelles_sont_les_dif_cc_s_l_information", "Q20", "Q21", "Q22", "Vous_pensez_que_l_av_chnique_de_mes_soins", "Vous_pensez_que_l_im_d_erreurs_m_dicales", "Je_re_ois_r_guli_rem_ction_de_mon_service", "Les_plaintes_des_pat_difier_nos_pratiques", "Le_manque_de_temps_e_des_retours_patients", "Les_retours_des_pati_utiles_cliniquement", "Pouvez_vous_citer_un_exemple_r", "Si_oui_pr_ciser", "Selon_vous_quels_so_l_h_pital_de_Pob_", "_9_Quels_sont_les_pri_antage_les_patients_", "Q23", "Q24"];
var HEADERS_COMMUNAUTE_H = ["dept", "zone", "Formation_Sanitaire", "Enqueteur", "Comment_avez_vous_tr_acc_der_l_h_pital_", "Avez_vous_rencontr_re_v_hicule_ou_moto_", "L_acc_s_aux_b_timent_mobilit_r_duite_", "Une_fois_l_int_rie_neaux_d_orientation_", "La_propret_et_le_co_ils_au_rendez_vous_", "Avez_vous_des_suggestions_part", "Si_oui_lesquelles", "Comment_jugez_vous_e_accueil_au_portail_", "Comment_jugez_vous_e_l_dans_les_services_", "Comment_jugez_vous_e_au_service_de_sortie", "Comment_trouvez_vous_les_dirigeants_", "Comment_appr_ciez_vo_es_agents_examiner", "Comment_appr_ciez_vo_diguer_de_bons_soins", "Comment_appr_ciez_vo_e_bonne_prescription", "Comment_appr_ciez_vo_niquer_sur_les_actes", "Pensez_vous_que_les_la_prise_en_charge_", "Comment_trouvez_vous_rapport_avec_le_mal", "Pensez_vous_que_les_ous_servir_au_besoin", "Les_m_dicaments_pres_ibles_dans_le_centre", "Les_m_dicaments_non_pharmacies_proches_", "Comment_trouvez_vous_ifs_des_m_dicaments_", "Comment_trouvez_vous_ents_hospitalisation", "Comment_trouvez_vous_la_cour_de_l_h_pital", "Comment_trouvez_vous_alle_de_consultation", "Comment_trouvez_vous_les_toilettes", "Comment_trouvez_vous_la_buanderie", "Pensez_vous_que_les_onn_es_sur_les_soins", "Pensez_vous_qu_on_a_est_pas_souvent_seul", "Avez_vous_acc_s_l_eau", "Vous_a_t_on_une_fois_des_soins_domicile", "Vous_est_il_jamais_a_r_des_agents_de_sant", "Des_agents_vous_ont_eux_directement_vers", "Pensez_vous_que_les_h_pital_g_rent_bien", "Pensez_vous_que_dans_u_sein_de_l_h_pital_", "Si_vous_devrez_reven_s_motivants_ce_choix", "Age", "Sexe", "Profession", "Instruction", "Avez_vous_autre_chose_dire"];
function setupFeuillesHopitaux() {
  [[SHEET_USAGERS_H, HEADERS_USAGERS_H], [SHEET_PERSONNELS_H, HEADERS_PERSONNELS_H], [SHEET_COMMUNAUTE_H, HEADERS_COMMUNAUTE_H]]
    .forEach(function (p) { getOrCreateSheet_(p[0], p[1]); });
}

/* ---------- Calculs du tableau de bord ---------- */
function readColumn_(sheet, hm, name) {
  var col = hm[name], lastRow = sheet.getLastRow();
  if (!col || lastRow < 2) return [];
  return sheet.getRange(2, col, lastRow - 1, 1).getValues().map(function (r) { return r[0]; });
}
function clampPct_(ratio) { return Math.min(100, Math.round(ratio * 1000) / 10); }
function columnLetterToIndex_(l) { var c = 0; for (var i = 0; i < l.length; i++) c = c * 26 + (l.charCodeAt(i) - 64); return c; }
function computeRowScoresForCategories_(tSheet, categories, nRaw) {
  if (!nRaw || nRaw <= 0 || tSheet.getLastRow() < 3) return [];
  var idxs = [];
  categories.forEach(function (cat) { cat.cols.forEach(function (l) { idxs.push(columnLetterToIndex_(l) - 1); }); });
  var nRows = Math.min(nRaw, tSheet.getLastRow() - 2);
  if (nRows <= 0) return [];
  return tSheet.getRange(3, 1, nRows, Math.max.apply(null, idxs) + 1).getValues().map(function (row) {
    var sum = 0, n = 0;
    idxs.forEach(function (ci) { var v = row[ci]; if (typeof v === "number") { sum += Math.max(0, Math.min(1, v)); n++; } });
    return n > 0 ? sum / n : null;
  });
}
function groupWithScore_(keys, scores) {
  var out = {}, result = [];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (k === "" || k === null || k === undefined) continue;
    k = k.toString();
    if (!out[k]) out[k] = { count: 0, sum: 0, n: 0 };
    out[k].count++;
    if (typeof scores[i] === "number") { out[k].sum += scores[i]; out[k].n++; }
  }
  for (var key in out) result.push({ label: key, count: out[key].count, avgScore: out[key].n > 0 ? clampPct_(out[key].sum / out[key].n) : null });
  result.sort(function (a, b) { return b.count - a.count; });
  return result;
}
function avgOf_(scores) {
  var s = 0, n = 0;
  scores.forEach(function (v) { if (typeof v === "number") { s += v; n++; } });
  return n > 0 ? clampPct_(s / n) : null;
}
var CATEGORY_MAP = {
  usagers: [
    { name: '1. Accueil et orientation', cols: ["G","H","I","J"] },
    { name: '2. Information et communication', cols: ["K","L","M","N","O","P","Q"] },
    { name: '3. Qualité de l’examen et de l’écoute', cols: ["U","V","W","X"] },
    { name: '4. Soins et continuité', cols: ["Y","Z","AA","AB","AC","AD","AE"] },
    { name: '5. Respect, intimité et confidentialité', cols: ["AF","AG","AH","AI"] },
    { name: '6. Accessibilité financière et médicaments', cols: ["AJ","AK","AL","AM","AN","AO"] },
    { name: '7. Environnement et hygiène', cols: ["AQ","AR","AS","AT","AU","AV","AW"] },
    { name: '8. Équité, respect et compétence du personnel', cols: ["AX","AY","AZ","BA","BB","BC","BD","BE"] },
    { name: '9. Transparence et pratiques abusives', cols: ["BF","BG","BH","BI","BJ"] },
    { name: '10. Rapidité et réactivité de la prise en charge', cols: ["BK","BL","BM"] },
    { name: '11. Vaccination', cols: ["R","S","T","AP","BN"] }
  ],
  personnels: [
    { name: 'Ressources matérielles et infrastructures', cols: ["I","J","K"] },
    { name: 'Organisation du travail et charge professionnelle', cols: ["L","M","N"] },
    { name: 'Valorisation, reconnaissance et perspectives professionnelles', cols: ["O","P","Q"] },
    { name: 'Relations interpersonnelles et climat social', cols: ["R","S","T"] },
    { name: 'Protocoles, directives et communication institutionnelle', cols: ["U","V","W"] },
    { name: 'Compétences techniques et fiabilité de l’équipe', cols: ["X","Y","Z"] },
    { name: 'Supervision, encadrement et formation', cols: ["AA","AB","AC","AD","AE","AF"] },
    { name: 'Information et charte des droits des patients', cols: ["AG","AH","AI","AJ","AK","AL","AM","AN"] },
    { name: 'Pratiques éthiques et relation patient', cols: ["AO","AP","AQ","AR","AS"] },
    { name: 'Gouvernance et collaboration avec le COGECS', cols: ["AT","AU","AV","AW","AX"] }
  ],
  cogecs: [
    { name: '1. Gouvernance et gestion du centre de santé', cols: ["E","F","G","H","I","J"] },
    { name: '2. Collaboration interne avec le personnel soignant', cols: ["K","L","M","N"] },
    { name: '3. Relation avec la coordination de zone sanitaire', cols: ["O","P","Q"] },
    { name: '4. Capacités et moyens des membres du COGECS', cols: ["R","S","T"] }
  ]
};
function computeCategoryScores_(tSheet, categories, nRaw) {
  var out = {}, empty = function () { categories.forEach(function (c) { out[c.name] = []; }); return out; };
  if (!nRaw || nRaw <= 0 || tSheet.getLastRow() < 3) return empty();
  var nRows = Math.min(nRaw, tSheet.getLastRow() - 2);
  if (nRows <= 0) return empty();
  var range = tSheet.getRange(3, 1, nRows, tSheet.getLastColumn()).getValues();
  categories.forEach(function (cat) {
    var idxs = cat.cols.map(function (l) { return columnLetterToIndex_(l) - 1; });
    out[cat.name] = range.map(function (row) {
      var s = 0, n = 0;
      idxs.forEach(function (ci) { var v = row[ci]; if (typeof v === "number") { s += Math.max(0, Math.min(1, v)); n++; } });
      return n > 0 ? s / n : null;
    });
  });
  return out;
}
/** Rapport générique par clé (commune / structure / enquêteur). */
function aggReport_(sheetName, tName, cats, keyHeader, withDemo, withCommune, mask) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(sheetName), tSheet = ss.getSheetByName(tName);
  var hm = buildHeaderMap_(sheet).map, keys = readColumn_(sheet, hm, keyHeader), n = keys.length;
  var cs = computeCategoryScores_(tSheet, cats, n), names = cats.map(function (c) { return c.name; });
  var ages = withDemo ? readColumn_(sheet, hm, "Âge") : [], sexes = withDemo ? readColumn_(sheet, hm, "Sexe") : [];
  var coms = withCommune ? readColumn_(sheet, hm, "Commune") : [], by = {};
  for (var i = 0; i < n; i++) {
    if (mask && !mask[i]) continue;
    var k = keys[i];
    if (k === "" || k === null || k === undefined) continue;
    k = k.toString();
    if (!by[k]) { by[k] = { count: 0, ageSum: 0, ageN: 0, g: {}, sum: {}, n: {}, commune: (coms[i] || "").toString() }; names.forEach(function (c) { by[k].sum[c] = 0; by[k].n[c] = 0; }); }
    var r = by[k]; r.count++;
    if (typeof ages[i] === "number") { r.ageSum += ages[i]; r.ageN++; }
    if (sexes[i] !== "" && sexes[i] !== null && sexes[i] !== undefined) { var sk = sexes[i].toString(); r.g[sk] = (r.g[sk] || 0) + 1; }
    names.forEach(function (c) { var s = cs[c][i]; if (typeof s === "number") { r.sum[c] += s; r.n[c]++; } });
  }
  var res = {};
  for (var key in by) {
    var rec = by[key], pct = {}, o;
    names.forEach(function (c) { pct[c] = rec.n[c] > 0 ? clampPct_(rec.sum[c] / rec.n[c]) : null; });
    o = { count: rec.count, categoryScores: pct };
    if (withDemo) {
      var gp = {}; for (var g in rec.g) gp[g] = Math.round((rec.g[g] / rec.count) * 1000) / 10;
      o.avgAge = rec.ageN > 0 ? Math.round((rec.ageSum / rec.ageN) * 10) / 10 : null; o.genderPct = gp;
    }
    if (withCommune) o.commune = rec.commune;
    res[key] = o;
  }
  return { categories: names, res: res };
}
function buildCommuneReport_(s, t, c, mk) { var r = aggReport_(s, t, c, "Commune", true, false, mk); return { categories: r.categories, byCommune: r.res }; }
function buildStructureReport_(s, t, c, mk) { var r = aggReport_(s, t, c, "Structure", false, true, mk); return { categories: r.categories, byStructure: r.res }; }
function buildEnqueteurReport_(s, t, c, mk) { var r = aggReport_(s, t, c, "Identification de l'enquêteur", false, false, mk); return { categories: r.categories, byEnqueteur: r.res }; }

function getEnqueteurPositions_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_USAGERS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var hm = buildHeaderMap_(sheet).map, rc = function (h) { return readColumn_(sheet, hm, h); };
  var dates = rc("today"), enq = rc("Identification de l'enquêteur"), com = rc("Commune"), str = rc("Structure");
  var lats = rc("_start-geopoint_latitude"), lons = rc("_start-geopoint_longitude"), accs = rc("_start-geopoint_precision"), starts = rc("start"), ends = rc("end"), byDay = {};
  for (var i = 0; i < dates.length; i++) {
    var lat = lats[i], lon = lons[i];
    if (lat === "" || lon === "" || lat == null || lon == null || isNaN(Number(lat)) || isNaN(Number(lon))) continue;
    var day = (dates[i] || "Date inconnue").toString(), dur = null;
    if (!byDay[day]) byDay[day] = [];
    if (starts[i] && ends[i]) { var a = new Date(starts[i]).getTime(), b = new Date(ends[i]).getTime(); if (!isNaN(a) && !isNaN(b) && b > a) dur = Math.round((b - a) / 6000) / 10; }
    byDay[day].push({ enqueteur: (enq[i] || "Non renseigné").toString(), commune: (com[i] || "").toString(), structure: (str[i] || "").toString(),
      lat: Number(lat), lon: Number(lon), accuracy: (accs[i] !== "" && accs[i] != null) ? Math.round(Number(accs[i])) : null, durationMin: dur });
  }
  return byDay;
}

/** Données du tableau de bord : réservé aux comptes connectés. filter.periode = id de période (ou « toutes »). */
function getDashboardData(sessionToken, filter) {
  filter = filter || {};
  var session = sessionToken ? findSession_(sessionToken) : null, acc = session ? findAccountRow_(session.username) : null;
  var account = acc ? accountPublicView_(acc.data) : null;
  if (!account || !account.active) return { ok: false, error: "Authentification requise (connectez-vous avec votre compte)." };
  var errors = [];
  function safe_(label, fn, fb) { try { return fn(); } catch (e) { errors.push(label + " : " + e.message); return fb; } }
  var periods = safe_("periodes", periodsList_, []), sel = null;
  if (filter.periode === 'auto') {
    var today_ = day_(new Date()), started = periods.filter(function (p) { return p.start <= today_; });
    sel = periods.filter(function (p) { return p.start <= today_ && p.end >= today_; })[0] || started[0] || null;
  } else if (filter.periode && filter.periode !== 'toutes') periods.forEach(function (p) { if (p.id === filter.periode) sel = p; });
  var ss = SpreadsheetApp.getActiveSpreadsheet(), U = ss.getSheetByName(SHEET_USAGERS), C = ss.getSheetByName(SHEET_COGECS), P = ss.getSheetByName(SHEET_PERSONNELS);
  var tU = ss.getSheetByName(T_USAGERS), tC = ss.getSheetByName(T_COGECS), tP = ss.getSheetByName(T_PERSONNELS);
  var uH = U ? buildHeaderMap_(U).map : {}, cH = C ? buildHeaderMap_(C).map : {}, pH = P ? buildHeaderMap_(P).map : {};
  var uCom = safe_("uCom", function () { return readColumn_(U, uH, "Commune"); }, []);
  var uStr = safe_("uStr", function () { return readColumn_(U, uH, "Structure"); }, []);
  var uEnq = safe_("uEnq", function () { return readColumn_(U, uH, "Identification de l'enquêteur"); }, []);
  var uSc = safe_("uSc", function () { return computeRowScoresForCategories_(tU, CATEGORY_MAP.usagers, uCom.length); }, []);
  var uDay = safe_("uDay", function () { return rowDays_(U, uH); }, []);
  var cCom = safe_("cCom", function () { return readColumn_(C, cH, "Commune"); }, []);
  var cStr = safe_("cStr", function () { return readColumn_(C, cH, "Structure"); }, []);
  var cSc = safe_("cSc", function () { return computeRowScoresForCategories_(tC, CATEGORY_MAP.cogecs, cCom.length); }, []);
  var cDay = safe_("cDay", function () { return rowDays_(C, cH); }, []);
  var pCom = safe_("pCom", function () { return readColumn_(P, pH, "Commune"); }, []);
  var pStr = safe_("pStr", function () { return readColumn_(P, pH, "Structure"); }, []);
  var pSc = safe_("pSc", function () { return computeRowScoresForCategories_(tP, CATEGORY_MAP.personnels, pCom.length); }, []);
  var pDay = safe_("pDay", function () { return rowDays_(P, pH); }, []);
  // portée du compte (par ligne) puis période choisie
  var memo = {};
  function inScope_(c) { var k = String(c === undefined || c === null ? '' : c); if (k === '') return false; if (!(k in memo)) memo[k] = communeInScope_(k, account); return memo[k]; }
  function scopeMask_(com) { return com.map(inScope_); }
  function periodMask_(sm, days) { return sm.map(function (ok, i) { return ok && inPeriod_(days[i], sel); }); }
  function pick_(a, m) { var o = []; for (var i = 0; i < m.length; i++) if (m[i]) o.push(a[i]); return o; }
  var sU = scopeMask_(uCom), sC = scopeMask_(cCom), sP = scopeMask_(pCom);
  var mU = periodMask_(sU, uDay), mC = periodMask_(sC, cDay), mP = periodMask_(sP, pDay);
  // évolution par période (toutes les périodes, dans la portée du compte)
  function trendOf_(sm, days, sc, p, undated) {
    var n = 0, s = 0, sn = 0;
    for (var i = 0; i < sm.length; i++) {
      if (!sm[i]) continue;
      if (undated) { if (days[i]) continue; } else if (!inPeriod_(days[i], p)) continue;
      n++; if (typeof sc[i] === 'number') { s += sc[i]; sn++; }
    }
    return { count: n, avgScore: sn > 0 ? clampPct_(s / sn) : null };
  }
  var hosp = safe_("hopitaux", function () { return hospitalData_(account, sel, periods); }, { data: {}, trend: {}, trendUndated: {}, errors: [] });
  hosp.errors.forEach(function (e) { errors.push(e); });
  var trend = periods.slice().sort(function (a, b) { return a.start < b.start ? -1 : a.start > b.start ? 1 : 0; }).map(function (p) {
    var h = hosp.trend[p.id] || {};
    return { id: p.id, label: p.label, start: p.start, end: p.end,
      usagers: trendOf_(sU, uDay, uSc, p), personnels: trendOf_(sP, pDay, pSc, p), cogecs: trendOf_(sC, cDay, cSc, p),
      hUsagers: h.usagers || { count: 0, avgScore: null }, hPersonnel: h.personnel || { count: 0, avgScore: null }, hCommunaute: h.communaute || { count: 0, avgScore: null } };
  });
  var und = { id: 'sans', label: 'Sans date', usagers: trendOf_(sU, uDay, uSc, null, true), personnels: trendOf_(sP, pDay, pSc, null, true), cogecs: trendOf_(sC, cDay, cSc, null, true),
    hUsagers: hosp.trendUndated.usagers || { count: 0, avgScore: null }, hPersonnel: hosp.trendUndated.personnel || { count: 0, avgScore: null }, hCommunaute: hosp.trendUndated.communaute || { count: 0, avgScore: null } };
  if (und.usagers.count + und.personnels.count + und.cogecs.count + und.hUsagers.count + und.hPersonnel.count + und.hCommunaute.count > 0) trend.push(und);
  // données de la période choisie
  uCom = pick_(uCom, mU); uStr = pick_(uStr, mU); uEnq = pick_(uEnq, mU); uSc = pick_(uSc, mU);
  cCom = pick_(cCom, mC); cStr = pick_(cStr, mC); cSc = pick_(cSc, mC);
  pCom = pick_(pCom, mP); pStr = pick_(pStr, mP); pSc = pick_(pSc, mP);
  var nz = function (a) { return a.filter(function (v) { return v !== ""; }).length; };
  var uN = nz(uCom), cN = nz(cCom), pN = nz(pCom), emptyC = { categories: [], byCommune: {} };
  var result = {
    ok: true, counts: { usagers: uN, cogecs: cN, personnels: pN, total: uN + cN + pN },
    periods: periods, selectedPeriod: sel ? sel.id : 'toutes', trend: trend,
    hopital: hosp.data,
    usagers: { byCommune: safe_("u.com", function () { return groupWithScore_(uCom, uSc); }, []), byStructure: safe_("u.str", function () { return groupWithScore_(uStr, uSc); }, []), byEnqueteur: safe_("u.enq", function () { return groupWithScore_(uEnq, uSc); }, []) },
    cogecs: { byCommune: safe_("c.com", function () { return groupWithScore_(cCom, cSc); }, []), byStructure: safe_("c.str", function () { return groupWithScore_(cStr, cSc); }, []) },
    personnels: { byCommune: safe_("p.com", function () { return groupWithScore_(pCom, pSc); }, []), byStructure: safe_("p.str", function () { return groupWithScore_(pStr, pSc); }, []) },
    global: {
      byCible: [{ label: "Usagers", count: uN, avgScore: safe_("avgU", function () { return avgOf_(uSc); }, null) }, { label: "Personnel soignant", count: pN, avgScore: safe_("avgP", function () { return avgOf_(pSc); }, null) }, { label: "COGECS", count: cN, avgScore: safe_("avgC", function () { return avgOf_(cSc); }, null) }],
      byCommune: safe_("g.com", function () { return groupWithScore_(uCom.concat(cCom, pCom), uSc.concat(cSc, pSc)); }, [])
    },
    communeReport: {
      usagers: safe_("cr.u", function () { return buildCommuneReport_(SHEET_USAGERS, T_USAGERS, CATEGORY_MAP.usagers, mU); }, emptyC),
      personnels: safe_("cr.p", function () { return buildCommuneReport_(SHEET_PERSONNELS, T_PERSONNELS, CATEGORY_MAP.personnels, mP); }, emptyC),
      cogecs: safe_("cr.c", function () { return buildCommuneReport_(SHEET_COGECS, T_COGECS, CATEGORY_MAP.cogecs, mC); }, emptyC)
    },
    structureReport: {
      usagers: safe_("sr.u", function () { return buildStructureReport_(SHEET_USAGERS, T_USAGERS, CATEGORY_MAP.usagers, mU); }, { categories: [], byStructure: {} }),
      personnels: safe_("sr.p", function () { return buildStructureReport_(SHEET_PERSONNELS, T_PERSONNELS, CATEGORY_MAP.personnels, mP); }, { categories: [], byStructure: {} }),
      cogecs: safe_("sr.c", function () { return buildStructureReport_(SHEET_COGECS, T_COGECS, CATEGORY_MAP.cogecs, mC); }, { categories: [], byStructure: {} })
    },
    enqueteurReport: {
      usagers: safe_("er.u", function () { return buildEnqueteurReport_(SHEET_USAGERS, T_USAGERS, CATEGORY_MAP.usagers, mU); }, { categories: [], byEnqueteur: {} }),
      cogecs: safe_("er.c", function () { return buildEnqueteurReport_(SHEET_COGECS, T_COGECS, CATEGORY_MAP.cogecs, mC); }, { categories: [], byEnqueteur: {} })
    },
    gpsByDay: safe_("gps", function () { return getEnqueteurPositions_(); }, {}),
    errors: errors, generatedAt: new Date().toISOString(), account: account
  };
  var out = applyScopeToResult_(result, account);
  var hc = out.hopital || {};
  out.counts.hopitalUsagers = (hc.usagers || {}).count || 0; out.counts.hopitalPersonnel = (hc.personnel || {}).count || 0; out.counts.hopitalCommunaute = (hc.communaute || {}).count || 0;
  out.counts.hopital = out.counts.hopitalUsagers + out.counts.hopitalPersonnel + out.counts.hopitalCommunaute;
  out.counts.total = out.counts.usagers + out.counts.personnels + out.counts.cogecs + out.counts.hopital;
  return out;
}

/* ---------- Comptes, sessions ---------- */
function sha256Hex_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(function (b) { return ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0'); }).join('');
}
function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); }
  return sh;
}
/** À exécuter UNE FOIS. Crée les feuilles manquantes, les 4 codes QR, et un compte « admin » si aucun compte n'existe
 *  (son mot de passe s'affiche une seule fois dans Affichage > Journaux). Les comptes existants ne sont pas touchés. */
function setupComptesEtQrTokens() {
  var comptes = getOrCreateSheet_(ACCOUNTS_SHEET_NAME, ['Identifiant', 'Hash', 'Niveau', 'Département', 'Zone Sanitaire', 'Commune', 'Libellé', 'DernièreModification', 'Actif']);
  if (comptes.getLastRow() < 2) {
    var pwd = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    comptes.appendRow(['admin', sha256Hex_('admin:' + pwd), 'admin', '', '', '', 'Administrateur', '', 'OUI']);
    Logger.log('Compte admin créé. Mot de passe (à noter, affiché une seule fois) : ' + pwd);
  }
  getOrCreateSheet_(SESSIONS_SHEET_NAME, ['Token', 'Identifiant', 'Créée le']);
  getOrCreateSheet_(ENVOIS_SHEET_NAME, ['cid', 'date']);
  codesSheet_();
  ensureQrRoles();
  Logger.log('Comptes, Sessions, QRTokens et Envois prêts.');
}
function ensureQrRoles() {
  var sh = getOrCreateSheet_(QRTOKENS_SHEET_NAME, ['Rôle', 'Token', 'GénéréLe']);
  QR_ROLES.forEach(function (r) { if (!findQrRow_(r)) sh.appendRow([r, Utilities.getUuid(), new Date().toISOString()]); });
}
/** À exécuter UNE FOIS : efface les anciens mots de passe en clair (colonne J de la feuille Comptes). */
function effacerMotsDePasse() {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  if (s && s.getLastRow() > 1 && s.getLastColumn() >= 10) s.getRange(2, 10, s.getLastRow() - 1, 1).clearContent();
}
/** Colonne dont le titre (ligne 1) est exactement « MotDePasseInitial » : index base 0, ou -1 si absente.
 *  Repérée par son titre pour ne jamais prendre une autre colonne (ex. nom d'hôpital) pour un mot de passe. */
function pwdColIndex_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  if (!sh || sh.getLastColumn() < 1) return -1;
  var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  for (var i = 0; i < h.length; i++) if (String(h[i]).trim() === 'MotDePasseInitial') return i;
  return -1;
}
/** Écrit le mot de passe en clair dans la colonne « MotDePasseInitial » (créée si absente), en format texte,
 *  pour qu'il reste visible et modifiable dans la feuille Comptes. */
function setPlainPwd_(row, pwd) {
  var sh = accountsSheet_(), pc = pwdColIndex_();
  if (pc < 0) { pc = Math.max(sh.getLastColumn(), 10); sh.getRange(1, pc + 1).setValue('MotDePasseInitial'); }
  sh.getRange(row, pc + 1).setNumberFormat('@').setValue(String(pwd));
}
var ACCOUNTS_HEADERS = ['Identifiant', 'Hash', 'Niveau', 'Département', 'Zone Sanitaire', 'Commune', 'Libellé', 'DernièreModification', 'Actif', '', 'MotDePasseInitial'];
function findAccountRow_(username) {
  var sh = getOrCreateSheet_(ACCOUNTS_SHEET_NAME, ACCOUNTS_HEADERS);
  var v = sh.getDataRange().getValues();
  username = String(username || '').trim();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === username) return { row: i + 1, data: v[i] };
  return null;
}
function accountPublicView_(d) { return { username: d[0], level: d[2], departement: d[3], zone: d[4], commune: d[5], label: d[6], active: (d[8] === undefined || d[8] === '' || d[8] === 'OUI') }; }
function createSession_(username) {
  var sh = getOrCreateSheet_(SESSIONS_SHEET_NAME, ['Token', 'Identifiant', 'Créée le']), v = sh.getDataRange().getValues(), now = Date.now();
  for (var i = v.length - 1; i >= 1; i--) if (now - new Date(v[i][2]).getTime() > SESSION_DURATION_MS) sh.deleteRow(i + 1);
  var token = Utilities.getUuid();
  sh.appendRow([token, username, new Date().toISOString()]);
  return token;
}
function findSession_(token) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET_NAME);
  if (!sh) return null;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === token) { if (Date.now() - new Date(v[i][2]).getTime() > SESSION_DURATION_MS) return null; return { row: i + 1, username: v[i][1] }; }
  }
  return null;
}
function loginBloque_(u) { return Number(CacheService.getScriptCache().get('f:' + String(u).slice(0, 100)) || 0) >= 5; }
function loginEchec_(u) { var c = CacheService.getScriptCache(), k = 'f:' + String(u).slice(0, 100); c.put(k, String(Number(c.get(k) || 0) + 1), 900); }
function handleAccountLogin(username, password) {
  if (!username || !password) return { ok: false, error: 'Identifiant et mot de passe requis.' };
  if (loginBloque_(username)) return { ok: false, error: "Trop d'essais. Réessayez dans 15 minutes." };
  var f = findAccountRow_(username);
  if (!f) { loginEchec_(username); return { ok: false, error: 'Identifiant ou mot de passe incorrect.' }; }
  if (f.data[8] === 'NON') return { ok: false, error: 'Ce compte a été désactivé par un administrateur.' };
  // Mot de passe saisi directement dans la feuille « Comptes » (colonne K « MotDePasseInitial ») :
  // vérifié à chaque connexion ; la cellule RESTE visible dans la feuille (modifier la cellule = changer le mot de passe).
  // L'empreinte (hash) est tenue à jour en parallèle.
  var pc = pwdColIndex_(), plain = pc < 0 || f.data[pc] === undefined || f.data[pc] === null ? '' : String(f.data[pc]).trim();
  if (plain) {
    if (String(password).trim() !== plain) { loginEchec_(username); return { ok: false, error: 'Identifiant ou mot de passe incorrect.' }; }
    var sh = accountsSheet_();
    sh.getRange(f.row, 2).setValue(sha256Hex_(String(f.data[0]).trim() + ':' + String(password).trim()));
    f.data[0] = String(f.data[0]).trim();
  } else if (sha256Hex_(String(f.data[0]).trim() + ':' + password) !== f.data[1]) { loginEchec_(username); return { ok: false, error: 'Identifiant ou mot de passe incorrect.' }; }
  return { ok: true, token: createSession_(String(f.data[0]).trim()), account: accountPublicView_(f.data) };
}
function requireAdminSession_(token) {
  var s = token ? findSession_(token) : null;
  if (!s) return { ok: false, error: 'Session invalide ou expirée, reconnectez-vous.' };
  var a = findAccountRow_(s.username);
  if (!a || a.data[2] !== 'admin' || a.data[8] === 'NON') return { ok: false, error: 'Réservé au compte administrateur.' };
  return { ok: true, acc: a };
}
function accountsSheet_() { return getOrCreateSheet_(ACCOUNTS_SHEET_NAME, ACCOUNTS_HEADERS); }
function handleListAccounts(token) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var v = accountsSheet_().getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) { var r = v[i]; out.push({ username: r[0], level: r[2], departement: r[3], zone: r[4], commune: r[5], label: r[6], lastModified: r[7] || '', active: (r[8] === undefined || r[8] === '' || r[8] === 'OUI') }); }
  return { ok: true, accounts: out };
}
function handleSetAccountActive(token, username, active) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var f = findAccountRow_(username); if (!f) return { ok: false, error: 'Compte introuvable.' };
  accountsSheet_().getRange(f.row, 9).setValue(active ? 'OUI' : 'NON');
  return { ok: true };
}
function handleUpdateAccount(token, username, fields) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var f = findAccountRow_(username); if (!f) return { ok: false, error: 'Compte introuvable.' };
  var sh = accountsSheet_(); fields = fields || {};
  [['level', 3], ['departement', 4], ['zone', 5], ['commune', 6], ['label', 7]].forEach(function (p) { if (fields[p[0]] !== undefined) sh.getRange(f.row, p[1]).setValue(fields[p[0]]); });
  sh.getRange(f.row, 8).setValue(new Date().toISOString());
  CacheService.getScriptCache().remove('ref');
  return { ok: true };
}
function handleDeleteAccount(token, username) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  if (username === c.acc.data[0]) return { ok: false, error: 'Vous ne pouvez pas supprimer votre propre compte connecté.' };
  var f = findAccountRow_(username); if (!f) return { ok: false, error: 'Compte introuvable.' };
  accountsSheet_().deleteRow(f.row);
  CacheService.getScriptCache().remove('ref');
  return { ok: true };
}
function handleChangePassword(token, oldPassword, newPassword) {
  var s = token ? findSession_(token) : null;
  if (!s) return { ok: false, error: 'Session invalide ou expirée, reconnectez-vous.' };
  if (!newPassword || String(newPassword).length < 8) return { ok: false, error: 'Le nouveau mot de passe doit contenir au moins 8 caractères.' };
  var f = findAccountRow_(s.username); if (!f) return { ok: false, error: 'Compte introuvable.' };
  var _pc = pwdColIndex_(), _pl = _pc < 0 || f.data[_pc] == null ? '' : String(f.data[_pc]).trim();
  if (_pl ? String(oldPassword).trim() !== _pl : sha256Hex_(s.username + ':' + oldPassword) !== f.data[1]) return { ok: false, error: 'Ancien mot de passe incorrect.' };
  var sh = accountsSheet_();
  sh.getRange(f.row, 2).setValue(sha256Hex_(s.username + ':' + newPassword));
  setPlainPwd_(f.row, newPassword);
  sh.getRange(f.row, 8).setValue(new Date().toISOString());
  return { ok: true };
}
function invalidateSessionsForUser_(username) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET_NAME), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (v[i][1] === username) sh.deleteRow(i + 1);
}
function handleChangeUsername(token, currentPassword, newUsername) {
  var s = token ? findSession_(token) : null;
  if (!s) return { ok: false, error: 'Session invalide ou expirée, reconnectez-vous.' };
  var f = findAccountRow_(s.username); if (!f) return { ok: false, error: 'Compte introuvable.' };
  var _pc2 = pwdColIndex_(), _pl2 = _pc2 < 0 || f.data[_pc2] == null ? '' : String(f.data[_pc2]).trim();
  if (_pl2 ? String(currentPassword).trim() !== _pl2 : sha256Hex_(s.username + ':' + currentPassword) !== f.data[1]) return { ok: false, error: 'Mot de passe incorrect.' };
  newUsername = (newUsername || '').toString().trim();
  if (!newUsername) return { ok: false, error: 'Identifiant vide.' };
  if (newUsername === s.username) return { ok: false, error: "C'est déjà votre identifiant actuel." };
  if (findAccountRow_(newUsername)) return { ok: false, error: 'Cet identifiant est déjà utilisé par un autre compte.' };
  var sh = accountsSheet_();
  sh.getRange(f.row, 1).setValue(newUsername);
  sh.getRange(f.row, 2).setValue(sha256Hex_(newUsername + ':' + currentPassword));
  setPlainPwd_(f.row, currentPassword);
  sh.getRange(f.row, 8).setValue(new Date().toISOString());
  invalidateSessionsForUser_(s.username);
  var d = f.data.slice(); d[0] = newUsername;
  return { ok: true, token: createSession_(newUsername), account: accountPublicView_(d) };
}

/* ---------- Codes QR enquêteurs ---------- */
function findQrRow_(role) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(QRTOKENS_SHEET_NAME);
  if (!sh) return null;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === role) return { row: i + 1, token: v[i][1] };
  return null;
}
function qrTokenValid_(token, roles) {
  token = (token || '').toString().trim();
  if (!token) return false;
  var levels = {};
  for (var i = 0; i < roles.length; i++) {
    var r = findQrRow_(roles[i]); if (r && String(r.token).trim() === token) return true;
    if (roles[i] === 'enqueteur') levels.zone = 1; else if (roles[i] === 'enqueteur_hopital') levels.hopital = 1;
  }
  if (accessCodesRows_().some(function (c) { return levels[c.level] && c.active && c.code === token; })) return true;
  var a = activityByCode_(token);
  return !!(a && levels[a.level]);
}
function handleGetQrToken(sessionToken, role) {
  var c = requireAdminSession_(sessionToken); if (!c.ok) return c;
  if (QR_ROLES.indexOf(role) === -1) return { ok: false, error: 'Rôle inconnu.' };
  var r = findQrRow_(role);
  if (r) return { ok: true, token: r.token };
  var t = Utilities.getUuid();
  getOrCreateSheet_(QRTOKENS_SHEET_NAME, ['Rôle', 'Token', 'GénéréLe']).appendRow([role, t, new Date().toISOString()]);
  return { ok: true, token: t };
}
function setQr_(role, code) {
  var sh = getOrCreateSheet_(QRTOKENS_SHEET_NAME, ['Rôle', 'Token', 'GénéréLe']), r = findQrRow_(role);
  if (r) { sh.getRange(r.row, 2).setNumberFormat('@').setValue(code); sh.getRange(r.row, 3).setValue(new Date().toISOString()); }
  else { sh.appendRow([role, code, new Date().toISOString()]); sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(code); }
}
function handleRefreshQrToken(sessionToken, role) {
  var c = requireAdminSession_(sessionToken); if (!c.ok) return c;
  if (QR_ROLES.indexOf(role) === -1) return { ok: false, error: 'Rôle inconnu.' };
  var t = Utilities.getUuid(); setQr_(role, t);
  return { ok: true, token: t };
}
function handleSetQrCode(sessionToken, role, code) {
  var c = requireAdminSession_(sessionToken); if (!c.ok) return c;
  if (QR_ROLES.indexOf(role) === -1) return { ok: false, error: 'Rôle inconnu.' };
  code = (code || '').toString().trim();
  if (!code) return { ok: false, error: 'Code vide ou non reconnu.' };
  setQr_(role, code);
  return { ok: true, token: code };
}
function handleMatchQrCode(rawText) {
  rawText = (rawText || '').toString().trim();
  if (!rawText) return { ok: false, error: 'QR code vide.' };
  var cand = rawText, q = rawText.indexOf('?');
  if (rawText.indexOf('http') === 0 && q !== -1) {
    rawText.substring(q + 1).split('#')[0].split('&').forEach(function (p) {
      var kv = p.split('=');
      if (kv[0] === 'code' || kv[0] === 't') cand = decodeURIComponent(kv.slice(1).join('='));
    });
  }
  cand = String(cand).trim();
  for (var i = 0; i < QR_ROLES.length; i++) {
    var r = findQrRow_(QR_ROLES[i]);
    if (r && String(r.token).trim() === cand) {
      var lv = QR_ROLES[i] === 'enqueteur_hopital' ? 'hopital' : QR_ROLES[i] === 'enqueteur' ? 'zone' : '';
      return { ok: true, role: QR_ROLES[i], level: lv, label: lv === 'hopital' ? 'Tous les hôpitaux' : lv === 'zone' ? 'Toutes les zones sanitaires' : '' };
    }
  }
  var rows = accessCodesRows_();
  for (var j = 0; j < rows.length; j++) {
    if (rows[j].code === cand && rows[j].active) return { ok: true, role: rows[j].level === 'hopital' ? 'enqueteur_hopital' : 'enqueteur', level: rows[j].level, label: rows[j].name };
  }
  var acts = activitiesRows_(), up = cand.toUpperCase(), today = day_(new Date());
  for (var k = 0; k < acts.length; k++) {
    var a = acts[k]; if (a.code.toUpperCase() !== up) continue;
    if (!a.active) return { ok: false, error: "Cette activité est désactivée. Demandez à l'administrateur." };
    if (a.debut && today < a.debut) return { ok: false, error: "Cette activité commence le " + a.debut + "." };
    if (a.fin && today > a.fin) return { ok: false, error: "Cette activité est terminée (fin le " + a.fin + ")." };
    return { ok: true, role: a.level === 'hopital' ? 'enqueteur_hopital' : 'enqueteur', level: a.level, label: a.name + (a.cible ? ' — ' + a.cible : ''), activity: activityPublic_(a) };
  }
  return { ok: false, error: "Ce code ne correspond à aucune cible active." };
}

/* ---------- Codes d'accès par zone sanitaire / par hôpital ---------- */
var CODES_SHEET_NAME = 'CodesAcces', CODES_HEADERS = ['Niveau', 'Nom', 'Code', 'GénéréLe', 'Actif'];
function codesSheet_() { return getOrCreateSheet_(CODES_SHEET_NAME, CODES_HEADERS); }
function accessCodesRows_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CODES_SHEET_NAME);
  if (!sh) return [];
  var v = sh.getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) if (v[i][0] && v[i][2] !== '') out.push({ row: i + 1, level: String(v[i][0]), name: String(v[i][1]), code: String(v[i][2]).trim(), active: v[i][4] !== 'NON' });
  return out;
}
function levelRole_(level) { return level === 'hopital' ? 'enqueteur_hopital' : 'enqueteur'; }
function codeEnUsage_(code, exceptRow, exceptRole, exceptActRow) {
  code = String(code).trim();
  if (activitiesRows_().some(function (a) { return a.code.toUpperCase() === code.toUpperCase() && a.row !== exceptActRow; })) return true;
  var general = QR_ROLES.some(function (r) { var q = findQrRow_(r); return r !== exceptRole && q && String(q.token).trim() === code; });
  return general || accessCodesRows_().some(function (r) { return r.code === code && r.row !== exceptRow; });
}
function handleListAccessCodes(token) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  ensureQrRoles(); codesSheet_();
  var items = [];
  [['zone', 'enqueteur'], ['hopital', 'enqueteur_hopital']].forEach(function (p) {
    var r = findQrRow_(p[1]); items.push({ level: p[0], name: '', general: true, code: String(r.token), active: true });
  });
  accessCodesRows_().forEach(function (r) { items.push({ level: r.level, name: r.name, general: false, code: r.code, active: r.active }); });
  var zones = {}, sh = accountsSheet_();
  if (sh) { var v = sh.getDataRange().getValues(); for (var i = 1; i < v.length; i++) if (v[i][4]) zones[String(v[i][4])] = 1; }
  return { ok: true, codes: items, zones: Object.keys(zones).sort() };
}
/** Crée ou modifie un code. name vide = code général du niveau. code vide = code généré automatiquement. */
function handleSaveAccessCode(token, level, name, code) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  if (level !== 'zone' && level !== 'hopital') return { ok: false, error: 'Niveau inconnu.' };
  name = String(name || '').trim().slice(0, 80); code = String(code || '').trim();
  if (code && code.length < 4) return { ok: false, error: 'Code trop court (4 caractères minimum).' };
  if (!name) {
    var role = levelRole_(level);
    if (!code) code = Utilities.getUuid();
    if (codeEnUsage_(code, 0, role)) return { ok: false, error: 'Ce code est déjà utilisé ailleurs.' };
    setQr_(role, code); return { ok: true, code: code };
  }
  var ex = null;
  accessCodesRows_().forEach(function (r) { if (r.level === level && r.name.toLowerCase() === name.toLowerCase()) ex = r; });
  if (!code) { do { code = Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase(); } while (codeEnUsage_(code, ex ? ex.row : 0, '')); }
  else if (codeEnUsage_(code, ex ? ex.row : 0, '')) return { ok: false, error: 'Ce code est déjà utilisé ailleurs.' };
  var sh = codesSheet_(), row = ex ? ex.row : sh.getLastRow() + 1, now = new Date().toISOString();
  sh.getRange(row, 1, 1, 5).setNumberFormat('@').setValues([[level, ex ? ex.name : name, code, now, ex ? (ex.active ? 'OUI' : 'NON') : 'OUI']]);
  return { ok: true, code: code };
}
function findSpecificCode_(level, name) {
  var f = null; name = String(name || '').trim().toLowerCase();
  accessCodesRows_().forEach(function (r) { if (r.level === level && r.name.toLowerCase() === name) f = r; });
  return f;
}
function handleSetAccessCodeActive(token, level, name, active) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var f = findSpecificCode_(level, name); if (!f) return { ok: false, error: 'Code introuvable.' };
  codesSheet_().getRange(f.row, 5).setNumberFormat('@').setValue(active ? 'OUI' : 'NON');
  return { ok: true };
}
function handleDeleteAccessCode(token, level, name) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var f = findSpecificCode_(level, name); if (!f) return { ok: false, error: 'Code introuvable.' };
  codesSheet_().deleteRow(f.row);
  return { ok: true };
}

/* ---------- Activités : un code QR par activité ---------- */
var AUDIO_COL = 'Audio', AUDIO_SHEET_NAME = '🎙️ Audios';
var ACT_SHEET_NAME = 'Activites', ACT_COL = 'Activité', ACT_ID_COL = 'IdActivité';
var ACT_HEADERS = ['ID', 'Activité', 'Niveau', 'Cible', 'Formulaires', 'Communes', 'Début', 'Fin', 'Code', 'Actif', 'Créée le'];
var ACT_FORMS = { zone: ['usagers', 'cogecs'], hopital: ['usagers', 'communaute'] };
var FORM_KEY_ = { submitUsagers: 'usagers', submitCogecs: 'cogecs', submitUsagersHopital: 'usagers', submitCommunauteHopital: 'communaute' };
var ACT_GRACE_DAYS = 1; // les réponses d'un enquêteur hors connexion restent acceptées 1 jour après la fin
function actSheet_() { return getOrCreateSheet_(ACT_SHEET_NAME, ACT_HEADERS); }
function csvList_(v) { return String(v || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean); }
function activitiesRows_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACT_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0] || v[i][8] === '') continue;
    out.push({ row: i + 1, id: String(v[i][0]), name: String(v[i][1]), level: String(v[i][2]), cible: String(v[i][3] || ''), forms: csvList_(v[i][4]), communes: [], debut: day_(v[i][6]), fin: day_(v[i][7]), code: String(v[i][8]).trim(), active: v[i][9] !== 'NON' });
  }
  return out;
}
function dayPlus_(s, n) {
  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/); if (!m) return s;
  return Utilities.formatDate(new Date(+m[1], +m[2] - 1, +m[3] + n, 12, 0, 0), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
/** Activité active correspondant au code, dans sa période (la fin est prolongée de ACT_GRACE_DAYS jour pour l'envoi). */
function activityByCode_(code) {
  code = String(code || '').trim().toUpperCase(); if (!code) return null;
  var today = day_(new Date()), f = null;
  activitiesRows_().forEach(function (a) {
    if (a.code.toUpperCase() !== code || !a.active) return;
    if (a.debut && today < a.debut) return;
    if (a.fin && today > dayPlus_(a.fin, ACT_GRACE_DAYS)) return;
    f = a;
  });
  return f;
}
function activityPublic_(a) { return { id: a.id, name: a.name, level: a.level, cible: a.cible, forms: a.forms, communes: a.communes, debut: a.debut, fin: a.fin }; }
function handleListActivities(token) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  actSheet_();
  var zones = {}, sh = accountsSheet_();
  if (sh) { var v = sh.getDataRange().getValues(); for (var i = 1; i < v.length; i++) if (v[i][4]) zones[String(v[i][4])] = 1; }
  return { ok: true, activities: activitiesRows_().map(function (a) { var p = activityPublic_(a); p.code = a.code; p.active = a.active; return p; }), zones: Object.keys(zones).sort() };
}
/** Crée (sans id) ou modifie (avec id) une activité. renew = true : nouveau code. */
function handleSaveActivity(token, p) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  p = p || {};
  var name = String(p.name || '').trim().slice(0, 80), level = p.level, cible = String(p.cible || '').trim().slice(0, 80);
  if (!name) return { ok: false, error: "Indiquez le nom de l'activité." };
  if (level !== 'zone' && level !== 'hopital') return { ok: false, error: 'Niveau inconnu.' };
  var okForms = ACT_FORMS[level], forms = (p.forms || []).filter(function (f) { return okForms.indexOf(f) > -1; });
  var communes = []; // les activités ne sont plus restreintes à des communes
  var rx = /^\d{4}-\d{2}-\d{2}$/, debut = String(p.debut || ''), fin = String(p.fin || '');
  if ((debut && !rx.test(debut)) || (fin && !rx.test(fin))) return { ok: false, error: 'Date invalide.' };
  if (debut && fin && fin < debut) return { ok: false, error: 'La date de fin précède la date de début.' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ex = null; activitiesRows_().forEach(function (a) { if (p.id && a.id === p.id) ex = a; });
    if (p.id && !ex) return { ok: false, error: 'Activité introuvable.' };
    var code = ex ? ex.code : '';
    if (!code || p.renew) { do { code = Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase(); } while (codeEnUsage_(code, 0, '', ex ? ex.row : 0)); }
    var sh = actSheet_(), row = ex ? ex.row : sh.getLastRow() + 1;
    var active = ex ? (ex.active ? 'OUI' : 'NON') : 'OUI';
    sh.getRange(row, 1, 1, ACT_HEADERS.length).setNumberFormat('@').setValues([[ex ? ex.id : 'A' + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase(), name, level, cible, forms.join(','), communes.join(','), debut, fin, code, active, new Date().toISOString()]]);
    return { ok: true, code: code };
  } finally { lock.releaseLock(); }
}
function handleSetActivityActive(token, id, active) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var f = null; activitiesRows_().forEach(function (a) { if (a.id === id) f = a; });
  if (!f) return { ok: false, error: 'Activité introuvable.' };
  actSheet_().getRange(f.row, 10).setNumberFormat('@').setValue(active ? 'OUI' : 'NON');
  return { ok: true };
}
function handleDeleteActivity(token, id) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var f = null; activitiesRows_().forEach(function (a) { if (a.id === id) f = a; });
  if (!f) return { ok: false, error: 'Activité introuvable.' };
  actSheet_().deleteRow(f.row);
  return { ok: true };
}

/* ---------- Filtrage serveur selon la portée du compte ---------- */
function referentiel_() {
  var cache = CacheService.getScriptCache(), k = cache.get('ref');
  if (k) return JSON.parse(k);
  var v = accountsSheet_().getDataRange().getValues(), r = {};
  for (var i = 1; i < v.length; i++) if (v[i][2] === 'commune' && v[i][5]) r[v[i][5]] = { departement: v[i][3], zone: v[i][4] };
  cache.put('ref', JSON.stringify(r), 600);
  return r;
}
var STRUCTURE_COMMUNE = {"CS Agbokpa": "Abomey", "CS Detohou": "Abomey", "CS Tangandji": "Abomey", "CS Djegbe (Abomey)": "Abomey", "CS Hounli": "Abomey", "CS Sehoun": "Abomey", "CS Adandokpodji": "Abomey", "CS Vidole": "Abomey", "CS Zounzonme": "Abomey", "CS Adanhondjigon": "Agbangnizoun", "CS Adingnigon": "Agbangnizoun", "CS Agbangnizoun": "Agbangnizoun", "CS Kinta": "Agbangnizoun", "CS Akodébakou": "Agbangnizoun", "MI Sinwe-Kpota": "Agbangnizoun", "CS Lissazounme": "Agbangnizoun", "CS Dovota": "Agbangnizoun", "CS Sahè": "Agbangnizoun", "CS Sinwe": "Agbangnizoun", "CS Tanve": "Agbangnizoun", "CS Kpoto": "Agbangnizoun", "CS Agondji": "Djidja", "CS Agouna": "Djidja", "CS Dan": "Djidja", "CS Djidja": "Djidja", "CS Dohouime": "Djidja", "CS Gobaix": "Djidja", "CS Outo": "Djidja", "CS Lobeta": "Djidja", "CS Monsourou": "Djidja", "CS Mougnon": "Djidja", "CS Oumgbega": "Djidja", "CS Saloudji": "Djidja", "CS Setto": "Djidja", "CS Ayogbé": "Djidja", "CS Sodohomè": "Bohicon", "CS Saclo": "Bohicon", "CS Gnidjazoun": "Bohicon", "CS Bohicon 2": "Bohicon", "Cs Lissèzoun": "Bohicon", "CS Ouassaho": "Bohicon", "CS Bohicon 1": "Bohicon", "CS Passagon": "Bohicon", "CS Avogbanna": "Bohicon", "CS Agongointo": "Bohicon", "CS Adjinagon": "Za-Kpota", "CS Za-kpota": "Za-Kpota", "CS Allahé": "Za-Kpota", "MI Za-Hla": "Za-Kpota", "CS Za-Tanta": "Za-Kpota", "CS Kpakpamè": "Za-Kpota", "CS Kpozoun": "Za-Kpota", "CS Houngomè": "Za-Kpota", "CS Assanlin": "Za-Kpota", "CS kpolokoué": "Za-Kpota", "CS zounzounmè": "Za-Kpota", "CS Tindji": "Za-Kpota", "CS Deme": "Zogbodomey", "CS Koussoukpa": "Zogbodomey", "CS Kpokissa": "Zogbodomey", "CS Massi": "Zogbodomey", "CS Cana 2": "Zogbodomey", "CS Zogbodome": "Zogbodomey", "MI Zoungbo-Bogon": "Zogbodomey", "CS Tanwé hessou": "Zogbodomey", "CS Akiza": "Zogbodomey", "CS Avlamè": "Zogbodomey", "CS Domè": "Zogbodomey", "CS Zoukou": "Zogbodomey", "CS Houin-Hounso": "Covè", "CS Covè": "Covè", "CS Lainta Cogbé": "Covè", "CS Naogon": "Covè", "CS Agonlin-Houegbo": "Zagnanado", "CS Baname": "Zagnanado", "CS Zakoumado": "Zagnanado", "MI Gossoé": "Zagnanado", "CS Don-Tan": "Zagnanado", "CS Dovi-Dove": "Zagnanado", "MI Dovi Zounnou": "Zagnanado", "CS Kpedekpo": "Zagnanado", "CS Zagnannado": "Zagnanado", "CS Dasso": "Ouinhi", "CS Ouinhi": "Ouinhi", "CS Aïzè": "Ouinhi", "CS Sagon": "Ouinhi", "MI Houédja": "Ouinhi", "MI Tévêdji": "Ouinhi", "CS Tohoues": "Ouinhi", "MI Akassa": "Ouinhi"};
function communeInScope_(label, account) {
  var lvl = account.level;
  if (lvl === 'admin' || lvl === 'national' || lvl === 'superviseur') return true;
  if (lvl === 'commune') return label === account.commune;
  var info = referentiel_()[label];
  if (!info) return false;
  if (lvl === 'departement') return info.departement === account.departement;
  if (lvl === 'zone') return info.zone === account.zone;
  return false;
}
function structureInScope_(label, account) { var c = STRUCTURE_COMMUNE[label]; return c ? communeInScope_(c, account) : false; }
function recomputeTotal_(rows) {
  var count = 0, w = 0, wn = 0;
  rows.forEach(function (r) { count += r.count; if (typeof r.avgScore === 'number') { w += r.avgScore * r.count; wn += r.count; } });
  return { count: count, avgScore: wn > 0 ? Math.min(100, Math.round((w / wn) * 10) / 10) : null };
}
function applyScopeToResult_(result, account) {
  var lvl = account.level;
  if (lvl === 'admin' || lvl === 'national' || lvl === 'superviseur') return result;
  var cibles = ['usagers', 'personnels', 'cogecs'];
  cibles.forEach(function (c) {
    if (result[c].byCommune) result[c].byCommune = result[c].byCommune.filter(function (r) { return communeInScope_(r.label, account); });
    if (result[c].byStructure) result[c].byStructure = result[c].byStructure.filter(function (r) { return structureInScope_(r.label, account); });
    if (result[c].byEnqueteur) result[c].byEnqueteur = [];
    var cr = result.communeReport[c]; Object.keys(cr.byCommune || {}).forEach(function (k) { if (!communeInScope_(k, account)) delete cr.byCommune[k]; });
    var sr = result.structureReport[c]; Object.keys(sr.byStructure || {}).forEach(function (k) { if (!structureInScope_(k, account)) delete sr.byStructure[k]; });
  });
  result.global.byCommune = result.global.byCommune.filter(function (r) { return communeInScope_(r.label, account); });
  var tU = recomputeTotal_(result.usagers.byCommune), tP = recomputeTotal_(result.personnels.byCommune), tC = recomputeTotal_(result.cogecs.byCommune);
  result.global.byCible = [{ label: "Usagers", count: tU.count, avgScore: tU.avgScore }, { label: "Personnel soignant", count: tP.count, avgScore: tP.avgScore }, { label: "COGECS", count: tC.count, avgScore: tC.avgScore }];
  result.counts = { usagers: tU.count, personnels: tP.count, cogecs: tC.count, total: tU.count + tP.count + tC.count };
  result.enqueteurReport = { usagers: { categories: [], byEnqueteur: {} }, cogecs: { categories: [], byEnqueteur: {} } };
  Object.keys(result.gpsByDay).forEach(function (day) { result.gpsByDay[day] = result.gpsByDay[day].filter(function (p) { return p.commune && communeInScope_(p.commune, account); }); });
  return result;
}

/** À exécuter si vous ne connaissez pas (ou plus) le mot de passe du compte « admin » :
 *  crée un nouveau mot de passe, affiché UNE SEULE FOIS dans Affichage > Journaux. */
function reinitialiserMotDePasseAdmin() {
  var pwd = Utilities.getUuid().replace(/-/g, '').slice(0, 10), f = findAccountRow_('admin');
  var sh = getOrCreateSheet_(ACCOUNTS_SHEET_NAME, ['Identifiant', 'Hash', 'Niveau', 'Département', 'Zone Sanitaire', 'Commune', 'Libellé', 'DernièreModification', 'Actif']);
  if (!f) { sh.appendRow(['admin', sha256Hex_('admin:' + pwd), 'admin', '', '', '', 'Administrateur', '', 'OUI']); setPlainPwd_(sh.getLastRow(), pwd); }
  else { sh.getRange(f.row, 2).setValue(sha256Hex_('admin:' + pwd)); sh.getRange(f.row, 3).setValue('admin'); sh.getRange(f.row, 9).setValue('OUI'); setPlainPwd_(f.row, pwd); }
  Logger.log('Identifiant : admin — Nouveau mot de passe : ' + pwd);
}

/* ====================================================================
 * PÉRIODES DE COLLECTE, DONNÉES HÔPITAUX, ÉVOLUTION
 * ==================================================================== */
var PERIODS_SHEET_NAME = 'Periodes', PERIODS_HEADERS = ['Id', 'Libellé', 'Début', 'Fin', 'Actif'];
var DATE_COLUMN = 'DateEnvoi', START_COL = 'Début saisie', END_COL = 'Fin saisie', DEV_COL = 'Appareil'; // heures de début et de fin du questionnaire (ISO), utilisées pour valider
var HOSP_CFG = {"usagers":{"sheet":SHEET_USAGERS_H,"hosp":"fs","zone":"zone","dept":"dept","cats":[{"n":"Accessibilité et orientation","items":[{"c":"Comment_avez_vous_tr_acc_der_l_h_pital_","l":"yh7ka41"},{"c":"Avez_vous_rencontr_re_v_hicule_ou_moto_","l":"vv2mg93"},{"c":"L_acc_s_aux_b_timent_mobilit_r_duite_","l":"do9ex82"},{"c":"Une_fois_l_int_rie_neaux_d_orientation_","l":"fx3yw25"},{"c":"La_propret_et_le_co_ils_au_rendez_vous_","l":"lr9zr25"},{"c":"Avez_vous_des_diffic_au_sein_de_l_h_pital","l":"tt7uf73","i":1}]},{"n":"Accueil et temps d'attente","items":[{"c":"A1","l":"tt3jk83"},{"c":"A2","l":"xc4ty77"},{"c":"A3","l":"qg8km11"},{"c":"A3a","l":"xc4ty77"},{"c":"A4","l":"qg8km11"},{"c":"A4a","l":"xc4ty77"},{"c":"A5","l":"or5hc40"},{"c":"A5a","l":"xc4ty77"},{"c":"A6","l":"qg8km11"},{"c":"A6a","l":"xc4ty77"}]},{"n":"Consultation, écoute et soins","items":[{"c":"C1","l":"lr1fn01"},{"c":"C2","l":"aq4ny65"},{"c":"D1","l":"et8pq06"},{"c":"D2","l":"xe8kz81","i":1},{"c":"E1","l":"xe8kz81"},{"c":"B1","l":"ft6fb55"},{"c":"B2","l":"xa3sf70"},{"c":"B3","l":"hj0ae68"},{"c":"E2","l":"nb71g36"},{"c":"E3","l":"lr1fn01"},{"c":"E4","l":"zj4kh59"},{"c":"A7","l":"hn23q50"},{"c":"A8","l":"el7ir34","i":1},{"c":"C4","l":"qm6xj64"},{"c":"C5","l":"to23d11"}]},{"n":"Médicaments et coûts","items":[{"c":"G1","l":"fr7fs28"},{"c":"G2","l":"fr7fs28"},{"c":"F1","l":"vj2kq12"},{"c":"F2","l":"vj2kq12"},{"c":"F3","l":"vj2kq12"},{"c":"C3","l":"yg8fo88"}]},{"n":"Droits du patient et plaintes","items":[{"c":"Avant_qu_une_d_cisio_vous_t_consult_e_","l":"dm08z48"},{"c":"Avez_vous_pu_exprime_les_soins_propos_s_","l":"oi06h09"},{"c":"Si_oui_votre_d_cisi_ression_ni_jugement_","l":"ac7oc61"},{"c":"Lors_de_votre_prise_en_charge_","l":"ft8yj05"},{"c":"Si_oui_cette_person_essionnels_de_sant_","l":"oe7lp13"},{"c":"Lors_de_votre_prise_en_charge","l":"oy6qx24"},{"c":"Lors_de_votre_prise_en_charge__001","l":"vn2fi03"}]},{"n":"Hygiène et conditions de séjour","items":[{"c":"H1","l":"ky86z20"},{"c":"H2","l":"wc9ng48"},{"c":"H3","l":"bm4ts82"},{"c":"H4","l":"zo1if20"},{"c":"H5","l":"my43j02"},{"c":"H6","l":"to4qf61"},{"c":"H7","l":"tg9gh34"},{"c":"C6","l":"bm2eu92"},{"c":"H8","l":"kx3na73"}]},{"n":"Pratiques parallèles et rançonnement","items":[{"c":"AA1","l":"si4fc54","i":1},{"c":"AA2","l":"st7nh97","i":1},{"c":"AA3","l":"uf3af06","i":1},{"c":"AA4","l":"nd9id61"}]},{"n":"Fidélité et recommandation","items":[{"c":"I1","l":"di0gd12"},{"c":"I2","l":"co65e31"}]}]},"personnel":{"sheet":SHEET_PERSONNELS_H,"hosp":"fs","zone":"zs","dept":"dept","cats":[{"n":"Conditions et organisation du travail","items":[{"c":"Q1","l":"dt1jh61"},{"c":"Q2","l":"au1wl40"},{"c":"Q4","l":"lc4jc27"},{"c":"Q5","l":"lz1xh46"},{"c":"Q6","l":"my0pl21"},{"c":"Q7","l":"ht7xl14"},{"c":"Q8","l":"cn08k85"},{"c":"Recevez_vous_les_inf_d_lais_raisonnables_","l":"ye6ms95"},{"c":"Q15","l":"sx8na20"},{"c":"Q17","l":"zo7fe17"},{"c":"Q18","l":"ui7si32"},{"c":"Q19","l":"fg4lo40"}]},{"n":"Droits et devoirs des patients","items":[{"c":"Avez_vous_re_u_une_f_evoirs_des_patients_","l":"pb2zc32"},{"c":"Vous_sentez_vous_sou_evoirs_des_patients_","l":"ep8ob11"},{"c":"Avez_vous_connaissance_des_dro","l":"fs59c72"},{"c":"Ces_droits_et_devoir_votre_tablissement_","l":"kf85b90"},{"c":"Pensez_vous_que_vos_onnement_de_travail_","l":"vt7ev88"},{"c":"Avez_vous_des_difficult_s_dans","l":"ps1pl37","i":1}]},{"n":"Collègues et collaboration","items":[{"c":"Q9","l":"ll2ik06"},{"c":"Les_coll_gues_de_vot_protocoles_de_soins_","l":"yh4vk06"},{"c":"Les_coll_gues_assure_esoins_des_patients_","l":"wu8sg46"},{"c":"Les_coll_gues_collab_ontinuit_des_soins_","l":"ye5ft78"},{"c":"Q10","l":"dn1qa13","i":1}]},{"n":"Direction, gestion et perspectives","items":[{"c":"Q11","l":"ke4iv96"},{"c":"Q12","l":"dq1ei30"},{"c":"Q13","l":"km19b24"},{"c":"Q14","l":"aq3kk41"},{"c":"Q20","l":"ie4cp00"},{"c":"Q22","l":"ak5xg12"}]},{"n":"Climat et motivation","items":[{"c":"Q3","l":"fj5iq62"},{"c":"Q16","l":"eg9su19"},{"c":"Q21","l":"nv6jq01"}]}]},"communaute":{"sheet":SHEET_COMMUNAUTE_H,"hosp":"Formation_Sanitaire","zone":"zone","dept":"dept","cats":[{"n":"Accessibilité de l'hôpital","items":[{"c":"Comment_avez_vous_tr_acc_der_l_h_pital_","l":"yh7ka41"},{"c":"Avez_vous_rencontr_re_v_hicule_ou_moto_","l":"vv2mg93"},{"c":"L_acc_s_aux_b_timent_mobilit_r_duite_","l":"do9ex82"},{"c":"Une_fois_l_int_rie_neaux_d_orientation_","l":"fx3yw25"},{"c":"La_propret_et_le_co_ils_au_rendez_vous_","l":"lr9zr25"},{"c":"Avez_vous_des_suggestions_part","l":"my17e38"}]},{"n":"Accueil","items":[{"c":"Comment_jugez_vous_e_accueil_au_portail_","l":"gc9bj55"},{"c":"Comment_jugez_vous_e_l_dans_les_services_","l":"rr4vi05"},{"c":"Comment_jugez_vous_e_au_service_de_sortie","l":"kr3pn80"},{"c":"Comment_trouvez_vous_les_dirigeants_","l":"yl4bj31"}]},{"n":"Prise en charge","items":[{"c":"Comment_appr_ciez_vo_es_agents_examiner","l":"kp2xf35"},{"c":"Comment_appr_ciez_vo_diguer_de_bons_soins","l":"he7qz74"},{"c":"Comment_appr_ciez_vo_e_bonne_prescription","l":"ll4vy48"},{"c":"Comment_appr_ciez_vo_niquer_sur_les_actes","l":"yw8ak40"},{"c":"Pensez_vous_que_les_la_prise_en_charge_","l":"hz7sr16"},{"c":"Comment_trouvez_vous_rapport_avec_le_mal","l":"dg0fr95"},{"c":"Pensez_vous_que_les_ous_servir_au_besoin","l":"go5va85"},{"c":"Les_m_dicaments_pres_ibles_dans_le_centre","l":"lu3ug49"},{"c":"Les_m_dicaments_non_pharmacies_proches_","l":"pc5hv66"},{"c":"Comment_trouvez_vous_ifs_des_m_dicaments_","l":"pr5ud52"},{"c":"Comment_trouvez_vous_ents_hospitalisation","l":"ga1tt79"}]},{"n":"Aisance de séjour","items":[{"c":"Comment_trouvez_vous_la_cour_de_l_h_pital","l":"qv3pt77"},{"c":"Comment_trouvez_vous_alle_de_consultation","l":"uw5ur14"},{"c":"Comment_trouvez_vous_les_toilettes","l":"qf8kr56"},{"c":"Comment_trouvez_vous_la_buanderie","l":"sn14p46"},{"c":"Pensez_vous_que_les_onn_es_sur_les_soins","l":"pw6gg98"},{"c":"Pensez_vous_qu_on_a_est_pas_souvent_seul","l":"wx0vw58"},{"c":"Avez_vous_acc_s_l_eau","l":"ca3ve40"}]},{"n":"Pratiques parallèles","items":[{"c":"Vous_a_t_on_une_fois_des_soins_domicile","l":"tw9po34"},{"c":"Vous_est_il_jamais_a_r_des_agents_de_sant","l":"em0ev27"},{"c":"Des_agents_vous_ont_eux_directement_vers","l":"hg9ab90"}]},{"n":"Gestion administrative de l’hôpital","items":[{"c":"Pensez_vous_que_les_h_pital_g_rent_bien","l":"rg7fs81"},{"c":"Pensez_vous_que_dans_u_sein_de_l_h_pital_","l":"ny6xw05"}]}]}};
var HOSP_SCALES = {"yh7ka41":{"Très claire":1,"Passable":0.33,"Insuffisante":0.33,"Inexistante":0},"vv2mg93":{"Aucune difficulté":1,"Quelques difficultés (recherche longue)":0.5,"Majeures (manque de places ou parking inadapté)":0},"do9ex82":{"Tout à fait adapté":1,"Plutôt adapté":0.67,"Peu adapté":0.33,"Pas du tout adapté":0},"fx3yw25":{"Très facilement":1,"Assez facilement":0.67,"Difficilement":0.33,"Je me suis perdu(e)":0},"lr9zr25":{"Très satisfaisant":1,"Satisfaisant":0.67,"Peu satisfaisant":0.33,"Pas du tout satisfaisant":0},"tt7uf73":{"Oui":1,"Non":0},"tt3jk83":{"Excellent":1,"Bon":0.67,"Passable":0.33,"Mauvais":0},"xc4ty77":{"Long":0.4,"Très long":0,"Acceptable":1},"qg8km11":{"Excellent":1,"Bon":0.67,"Passable":0.33,"Mauvais":0},"or5hc40":{"Excellent":1,"Bon":0.67,"Passable":0.33,"Mauvais":0},"lr1fn01":{"Très bien":1,"Bien":0.67,"Acceptable":0.5,"Mauvais":0},"aq4ny65":{"Oui, très bien":1,"Oui, bien":0.75,"Oui, passablement":0.4,"Non":0},"et8pq06":{"Correctement":1,"Passablement":0.5,"Non":0},"xe8kz81":{"Oui":1,"Non":0},"ft6fb55":{"Oui, tout ce qu'il faut":1,"Oui, un peu":0.5,"Non, rien du tout":0},"xa3sf70":{"Toujours":1,"Très souvent":0.75,"Quelques fois":0.5,"Rarement":0.25,"Jamais":0},"hj0ae68":{"Toujours":1,"Très souvent":0.75,"Quelques fois":0.5,"Rarement":0.25,"Jamais":0},"nb71g36":{"Très compétent":1,"Compétent":0.75,"Pas du tout compétent":0},"zj4kh59":{"Très satisfaisant":1,"Satisfaisant":0.67,"Peu satisfaisant":0.33,"Pas du tout satisfaisant":0},"hn23q50":{"Nous dépasse comme au marché":0,"Amicaux et serviables":1,"Ne nous considère pas":0,"Cherchent à nous encourager":1},"el7ir34":{"Très souvent":0.75,"Assez souvent":0.5,"Parfois":0.5,"Très rarement":0.1,"Jamais":0},"qm6xj64":{"Oui, promptement":1,"Oui, mais négligemment":0.4,"Non, pas du tout":0},"to23d11":{"Oui, toujours":1,"Oui, parfois":0.5,"Non, presque jamais":0},"fr7fs28":{"Oui":1,"Non":0,"Pas toujours":0.5},"vj2kq12":{"Très chers":0,"Chers":0.33,"Abordables":1},"yg8fo88":{"Oui, tout à faire":1,"En partie":0.5,"Non, pas du tout":0},"dm08z48":{"Oui":1,"Non":0},"oi06h09":{"Oui":1,"Non":0},"ac7oc61":{"Oui":1,"Non":0},"ft8yj05":{"Oui":1,"Non":0},"oe7lp13":{"Oui":1,"Non":0},"oy6qx24":{"Oui":1,"Non":0},"vn2fi03":{"Oui":1,"Non":0},"ky86z20":{"Convenable":1,"Acceptable":0.5,"Pas du tout convenable":0},"wc9ng48":{"Oui":1,"Non":0},"bm4ts82":{"Très convenable":1,"Acceptable":0.5,"Pas du tout convenable":0},"zo1if20":{"Très convenable":1,"Acceptable":0.5,"Pas du tout convenable":0},"my43j02":{"Très propres":1,"Acceptables":0.5,"Sales":0},"to4qf61":{"Très facilement":1,"Difficilement":0.33,"Pas du tout":0},"tg9gh34":{"Très convenable":1,"Acceptable":0.5,"Pas du tout convenable":0},"bm2eu92":{"Trop longue":0,"Acceptable selon le mal":1,"Trop courte":0},"kx3na73":{"Très satisfait":1,"Satisfait":0.67,"Peu satisfait":0.33,"Pas du tout satisfait":0},"si4fc54":{"Oui":1,"Non":0},"st7nh97":{"Oui":1,"Non":0},"uf3af06":{"Oui":1,"Non":0},"nd9id61":{"Oui":1,"Non":0},"di0gd12":{"Oui":1,"Non":0},"co65e31":{"Oui":1,"Non":0},"dt1jh61":{"Bonnes":1,"Acceptables":0.5,"Mauvaise":0},"au1wl40":{"Oui":1,"Non":0},"lc4jc27":{"Oui":1,"Pas tellement":0.5,"Non":0},"lz1xh46":{"Oui":1,"Pas tellement":0.5,"Non":0},"my0pl21":{"Oui":1,"Pas tellement":0.5,"Non":0},"ht7xl14":{"Oui":1,"Pas tellement":0.5,"Non":0},"cn08k85":{"Oui, toujours claires":1,"Parfois ambiguës":0.5,"Non, Souvent confuses ou contradictoires":0},"ye6ms95":{"Oui, en temps utile":1,"Parfois en retard":0.5,"Non, souvent trop tard ou pas du tout":0},"sx8na20":{"Oui":1,"Pas tellement":0.5,"Non":0},"zo7fe17":{"Oui":1,"Pas tellement":0.5,"Non":0},"ui7si32":{"Oui":1,"Pas tellement":0.5,"Non":0},"fg4lo40":{"Oui":1,"Pas tellement":0.5,"Non":0},"pb2zc32":{"Oui, de manière complète":1,"Partiellement ou de façon informelle":0.5,"Non, aucune formation ou sensibilisation reçue":0},"ep8ob11":{"Oui, pleinement soutenu":1,"Moyennement soutenu":0.5,"Non, peu ou pas soutenu":0},"fs59c72":{"Oui, clairement identifiés":1,"Partiellement connus":0.5,"Non, pas du tout connus":0},"kf85b90":{"Oui, ils sont bien expliqués":1,"Moyennement ou de façon peu claire":0.5,"Non, ils ne sont pas expliqués":0},"vt7ev88":{"Oui, dans l’ensemble":1,"Partiellement ou selon les situations":0.5,"Non, rarement ou jamais":0},"ps1pl37":{"Oui":1,"Non":0},"ll2ik06":{"Oui":1,"Pas tellement":0.5,"Non":0},"yh4vk06":{"Oui, la maîtrise est globalement très bonne":1,"Moyennement maîtrisée selon les personnes":0.5,"Non, des lacunes importantes sont observées":0},"wu8sg46":{"Oui, de manière constante":1,"Parfois, selon les situations":0.5,"Non, souvent inadaptée ou peu respectueuse":0},"ye5ft78":{"Oui, la collaboration est fluide et efficace":1,"Moyennement efficace selon les équipes":0.5,"Non, la coordination est souvent difficile":0},"dn1qa13":{"Oui":1,"Pas tellement":0.5,"Non":0},"ke4iv96":{"Oui":1,"Pas tellement":0.5,"Non":0},"dq1ei30":{"Oui":1,"Pas tellement":0.5,"Non":0},"km19b24":{"Oui":1,"Pas tellement":0.5,"Non":0},"aq3kk41":{"Oui":1,"Pas tellement":0.5,"Non":0},"ie4cp00":{"Oui":1,"Pas toujours":0.5,"Non":0},"ak5xg12":{"Oui":1,"Pas tellement":0.5,"Non":0},"fj5iq62":{"Oui":1,"Non":0},"eg9su19":{"Oui":1,"Pas tellement":0.5,"Non":0},"nv6jq01":{"Oui":1,"Pas tellement":0.5,"Non":0},"my17e38":{"Oui":1,"Non":0},"gc9bj55":{"Excellent/Bon":1,"Mauvais":0,"Passable":0.33},"rr4vi05":{"Excellent/Bon":1,"Mauvais":0,"Passable":0.33},"kr3pn80":{"Excellent/Bon":1,"Mauvais":0,"Passable":0.33},"yl4bj31":{"Accueillants":1,"Pas du tout accueillants":0},"kp2xf35":{"Très compétents/Compétents":1,"Pas du tout compétents":0},"he7qz74":{"Très compétents/Compétents":1,"Pas du tout compétents":0},"ll4vy48":{"Très compétents/Compétents":1,"Pas du tout compétents":0},"yw8ak40":{"Très compétents/Compétents":1,"Pas du tout compétents":0},"hz7sr16":{"Oui":1,"Non":0},"dg0fr95":{"Trop long/Long":0,"Raisonnable selon le mal":1,"Trop court":0},"go5va85":{"Oui":1,"Non":0},"lu3ug49":{"Tous les médicaments":1,"Quelques médicaments":0.5,"Aucun médicament":0},"pc5hv66":{"Oui":1,"Pas toujours":0.5,"Non":0},"pr5ud52":{"Très chers/Chers":0,"Abordables":1,"Pas du tout chers":1},"ga1tt79":{"Très chers/Chers":0,"Abordables":1,"Pas du tout chers":1},"qv3pt77":{"Propre":1,"Acceptable":0.5,"Salle":0},"uw5ur14":{"Propre":1,"Acceptable":0.5,"Salle":0},"qf8kr56":{"Propre":1,"Acceptable":0.5,"Salle":0},"sn14p46":{"Propre":1,"Acceptable":0.5,"Salle":0},"pw6gg98":{"Oui, tout à faire":1,"Oui, un peu":0.5,"Pas du tout":0},"wx0vw58":{"Oui":1,"Non":0},"ca3ve40":{"Facilement":1,"Difficilement":0.33,"Pas du tout":0},"tw9po34":{"Oui":1,"Non":0,"Pas directement":0.5},"em0ev27":{"Oui":1,"Non":0,"Pas directement":0.5},"hg9ab90":{"Oui":1,"Non":0,"Pas directement":0.5},"rg7fs81":{"Très bien/Bien":1,"Passable/Pas du tout":0.2},"ny6xw05":{"Oui":1,"Non":0}};
var HOSP_KEYS = ['usagers', 'personnel', 'communaute'];

/** Date « yyyy-MM-dd » d'une cellule (Date, texte ISO ou jj/mm/aaaa) ; '' si illisible. */
function day_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v).trim(), m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}
/** Date de chaque ligne : DateEnvoi (posée par le serveur), sinon « today », sinon « start ». */
function rowDays_(sheet, hm) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var a = readColumn_(sheet, hm, DATE_COLUMN), b = readColumn_(sheet, hm, 'today'), c = readColumn_(sheet, hm, 'start'), n = Math.max(a.length, b.length, c.length), out = [];
  for (var i = 0; i < n; i++) out.push(day_(a[i]) || day_(b[i]) || day_(c[i]));
  return out;
}

function periodsSheet_() {
  var sh = getOrCreateSheet_(PERIODS_SHEET_NAME, PERIODS_HEADERS);
  try { sh.getRange(2, 3, Math.max(1, sh.getMaxRows() - 1), 2).setNumberFormat('@'); } catch (e) {}
  return sh;
}
function periodsList_() {
  var v = periodsSheet_().getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0] || !v[i][1]) continue;
    var st = day_(v[i][2]), en = day_(v[i][3]);
    if (!st || !en) continue;
    out.push({ id: String(v[i][0]), label: String(v[i][1]), start: st, end: en, active: v[i][4] !== 'NON' });
  }
  out.sort(function (a, b) { return a.start < b.start ? 1 : a.start > b.start ? -1 : 0; });
  return out;
}
function sessionAccountView_(token) {
  var s = token ? findSession_(token) : null, a = s ? findAccountRow_(s.username) : null, acc = a ? accountPublicView_(a.data) : null;
  return acc && acc.active ? acc : null;
}
function handleListPeriods(token) {
  if (!sessionAccountView_(token)) return { ok: false, error: 'Session invalide ou expirée, reconnectez-vous.' };
  return { ok: true, periods: periodsList_() };
}
function handleSavePeriod(token, p) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  p = p || {};
  var label = String(p.label || '').trim(), st = day_(p.start), en = day_(p.end);
  if (!label) return { ok: false, error: 'Le nom de la période est obligatoire.' };
  if (!st || !en) return { ok: false, error: 'Dates de début et de fin obligatoires.' };
  if (st > en) return { ok: false, error: 'La date de début doit précéder la date de fin.' };
  var sh = periodsSheet_(), v = sh.getDataRange().getValues(), row = -1;
  for (var i = 1; i < v.length; i++) {
    if (p.id && String(v[i][0]) === String(p.id)) row = i + 1;
    if (String(v[i][1]).trim().toLowerCase() === label.toLowerCase() && String(v[i][0]) !== String(p.id || '')) return { ok: false, error: 'Une période porte déjà ce nom.' };
  }
  if (row > 0) sh.getRange(row, 2, 1, 4).setValues([[label, st, en, 'OUI']]);
  else sh.appendRow([Utilities.getUuid().slice(0, 8), label, st, en, 'OUI']);
  return { ok: true, periods: periodsList_() };
}
function handleDeletePeriod(token, id) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var sh = periodsSheet_(), v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(id)) sh.deleteRow(i + 1);
  return { ok: true, periods: periodsList_() };
}

/** Texte sans accents ni ponctuation, pour comparer noms d'hôpitaux et de zones. */
function strip_(s) {
  var m = { 'à': 'a', 'â': 'a', 'ä': 'a', 'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e', 'î': 'i', 'ï': 'i', 'ô': 'o', 'ö': 'o', 'ù': 'u', 'û': 'u', 'ü': 'u', 'ç': 'c', 'œ': 'oe' };
  return String(s || '').toLowerCase().replace(/[àâäéèêëîïôöùûüçœ]/g, function (c) { return m[c]; }).replace(/[^a-z0-9]+/g, ' ').trim();
}
function zoneKey_(s) { return strip_(s).replace(/\bzone sanitaire\b|\bzs\b/g, ' ').replace(/\s+/g, ' ').trim(); }
/** Portée d'un compte sur une ligne d'enquête hôpital. Compte « hopital » : nom de l'hôpital dans la colonne Commune de la feuille Comptes. */
function hospInScope_(hosp, zone, dept, account) {
  var lvl = account.level;
  if (lvl === 'admin' || lvl === 'national' || lvl === 'superviseur') return true;
  if (lvl === 'hopital') return strip_(hosp) === strip_(account.commune);
  if (lvl === 'zone') return zoneKey_(zone) !== '' && zoneKey_(zone) === zoneKey_(account.zone);
  if (lvl === 'departement') return strip_(dept) !== '' && strip_(dept) === strip_(account.departement);
  return false;
}
function inPeriod_(day, p) { return !p || (!!day && day >= p.start && day <= p.end); }
function aggRows_(rows, p) {
  var n = 0, s = 0, sn = 0;
  rows.forEach(function (r) { if (!inPeriod_(r.d, p)) return; n++; if (typeof r.o === 'number') { s += r.o; sn++; } });
  return { count: n, avgScore: sn > 0 ? clampPct_(s / sn) : null };
}

/** Enquêtes hôpitaux : score par rubrique (moyenne des réponses notées, 0 à 100 %). */
function hospitalData_(account, sel, periods) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), out = {}, trendRows = {}, errors = [];
  HOSP_KEYS.forEach(function (key) {
    var cfg = HOSP_CFG[key], names = cfg.cats.map(function (c) { return c.n; });
    var res = { count: 0, avgScore: null, categories: names, all: { count: 0, categoryScores: {} }, byHopital: [] }, rows = [];
    try {
      var sh = ss.getSheetByName(cfg.sheet);
      if (sh && sh.getLastRow() >= 2) {
        var v = sh.getDataRange().getValues(), ci = {};
        v[0].forEach(function (h, i) { var t = String(h || '').trim(); if (t && ci[t] === undefined) ci[t] = i; });
        var iH = ci[cfg.hosp], iZ = ci[cfg.zone], iD = ci[cfg.dept], i1 = ci[DATE_COLUMN], i2 = ci['today'], i3 = ci['start'];
        if (iH !== undefined) for (var r = 1; r < v.length; r++) {
          var row = v[r], hosp = String(row[iH] || '').trim();
          if (!hosp) continue;
          var zone = iZ === undefined ? '' : String(row[iZ] || '').trim(), dept = iD === undefined ? '' : String(row[iD] || '').trim();
          if (!hospInScope_(hosp, zone, dept, account)) continue;
          var cs = {}, tot = 0, tn = 0;
          cfg.cats.forEach(function (cat) {
            var s = 0, n = 0;
            cat.items.forEach(function (it) {
              var k = ci[it.c]; if (k === undefined) return;
              var lab = String(row[k] === undefined || row[k] === null ? '' : row[k]).trim(), m = HOSP_SCALES[it.l];
              if (!lab || !m || m[lab] === undefined) return;
              var sc = m[lab]; if (it.i) sc = 1 - sc;
              s += sc; n++;
            });
            cs[cat.n] = n > 0 ? s / n : null;
            tot += s; tn += n;
          });
          rows.push({ h: hosp, z: zone, d: (i1 !== undefined ? day_(row[i1]) : '') || (i2 !== undefined ? day_(row[i2]) : '') || (i3 !== undefined ? day_(row[i3]) : ''), cs: cs, o: tn > 0 ? tot / tn : null });
        }
      }
    } catch (e) { errors.push(key + ' : ' + e.message); }
    trendRows[key] = rows;
    var sel_ = rows.filter(function (r) { return inPeriod_(r.d, sel); }), by = {};
    function add(o, r) {
      o.count++;
      names.forEach(function (n) { if (typeof r.cs[n] === 'number') { o.sum[n] = (o.sum[n] || 0) + r.cs[n]; o.n[n] = (o.n[n] || 0) + 1; } });
      if (typeof r.o === 'number') { o.os += r.o; o.on++; }
    }
    function fin(o) {
      var pct = {}; names.forEach(function (n) { pct[n] = o.n[n] > 0 ? clampPct_(o.sum[n] / o.n[n]) : null; });
      return { count: o.count, avgScore: o.on > 0 ? clampPct_(o.os / o.on) : null, categoryScores: pct };
    }
    var all = { count: 0, sum: {}, n: {}, os: 0, on: 0 };
    sel_.forEach(function (r) {
      add(all, r);
      if (!by[r.h]) by[r.h] = { o: { count: 0, sum: {}, n: {}, os: 0, on: 0 }, z: r.z };
      add(by[r.h].o, r);
    });
    var fa = fin(all);
    res.count = fa.count; res.avgScore = fa.avgScore; res.all = { count: fa.count, categoryScores: fa.categoryScores };
    Object.keys(by).forEach(function (h) { var f = fin(by[h].o); res.byHopital.push({ label: h, zone: by[h].z, count: f.count, avgScore: f.avgScore, categoryScores: f.categoryScores }); });
    res.byHopital.sort(function (a, b) { return b.count - a.count; });
    out[key] = res;
  });
  var trend = {};
  periods.forEach(function (p) { trend[p.id] = {}; HOSP_KEYS.forEach(function (k) { trend[p.id][k] = aggRows_(trendRows[k], p); }); });
  var und = {}; HOSP_KEYS.forEach(function (k) { und[k] = aggRows_(trendRows[k].filter(function (r) { return !r.d; }), null); });
  return { data: out, trend: trend, trendUndated: und, errors: errors };
}

/* ---------- Enquêteurs : factures normalisées (PDF) et attestations de bonne fin d'exécution ---------- */
// Les factures PDF sont enregistrées directement dans Google Sheets (feuille « FacturesFichiers », découpées en morceaux : aucune autorisation Google Drive n'est nécessaire). Les données (nom, prénom, validation) sont dans
// les feuilles Enqueteurs, Factures et AttestParams. Seul le compte « admin » peut les lire, les valider et les modifier.
var ENQ_SHEET_NAME = 'Enqueteurs', FACT_SHEET_NAME = 'Factures', ATT_SHEET_NAME = 'AttestParams';
var ENQ_HEADERS = ['ID', 'IdActivité', 'Activité', 'Niveau', 'Cible', 'Début', 'Fin', 'Nom', 'Prénom', 'Clé', 'CrééLe', 'Validée', 'ValidéeLe', 'Numéro', 'MoMo'];
var ENQ_ID_COL = 'IdEnquêteur', ENQ_NUM_COL = 'N° enquêteur'; // colonnes ajoutées par le serveur aux feuilles de réponses
var FACT_HEADERS = ['ID', 'IdEnqueteur', 'Fichier', 'FileId', 'Taille', 'AjoutéLe'];
var ATT_HEADERS = ['IdActivité', 'Libellé', 'Début', 'Fin', 'Année', 'Lieu', 'DateSignature', 'Qualité', 'ModifiéLe'];
var FACT_MAX_FILES = 5;                 // factures par envoi
var FACT_MAX_BYTES = 6 * 1024 * 1024;   // 6 Mo par facture
var FACT_MAX_PER_ENQ = 40;              // factures au total par enquêteur et par activité
var FACT_GRACE_DAYS = 60;               // le dépôt reste ouvert 60 jours après la fin de l'activité
var FICH_SHEET_NAME = 'FacturesFichiers', FICH_HEADERS = ['IdFacture', 'Partie', 'Données'];
var FICH_CHUNK = 40000;                 // caractères par cellule (limite Google Sheets : 50 000)

/** À exécuter UNE FOIS dans l'éditeur Apps Script : crée les feuilles (Enqueteurs, Factures, FacturesFichiers, AttestParams). */
function setupEnqueteursEtFactures() {
  enqSheet_(); tsheet_(FACT_SHEET_NAME, FACT_HEADERS); tsheet_(ATT_SHEET_NAME, ATT_HEADERS); tsheet_(FICH_SHEET_NAME, FICH_HEADERS);
  Logger.log('Feuilles prêtes. Les factures sont stockées dans la feuille ' + FICH_SHEET_NAME + '.');
}
/** Feuille Enqueteurs (registre des enquêteurs inscrits) ; ajoute la colonne « Numéro » aux anciennes feuilles. */
function enqSheet_() {
  var sh = tsheet_(ENQ_SHEET_NAME, ENQ_HEADERS), n = ENQ_HEADERS.length;
  if (sh.getLastColumn() < n || String(sh.getRange(1, n).getValue()).trim() !== ENQ_HEADERS[n - 1]) {
    if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
    sh.getRange(1, n).setValue(ENQ_HEADERS[n - 1]);
    try { sh.getRange(2, n, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@'); } catch (e) {}
  }
  return sh;
}
function tsheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name); sh.appendRow(headers);
    try { sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@'); sh.setFrozenRows(1); } catch (e) {}
  }
  return sh;
}
/** Enregistre un PDF (base64) dans la feuille FacturesFichiers, en morceaux de FICH_CHUNK caractères. */
function storeFichier_(id, b64) {
  var sh = tsheet_(FICH_SHEET_NAME, FICH_HEADERS), rows = [];
  for (var i = 0, p = 1; i < b64.length; i += FICH_CHUNK, p++) rows.push([id, String(p), b64.substr(i, FICH_CHUNK)]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
}
/** Lit le PDF (base64) d'une facture depuis la feuille FacturesFichiers. */
function readFichier_(id) {
  var sh = tsheet_(FICH_SHEET_NAME, FICH_HEADERS), last = sh.getLastRow(); if (last < 2) return null;
  var ids = sh.getRange(2, 1, last - 1, 2).getValues(), parts = [];
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) parts.push({ row: i + 2, n: Number(ids[i][1]) || 0 });
  if (!parts.length) return null;
  parts.sort(function (a, b) { return a.n - b.n; });
  var out = [], k = 0;
  while (k < parts.length) { // lecture par blocs de lignes contiguës
    var j = k; while (j + 1 < parts.length && parts[j + 1].row === parts[j].row + 1) j++;
    sh.getRange(parts[k].row, 3, j - k + 1, 1).getValues().forEach(function (r) { out.push(String(r[0])); });
    k = j + 1;
  }
  return out.join('');
}
/** Supprime les morceaux d'une facture (ou, pour d'anciennes factures, met le fichier Drive à la corbeille). */
function trashFactureFile_(fileId, factId) {
  if (String(fileId) !== 'SHEET') { try { DriveApp.getFileById(String(fileId)).setTrashed(true); } catch (e) {} return; }
  var sh = tsheet_(FICH_SHEET_NAME, FICH_HEADERS), last = sh.getLastRow(); if (last < 2) return;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) !== String(factId)) continue;
    var end = i; while (i - 1 >= 0 && String(ids[i - 1][0]) === String(factId)) i--;
    sh.deleteRows(i + 2, end - i + 1);
  }
}
function rowsOf_(sh, ncol) {
  if (sh.getLastRow() < 2) return [];
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, ncol).getValues(), out = [];
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() !== '') out.push({ row: i + 2, v: v[i] });
  return out;
}
function cleanName_(s) { return String(s || '').replace(/[\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').replace(/^[=+\-@\s]+/, '').trim().slice(0, 60); }
function fmtNom_(s) { return cleanName_(s).toUpperCase(); }
function fmtPrenom_(s) {
  s = cleanName_(s);
  if (s !== s.toLowerCase() && s !== s.toUpperCase()) return s; // déjà saisi avec majuscules et minuscules : on ne touche pas
  return s.toLowerCase().replace(/(^|[\s\-'’])([^\s\-'’])/g, function (m, a, b) { return a + b.toUpperCase(); });
}
function nameKey_(nom, prenom) { return strip_(nom) + '|' + strip_(prenom); }
function newId_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase(); }
function codeFromText_(raw) {
  raw = String(raw || '').trim(); var cand = raw, q = raw.indexOf('?');
  if (raw.indexOf('http') === 0 && q !== -1) {
    raw.substring(q + 1).split('#')[0].split('&').forEach(function (p) { var kv = p.split('='); if (kv[0] === 'code' || kv[0] === 't') cand = decodeURIComponent(kv.slice(1).join('=')); });
  }
  return String(cand).trim();
}
/** Activité (non désactivée) correspondant au code, jusqu'à FACT_GRACE_DAYS jours après sa fin. */
function activityForInvoice_(code) {
  code = String(code || '').trim().toUpperCase(); if (!code) return null;
  var today = day_(new Date()), f = null;
  activitiesRows_().forEach(function (a) {
    if (a.code.toUpperCase() !== code || !a.active) return;
    if (a.fin && today > dayPlus_(a.fin, FACT_GRACE_DAYS)) return;
    f = a;
  });
  return f;
}
function cleanNum_(s) { return String(s || '').replace(/[^A-Za-z0-9\-\/]/g, '').slice(0, 15); }
function numKey_(s) { var k = String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); return /^\d+$/.test(k) ? k.replace(/^0+(?=\d)/, '') : k; } // « 07 » et « 7 » = même numéro
function enqView_(r) { var v = r.v; return { id: String(v[0]), actId: String(v[1]), nom: String(v[7]), prenom: String(v[8]), numero: String(v[13] || ''), valide: String(v[11]) === 'OUI', momo: String(v[14] || '') }; }
/** Enquêteur inscrit d'une activité, par son identifiant. */
function enqById_(actId, id) {
  var f = null;
  rowsOf_(enqSheet_(), ENQ_HEADERS.length).forEach(function (r) { if (String(r.v[0]) === String(id) && String(r.v[1]) === String(actId)) f = enqView_(r); });
  return f;
}
/** Liste des enquêteurs inscrits à une activité (ordre d'inscription). */
function enqList_(actId) {
  return rowsOf_(enqSheet_(), ENQ_HEADERS.length).filter(function (r) { return String(r.v[1]) === String(actId); }).map(enqView_);
}
/** Inscrit l'enquêteur (ou le retrouve) : le numéro d'enquêteur est unique dans l'activité. À appeler sous verrou. */
function enqUpsert_(a, nom, prenom, numero) {
  var sh = enqSheet_(), rows = rowsOf_(sh, ENQ_HEADERS.length), key = nameKey_(nom, prenom), nk = numKey_(numero), byNum = null, byName = null;
  rows.forEach(function (r) {
    if (String(r.v[1]) !== a.id) return;
    if (nk && numKey_(r.v[13]) === nk) byNum = r;
    if (String(r.v[9]) === key) byName = r;
  });
  if (byNum) {
    if (String(byNum.v[9]) !== key) return { ok: false, error: 'Ce numéro d\'enquêteur (' + numero + ') est déjà attribué à ' + String(byNum.v[8]) + ' ' + String(byNum.v[7]) + '. Vérifiez votre numéro, votre nom et votre prénom.' };
    return { ok: true, created: false, row: byNum.row, id: String(byNum.v[0]), nom: String(byNum.v[7]), prenom: String(byNum.v[8]), numero: String(byNum.v[13]) };
  }
  if (byName) {
    var old = numKey_(byName.v[13]);
    if (old && old !== nk) return { ok: false, error: 'Ce nom et ce prénom sont déjà enregistrés avec le numéro d\'enquêteur ' + String(byName.v[13]) + '.' };
    if (!old) sh.getRange(byName.row, 14).setNumberFormat('@').setValue(numero); // ancien enquêteur sans numéro : on le complète
    return { ok: true, created: false, row: byName.row, id: String(byName.v[0]), nom: String(byName.v[7]), prenom: String(byName.v[8]), numero: numero };
  }
  var id = newId_('E');
  sh.getRange(sh.getLastRow() + 1, 1, 1, ENQ_HEADERS.length).setNumberFormat('@').setValues([[id, a.id, a.name, a.level, a.cible, a.debut, a.fin, nom, prenom, key, new Date().toISOString(), 'NON', '', numero, '']]);
  return { ok: true, created: true, id: id, nom: nom, prenom: prenom, numero: numero };
}
/** Enquêteur déjà inscrit (lecture seule) : numéro + nom + prénom ; pour les anciens enquêteurs sans numéro, nom + prénom. */
function enqFind_(a, nom, prenom, numero) {
  var key = nameKey_(nom, prenom), nk = numKey_(numero), f = null;
  rowsOf_(enqSheet_(), ENQ_HEADERS.length).forEach(function (r) {
    if (String(r.v[1]) !== a.id || String(r.v[9]) !== key) return;
    var rn = numKey_(r.v[13]); if (rn && nk && rn !== nk) return;
    f = r;
  });
  return f;
}
/** Première connexion de l'enquêteur : nom, prénom et numéro d'enquêteur (accès par le code de son activité). */
/** Tableau de bord de l'enquêteur inscrit : par cible, questionnaires validés, rejetés, objectif et reste à atteindre sur la période. */
/** Numéro MTN MoMo : exactement 8 chiffres (sans 01 ni indicatif). */
function cleanMomo_(x) { var t = String(x == null ? '' : x).replace(/[\s.\-]/g, ''); return /^\d{8}$/.test(t) ? t : ''; }
function enqSetMomo_(actId, enqId, momo) {
  var sh = enqSheet_(), ok = false;
  rowsOf_(sh, ENQ_HEADERS.length).forEach(function (r) { if (String(r.v[0]) === String(enqId) && String(r.v[1]) === String(actId)) { sh.getRange(r.row, 15).setNumberFormat('@').setValue(momo); ok = true; } });
  return ok;
}
var PAY_SHEET = 'Paiements', PAY_HEADERS = ['IdActivité', 'IdEnquêteur', 'Montant', 'Statut', 'Date', 'Par', 'NumeroMoMo'];
/** Paiement enregistré (statut PAYE) d'un enquêteur pour une activité, ou null. */
function paymentOf_(actId, enqId) {
  var f = null; suiviRows_(PAY_SHEET, PAY_HEADERS).forEach(function (r) { if (String(r.v[0]) === String(actId) && String(r.v[1]) === String(enqId) && String(r.v[3]) === 'PAYE') f = { montant: Number(r.v[2]) || 0, date: String(r.v[4]), par: String(r.v[5]), momo: String(r.v[6]), row: r.row }; });
  return f;
}
/** L'enquêteur ajoute ou modifie son numéro MTN MoMo (8 chiffres). */
function handleSetMomo(d) {
  var a = activityForInvoice_(d.code); if (!a) return { ok: false, error: 'ACT_INVALID' };
  var en = enqById_(a.id, d.id); if (!en) return { ok: false, error: 'ENQ_INVALID' };
  var m = cleanMomo_(d.momo); if (!m) return { ok: false, error: 'Le numéro MTN MoMo doit comporter exactement 8 chiffres (sans le 01).' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try { enqSetMomo_(a.id, en.id, m); SpreadsheetApp.flush(); return { ok: true, momo: m }; } finally { lock.releaseLock(); }
}
function handleMyCounts(d) {
  var a = activityForInvoice_(d.code); if (!a) return { ok: false, error: 'ACT_INVALID' };
  var en = enqById_(a.id, d.id); if (!en) return { ok: false, error: 'ENQ_INVALID' };
  var act = actById_(a.id); if (!act) return { ok: false, error: 'ACT_INVALID' };
  var appareil = devCheck_(a.id, en.id, d.dev), c = suiviCompute_(act), me = null, pay = paymentOf_(a.id, en.id);
  c.enqueteurs.forEach(function (e) { if (String(e.id) === String(en.id)) me = e; });
  var cibles = {};
  c.forms.forEach(function (f) {
    var v = me ? (me.valid[f] || 0) : 0, r = me ? (me.rej[f] || 0) : 0, o = me && me.objByForm[f] !== undefined ? me.objByForm[f] : null;
    cibles[f] = { valides: v, rejetes: r, objectif: o, reste: o === null ? null : Math.max(0, o - v) };
  });
  return { ok: true, appareil: appareil, momo: en.momo, paiement: pay ? { montant: pay.montant, date: pay.date, momo: pay.momo } : null, cibles: cibles, mins: c.mins, debut: act.debut || '', fin: act.fin || '', joursRestants: c.totalDays ? Math.max(0, c.totalDays - c.elapsed) : null,
    rejDuree: me ? me.rejDuree : 0, rejInterv: me ? me.rejInterv : 0, rejPlafond: me ? me.rejPlafond : 0,
    quotas: c.quotas, montantCentre: c.mins.montantCentre, modeCentre: !!(c.mins.montantCentre > 0 && Object.keys(c.quotas).length), centres: me && me.centresDetail ? me.centresDetail : [], gain: me && me.centresDetail ? me.aPayer : null };
}
function handleRegisterEnqueteur(d) {
  var a = activityForInvoice_(d.code); if (!a) return { ok: false, error: 'ACT_INVALID' };
  var nom = fmtNom_(d.nom), prenom = fmtPrenom_(d.prenom), numero = cleanNum_(d.numero);
  if (!nom || !prenom) return { ok: false, error: 'Indiquez votre nom et votre prénom.' };
  if (!numero) return { ok: false, error: 'Indiquez votre numéro d\'enquêteur (lettres, chiffres et tirets).' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var u = enqUpsert_(a, nom, prenom, numero); if (!u.ok) return u;
    var mo = String(d.momo == null ? '' : d.momo).trim();
    if (mo) { var m8 = cleanMomo_(mo); if (!m8) return { ok: false, error: 'Le numéro MTN MoMo doit comporter exactement 8 chiffres (sans le 01).' }; enqSetMomo_(a.id, u.id, m8); }
    SpreadsheetApp.flush();
    var ap = devCheck_(a.id, u.id, d.dev);
    return { ok: true, id: u.id, nom: u.nom, prenom: u.prenom, numero: u.numero, created: u.created, activite: a.name, appareil: ap, momo: (enqById_(a.id, u.id) || {}).momo || '' };
  } finally { lock.releaseLock(); }
}
function handleMatchFactureCode(rawText) {
  var cand = codeFromText_(rawText); if (!cand) return { ok: false, error: 'QR code vide.' };
  var a = activityForInvoice_(cand);
  if (!a) return { ok: false, error: "Ce code ne correspond à aucune activité ouverte au dépôt de factures (activité inconnue, désactivée ou terminée depuis trop longtemps)." };
  return { ok: true, activity: activityPublic_(a), label: a.name + (a.cible ? ' — ' + a.cible : '') };
}
/** Dépôt d'une ou plusieurs factures PDF par un enquêteur (accès par le code de son activité). */
function handleSubmitFacture(d) {
  var a = activityForInvoice_(d.code); if (!a) return { ok: false, error: 'ACT_INVALID' };
  var nom = fmtNom_(d.nom), prenom = fmtPrenom_(d.prenom);
  var numero = cleanNum_(d.numero);
  if (!nom || !prenom) return { ok: false, error: 'Indiquez votre nom et votre prénom.' };
  if (!numero) return { ok: false, error: 'Indiquez votre numéro d\'enquêteur.' };
  var files = d.files || [];
  if (!files.length) return { ok: false, error: 'Ajoutez au moins une facture au format PDF.' };
  if (files.length > FACT_MAX_FILES) return { ok: false, error: 'Maximum ' + FACT_MAX_FILES + ' factures par envoi.' };
  var blobs = [];
  for (var i = 0; i < files.length; i++) {
    var bytes;
    try { bytes = Utilities.base64Decode(String(files[i].b64 || '')); } catch (e) { return { ok: false, error: 'Fichier illisible : ' + cleanName_(files[i].name) }; }
    if (!bytes.length || bytes.length > FACT_MAX_BYTES) return { ok: false, error: 'La facture « ' + cleanName_(files[i].name) + ' » dépasse ' + Math.round(FACT_MAX_BYTES / 1048576) + ' Mo ou est vide.' };
    if (!(bytes[0] === 37 && bytes[1] === 80 && bytes[2] === 68 && bytes[3] === 70)) return { ok: false, error: 'La facture « ' + cleanName_(files[i].name) + ' » n\'est pas un fichier PDF.' };
    var orig = cleanName_(files[i].name).replace(/[\\\/:*?"<>|]+/g, '-').replace(/\.pdf$/i, '').slice(0, 60) || 'facture';
    blobs.push({ bytes: bytes, name: orig, b64: String(files[i].b64 || '').replace(/\s+/g, '') });
  }
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var fsh = tsheet_(FACT_SHEET_NAME, FACT_HEADERS), u = enqUpsert_(a, nom, prenom, numero);
    if (!u.ok) return u;
    var eid = u.id; nom = u.nom; prenom = u.prenom;
    var have = rowsOf_(fsh, FACT_HEADERS.length).filter(function (r) { return String(r.v[1]) === eid; }).length;
    if (have + blobs.length > FACT_MAX_PER_ENQ) return { ok: false, error: 'Limite de ' + FACT_MAX_PER_ENQ + ' factures atteinte pour cet enquêteur.' };
    var now = new Date().toISOString();
    blobs.forEach(function (b) {
      var fid = newId_('F');
      storeFichier_(fid, b.b64);
      fsh.getRange(fsh.getLastRow() + 1, 1, 1, FACT_HEADERS.length).setNumberFormat('@').setValues([[fid, eid, b.name + '.pdf', 'SHEET', String(b.bytes.length), now]]);
    });
    SpreadsheetApp.flush();
    return { ok: true, nom: nom, prenom: prenom, ajoutees: blobs.length, total: have + blobs.length, activite: a.name };
  } finally { lock.releaseLock(); }
}
/** Pour l'enquêteur : renvoie les données de SON attestation, uniquement si l'administrateur l'a validée.
 *  Accès : code de l'activité + nom + prénom (les mêmes que pour le dépôt). */
function handleGetAttestation(d) {
  var a = activityForInvoice_(d.code); if (!a) return { ok: false, error: 'ACT_INVALID' };
  var nom = fmtNom_(d.nom), prenom = fmtPrenom_(d.prenom);
  if (!nom || !prenom) return { ok: false, error: 'Indiquez votre nom et votre prénom.' };
  var ex = enqFind_(a, nom, prenom, cleanNum_(d.numero));
  if (!ex) return { ok: true, status: 'inconnu' };
  var nf = rowsOf_(tsheet_(FACT_SHEET_NAME, FACT_HEADERS), FACT_HEADERS.length).filter(function (r) { return String(r.v[1]) === String(ex.v[0]); }).length;
  if (String(ex.v[11]) !== 'OUI') return { ok: true, status: 'attente', factures: nf, nom: String(ex.v[7]), prenom: String(ex.v[8]) };
  var p = attestParamsMap_()[a.id];
  if (!p || !p.dateSignature) return { ok: true, status: 'attente', factures: nf, nom: String(ex.v[7]), prenom: String(ex.v[8]) };
  return { ok: true, status: 'valide', nom: String(ex.v[7]), prenom: String(ex.v[8]), level: a.level, cible: a.cible, debut: a.debut, fin: a.fin, params: p };
}
function attestParamsMap_() {
  var m = {};
  rowsOf_(tsheet_(ATT_SHEET_NAME, ATT_HEADERS), ATT_HEADERS.length).forEach(function (r) {
    var v = r.v; m[String(v[0])] = { libelle: String(v[1] || ''), debut: day_(v[2]), fin: day_(v[3]), annee: String(v[4] || ''), lieu: String(v[5] || ''), dateSignature: day_(v[6]), qualite: String(v[7] || '') };
  });
  return m;
}
/** Liste complète pour l'administrateur : activités (avec paramètres d'attestation), enquêteurs et leurs factures. */
function handleListEnqueteurs(token) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var live = {}; activitiesRows_().forEach(function (a) { live[a.id] = a; });
  var params = attestParamsMap_(), fm = {}, acts = {}, enq = [];
  rowsOf_(tsheet_(FACT_SHEET_NAME, FACT_HEADERS), FACT_HEADERS.length).forEach(function (r) {
    var k = String(r.v[1]); (fm[k] = fm[k] || []).push({ id: String(r.v[0]), name: String(r.v[2]), size: Number(r.v[4]) || 0, date: String(r.v[5]) });
  });
  rowsOf_(enqSheet_(), ENQ_HEADERS.length).forEach(function (r) {
    var v = r.v, aid = String(v[1]);
    if (!acts[aid]) {
      var l = live[aid];
      acts[aid] = l ? { id: aid, name: l.name, level: l.level, cible: l.cible, debut: l.debut, fin: l.fin, active: l.active, supprimee: false }
                    : { id: aid, name: String(v[2]), level: String(v[3]), cible: String(v[4]), debut: day_(v[5]), fin: day_(v[6]), active: false, supprimee: true };
      acts[aid].params = params[aid] || null;
    }
    enq.push({ id: String(v[0]), actId: aid, nom: String(v[7]), prenom: String(v[8]), numero: String(v[13] || ''), created: String(v[10]), valide: String(v[11]) === 'OUI', valideLe: String(v[12] || ''), factures: (fm[String(v[0])] || []) });
  });
  var list = Object.keys(acts).map(function (k) { return acts[k]; });
  // les activités encore ouvertes sans enquêteur sont aussi proposées (pour régler leurs paramètres à l'avance)
  Object.keys(live).forEach(function (k) { if (!acts[k]) { var l = live[k]; list.push({ id: k, name: l.name, level: l.level, cible: l.cible, debut: l.debut, fin: l.fin, active: l.active, supprimee: false, params: params[k] || null }); } });
  list.sort(function (x, y) { return (x.debut || '') < (y.debut || '') ? 1 : (x.debut || '') > (y.debut || '') ? -1 : 0; });
  return { ok: true, activities: list, enqueteurs: enq };
}
function handleGetFacture(token, id) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var f = rowsOf_(tsheet_(FACT_SHEET_NAME, FACT_HEADERS), FACT_HEADERS.length).filter(function (r) { return String(r.v[0]) === String(id); })[0];
  if (!f) return { ok: false, error: 'Facture introuvable.' };
  if (String(f.v[3]) === 'SHEET') {
    var b64 = readFichier_(id);
    return b64 ? { ok: true, name: String(f.v[2]), b64: b64 } : { ok: false, error: 'Contenu de la facture introuvable dans la feuille ' + FICH_SHEET_NAME + '.' };
  }
  try { var b = DriveApp.getFileById(String(f.v[3])).getBlob(); return { ok: true, name: String(f.v[2]), b64: Utilities.base64Encode(b.getBytes()) }; } // anciennes factures stockées sur Drive
  catch (e) { return { ok: false, error: "Fichier introuvable (supprimé ?)." }; }
}
/** Valide ou retire la validation de l'attestation d'une liste d'enquêteurs. */
function handleSetAttestationValide(token, ids, valide) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  if (!ids || !ids.length) return { ok: false, error: 'Aucun enquêteur.' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = enqSheet_(), set = {}; ids.forEach(function (i) { set[String(i)] = 1; });
    var n = 0, now = new Date().toISOString();
    rowsOf_(sh, ENQ_HEADERS.length).forEach(function (r) {
      if (!set[String(r.v[0])]) return;
      sh.getRange(r.row, 12, 1, 2).setNumberFormat('@').setValues([[valide ? 'OUI' : 'NON', valide ? now : '']]); n++;
    });
    return { ok: true, modifies: n, date: valide ? now : '' };
  } finally { lock.releaseLock(); }
}
function handleUpdateEnqueteur(token, id, nom, prenom) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  nom = fmtNom_(nom); prenom = fmtPrenom_(prenom);
  if (!nom || !prenom) return { ok: false, error: 'Nom et prénom obligatoires.' };
  var sh = enqSheet_(), rows = rowsOf_(sh, ENQ_HEADERS.length), me = null, key = nameKey_(nom, prenom);
  rows.forEach(function (r) { if (String(r.v[0]) === String(id)) me = r; });
  if (!me) return { ok: false, error: 'Enquêteur introuvable.' };
  if (rows.some(function (r) { return String(r.v[0]) !== String(id) && String(r.v[1]) === String(me.v[1]) && String(r.v[9]) === key; })) return { ok: false, error: 'Un autre enquêteur de cette activité porte déjà ce nom et ce prénom.' };
  sh.getRange(me.row, 8, 1, 3).setNumberFormat('@').setValues([[nom, prenom, key]]);
  return { ok: true, nom: nom, prenom: prenom };
}
function handleDeleteFacture(token, id) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var sh = tsheet_(FACT_SHEET_NAME, FACT_HEADERS), f = rowsOf_(sh, FACT_HEADERS.length).filter(function (r) { return String(r.v[0]) === String(id); })[0];
  if (!f) return { ok: false, error: 'Facture introuvable.' };
  trashFactureFile_(f.v[3], f.v[0]); sh.deleteRow(f.row);
  return { ok: true };
}
function handleDeleteEnqueteur(token, id) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var fsh = tsheet_(FACT_SHEET_NAME, FACT_HEADERS), rows = rowsOf_(fsh, FACT_HEADERS.length);
    for (var i = rows.length - 1; i >= 0; i--) if (String(rows[i].v[1]) === String(id)) { trashFactureFile_(rows[i].v[3], rows[i].v[0]); fsh.deleteRow(rows[i].row); }
    var esh = enqSheet_(), e = rowsOf_(esh, ENQ_HEADERS.length).filter(function (r) { return String(r.v[0]) === String(id); })[0];
    if (!e) return { ok: false, error: 'Enquêteur introuvable.' };
    esh.deleteRow(e.row);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
/** Paramètres modifiables de l'attestation d'une activité (zone / hôpital affiché, dates, année, lieu et date de signature). */
function handleSaveAttestParams(token, p) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  p = p || {}; var id = String(p.actId || '').trim(); if (!id) return { ok: false, error: 'Activité inconnue.' };
  var rx = /^\d{4}-\d{2}-\d{2}$/, debut = String(p.debut || ''), fin = String(p.fin || ''), ds = String(p.dateSignature || ''), annee = String(p.annee || '').trim();
  if ((debut && !rx.test(debut)) || (fin && !rx.test(fin)) || (ds && !rx.test(ds))) return { ok: false, error: 'Date invalide.' };
  if (debut && fin && fin < debut) return { ok: false, error: 'La date de fin précède la date de début.' };
  if (annee && !/^\d{4}$/.test(annee)) return { ok: false, error: 'Année invalide.' };
  var row = [id, cleanName_(p.libelle).slice(0, 80), debut, fin, annee, cleanName_(p.lieu).slice(0, 60), ds, cleanName_(p.qualite).slice(0, 60), new Date().toISOString()];
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = tsheet_(ATT_SHEET_NAME, ATT_HEADERS), ex = rowsOf_(sh, ATT_HEADERS.length).filter(function (r) { return String(r.v[0]) === id; })[0];
    sh.getRange(ex ? ex.row : sh.getLastRow() + 1, 1, 1, ATT_HEADERS.length).setNumberFormat('@').setValues([row]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* ---------- Suivi de la collecte : superviseurs d'activité, objectifs par enquêteur, effectif du personnel ---------- */
// Feuilles : Affectations (superviseur ↔ activité), Objectifs (par activité, enquêteur et formulaire), Effectifs (personnel par structure).
// Compte « superviseur d'activité » (niveau superviseur_act) : ne voit QUE le suivi des activités qui lui sont affectées.
var SUP_LEVEL = 'superviseur_act';
var AFF_SHEET = 'Affectations', AFF_HEADERS = ['IdActivité', 'Superviseur'];
var OBJ_SHEET = 'Objectifs', OBJ_HEADERS = ['IdActivité', 'Enquêteur', 'Cible', 'Objectif'];
var EFF_SHEET = 'Effectifs', EFF_HEADERS = ['IdActivité', 'Structure', 'Effectif'];
var PAR_SHEET = 'SuiviParams', PAR_HEADERS = ['IdActivité', 'MontantTotal', 'ModifiéLe', 'ModifiéPar', 'DuréeMinMin', 'IntervalleMinMin', 'TailleCentre', 'PlafondPct', 'MontantCentre', 'NbCentres'];
var QC_SHEET = 'QuotasCentre', QC_HEADERS = ['IdActivité', 'Cible', 'Quota'];
var AP_SHEET = 'AppareilsEnq', AP_HEADERS = ['IdActivité', 'IdEnquêteur', 'Appareil', 'Statut', 'Demandé le', 'Décidé le', 'Décidé par'], MAX_APPAREILS = 2;
var PJ_SHEET = 'PlafondsJour', PJ_HEADERS = ['IdActivité', 'Formulaire', 'PlafondParJour'];
var OBJ_ALL = '*'; // ligne d'objectif commune à tous les enquêteurs inscrits de l'activité
var SUIVI_FORMS = { usagers: 'Usagers', cogecs: 'COGECS', communaute: 'Communauté' };
var SUIVI_MAXDAYS = 400;

function setupSuivi() { tsheet_(AFF_SHEET, AFF_HEADERS); tsheet_(OBJ_SHEET, OBJ_HEADERS); tsheet_(EFF_SHEET, EFF_HEADERS); tsheet_(PAR_SHEET, PAR_HEADERS); enqSheet_(); Logger.log('Feuilles Affectations, Objectifs, Effectifs, SuiviParams et Enqueteurs prêtes.'); }

function suiviRows_(name, headers) { var sh = tsheet_(name, headers); return rowsOf_(sh, headers.length); }
function suiviReplace_(name, headers, actId, newRows) {
  var sh = tsheet_(name, headers), keep = suiviRows_(name, headers).filter(function (r) { return String(r.v[0]) !== String(actId); }).map(function (r) { return r.v; });
  var all = keep.concat(newRows), last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, headers.length).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, headers.length).setNumberFormat('@').setValues(all);
}
function actById_(id) { var f = null; activitiesRows_().forEach(function (a) { if (a.id === String(id)) f = a; }); return f; }
function assignedTo_(user) { var o = {}; suiviRows_(AFF_SHEET, AFF_HEADERS).forEach(function (r) { if (String(r.v[1]) === String(user)) o[String(r.v[0])] = 1; }); return o; }

/** Accès au suivi : administrateur (tout) ou superviseur d'activité (activités affectées seulement). */
function suiviAccess_(token, actId) {
  var s = token ? findSession_(token) : null;
  if (!s) return { ok: false, error: 'Session invalide ou expirée, reconnectez-vous.' };
  var a = findAccountRow_(s.username);
  if (!a || a.data[8] === 'NON') return { ok: false, error: 'Compte désactivé ou introuvable.' };
  var lvl = String(a.data[2]);
  if (lvl === 'admin') return { ok: true, admin: true, user: s.username, label: String(a.data[6] || s.username) };
  if (lvl === SUP_LEVEL) {
    if (actId && !assignedTo_(s.username)[String(actId)]) return { ok: false, error: "Cette activité ne vous est pas affectée." };
    return { ok: true, admin: false, user: s.username, label: String(a.data[6] || s.username) };
  }
  return { ok: false, error: "Accès réservé à l'administrateur et aux superviseurs d'activité." };
}
function actSummary_(a) { return { id: a.id, name: a.name, level: a.level, cible: a.cible, debut: a.debut, fin: a.fin, forms: a.forms, active: a.active }; }
function handleSuiviList(token) {
  var c = suiviAccess_(token, ''); if (!c.ok) return c;
  var mine = c.admin ? null : assignedTo_(c.user);
  var list = activitiesRows_().filter(function (a) { return !mine || mine[a.id]; }).map(actSummary_);
  list.sort(function (x, y) { return (x.debut || '') < (y.debut || '') ? 1 : (x.debut || '') > (y.debut || '') ? -1 : 0; });
  return { ok: true, admin: c.admin, user: c.user, label: c.label, activities: list };
}
function daysBetween_(a, b) { // nombre de jours entre deux dates yyyy-MM-dd (b - a)
  var x = String(a).match(/^(\d{4})-(\d{2})-(\d{2})$/), y = String(b).match(/^(\d{4})-(\d{2})-(\d{2})$/); if (!x || !y) return 0;
  return Math.round((Date.UTC(+y[1], +y[2] - 1, +y[3]) - Date.UTC(+x[1], +x[2] - 1, +x[3])) / 86400000);
}
function actForms_(act) {
  return (act.level === 'hopital' ? ['usagers', 'communaute'] : ['usagers', 'cogecs']).filter(function (x) { return !act.forms.length || act.forms.indexOf(x) > -1; });
}
/** Objectif commun (par cible) appliqué à TOUS les enquêteurs inscrits. À défaut de ligne commune, reprend la valeur la plus fréquente des anciens objectifs individuels. */
function objGlobal_(actId) {
  var g = {}, has = false, legacy = {};
  suiviRows_(OBJ_SHEET, OBJ_HEADERS).forEach(function (r) {
    if (String(r.v[0]) !== String(actId)) return;
    var n = Number(r.v[3]); if (!(n >= 0)) return; var f = String(r.v[2]);
    if (String(r.v[1]).trim() === OBJ_ALL) { g[f] = n; has = true; } else { (legacy[f] = legacy[f] || {})[n] = (legacy[f][n] || 0) + 1; }
  });
  if (has) return g;
  Object.keys(legacy).forEach(function (f) { var best = null; Object.keys(legacy[f]).forEach(function (n) { if (best === null || legacy[f][n] > legacy[f][best]) best = n; }); g[f] = Number(best); });
  return g;
}
/** Durées minimales (en minutes) fixées par l'administrateur pour valider un questionnaire : durée de saisie et intervalle depuis le précédent. */
function minsOf_(actId) {
  try { var psh = tsheet_(PAR_SHEET, PAR_HEADERS); if (!String(psh.getRange(1, 10).getValue())) psh.getRange(1, 1, 1, PAR_HEADERS.length).setValues([PAR_HEADERS]); } catch (e) {}
  var o = { duree: 0, intervalle: 0, centre: 0, plafondJour: 0, plafond: 0, montantCentre: 0, nbCentres: 0 }; suiviRows_(PAR_SHEET, PAR_HEADERS).forEach(function (r) { if (String(r.v[0]) === String(actId)) { o.duree = Math.max(0, Number(r.v[4]) || 0); o.intervalle = Math.max(0, Number(r.v[5]) || 0); o.centre = Math.max(0, Number(r.v[6]) || 0); o.plafondJour = Math.max(0, Math.round(Number(r.v[7]) || 0)); o.montantCentre = Math.max(0, Number(r.v[8]) || 0); o.nbCentres = Math.max(0, Number(r.v[9]) || 0); } });
  o.plafond = o.plafondJour; // nombre maximal de questionnaires par téléphone, par centre et par jour (valeur fixée par l'administrateur)
  return o;
}
/** Quota attendu par centre et par cible (ex. 60 usagers, 3 COGECS) : fixé par l'administrateur pour chaque activité (zone sanitaire ou hôpital). */
/** Plafond maximal de questionnaires par jour, fixé par l'administrateur pour chaque type de questionnaire (par téléphone et par centre). */
function plafondsOf_(actId) {
  var q = {}; suiviRows_(PJ_SHEET, PJ_HEADERS).forEach(function (r) { if (String(r.v[0]) === String(actId)) { var n = Math.round(Number(r.v[2])); if (n > 0) q[String(r.v[1])] = n; } });
  return q;
}
/** Téléphones d'un enquêteur : 2 maximum sont autorisés automatiquement ; un troisième reste « ATTENTE » jusqu'à la décision de l'administrateur. */
function devCheck_(actId, enqId, dev) {
  dev = String(dev || '').trim(); if (!dev || !enqId) return 'OK';
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var rows = suiviRows_(AP_SHEET, AP_HEADERS), mine = rows.filter(function (r) { return String(r.v[0]) === String(actId) && String(r.v[1]) === String(enqId); }), found = null, nOk = 0;
    mine.forEach(function (r) { if (String(r.v[2]) === dev) found = r; if (String(r.v[3]) === 'AUTORISE') nOk++; });
    if (found) { var st = String(found.v[3]); return st === 'AUTORISE' ? 'OK' : st; }
    var status = nOk < MAX_APPAREILS ? 'AUTORISE' : 'ATTENTE', now = new Date().toISOString(), sh = tsheet_(AP_SHEET, AP_HEADERS);
    sh.getRange(sh.getLastRow() + 1, 1, 1, AP_HEADERS.length).setNumberFormat('@').setValues([[String(actId), String(enqId), dev, status, now, status === 'AUTORISE' ? now : '', status === 'AUTORISE' ? 'automatique' : '']]);
    return status === 'AUTORISE' ? 'OK' : 'ATTENTE';
  } finally { lock.releaseLock(); }
}
function quotasOf_(actId) {
  var q = {}; suiviRows_(QC_SHEET, QC_HEADERS).forEach(function (r) { if (String(r.v[0]) === String(actId)) { var n = Number(r.v[2]); if (n > 0) q[String(r.v[1])] = n; } });
  return q;
}
function montantOf_(actId) {
  var m = null; suiviRows_(PAR_SHEET, PAR_HEADERS).forEach(function (r) { if (String(r.v[0]) === String(actId)) { var n = Number(r.v[1]); if (n >= 0 && String(r.v[1]) !== '') m = n; } });
  return m;
}
function enqColumn_(hm) { var names = ["Identification de l'enquêteur", 'Enqueteur', 'Enquêteur', 'enqueteur']; for (var i = 0; i < names.length; i++) if (hm[names[i]]) return names[i]; return ''; }

/** Calcule tout le suivi d'une activité : réalisé par enquêteur et par jour, objectifs, attendu, personnel. */
function suiviCompute_(act) {
  var today = day_(new Date()), ss = SpreadsheetApp.getActiveSpreadsheet();
  var srcs = (act.level === 'hopital' ? [['usagers', SHEET_USAGERS_H], ['communaute', SHEET_COMMUNAUTE_H]] : [['usagers', SHEET_USAGERS], ['cogecs', SHEET_COGECS]])
    .filter(function (x) { return !act.forms.length || act.forms.indexOf(x[0]) > -1; });
  var enq = {}, order = [], allDays = {}, glob = objGlobal_(act.id), mins = minsOf_(act.id), quotas = quotasOf_(act.id), plafonds = plafondsOf_(act.id);
  // Enquêteurs INSCRITS (nom, prénom, numéro saisis à leur première connexion) : ils forment la liste de suivi, même sans réponse
  enqList_(act.id).forEach(function (x) {
    var k = 'id:' + x.id; enq[k] = { key: k, id: x.id, name: x.prenom, nom: x.nom, numero: x.numero, inscrit: true, byForm: {}, daily: {}, objByForm: {} };
    Object.keys(glob).forEach(function (f) { enq[k].objByForm[f] = glob[f]; }); order.push(k);
  });
  function getE_(name) { // réponses sans enquêteur inscrit (anciennes réponses, code sans activité) : regroupées par nom
    name = String(name || '').replace(/\s+/g, ' ').trim() || 'Non renseigné'; var k = 'n:' + (strip_(name) || 'non renseigne');
    if (!enq[k]) { enq[k] = { key: k, id: '', name: name, nom: '', numero: '', inscrit: false, byForm: {}, daily: {}, objByForm: {} }; order.push(k); }
    return enq[k];
  }
  srcs.forEach(function (sr) {
    var sh = ss.getSheetByName(sr[1]); if (!sh || sh.getLastRow() < 2) return;
    var hm = buildHeaderMap_(sh).map; if (!hm[ACT_ID_COL]) return;
    var tsC = hm[START_COL] ? START_COL : (hm['start'] ? 'start' : ''), teC = hm[END_COL] ? END_COL : (hm['end'] ? 'end' : '');
    var cenC = hm['Structure'] ? 'Structure' : (hm['fs'] ? 'fs' : ''), cenv = cenC ? readColumn_(sh, hm, cenC) : [], dvv = hm[DEV_COL] ? readColumn_(sh, hm, DEV_COL) : [], tsv = tsC ? readColumn_(sh, hm, tsC) : [], tev = teC ? readColumn_(sh, hm, teC) : [];
    var ids = readColumn_(sh, hm, ACT_ID_COL), ec = enqColumn_(hm), names = ec ? readColumn_(sh, hm, ec) : [], days = rowDays_(sh, hm), eids = hm[ENQ_ID_COL] ? readColumn_(sh, hm, ENQ_ID_COL) : [];
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i]) !== act.id) continue;
      var e = (eids[i] && enq['id:' + String(eids[i])]) || getE_(names[i]), d = days[i] || '';
      e.byForm[sr[0]] = (e.byForm[sr[0]] || 0) + 1;
      (e.items = e.items || []).push({ f: sr[0], s: Date.parse(tsv[i]), e: Date.parse(tev[i]), dev: String(dvv[i] || ''), c: strip_(cenv[i]), cl: String(cenv[i] || '').trim() });
      if (d) { e.daily[d] = (e.daily[d] || 0) + 1; allDays[d] = 1; }
    }
  });
  // Validation : durée de saisie ≥ minimum, et intervalle depuis la fin du questionnaire précédent du même enquêteur ≥ minimum.
  // Un questionnaire sans horodatage (anciennes réponses) ne peut pas être évalué : il reste valide.
  order.forEach(function (k) {
    var e = enq[k], its = (e.items || []).slice().sort(function (a, b) { return (isNaN(a.s) ? 0 : a.s) - (isNaN(b.s) ? 0 : b.s); }), prevEnd;
    e.valid = {}; e.rej = {}; e.rejDuree = 0; e.rejInterv = 0; e.rejPlafond = 0; e.centres = {}; e.centresAll = {}; var perDay = {}, devs = {}; prevEnd = {};
    var devInfo = {}, devOrder = []; e.rejets = [];
    its.forEach(function (it) {
      var bad = false, rj = null, dkey = it.dev || '?'; if (it.dev) devs[it.dev] = 1;
      if (!devInfo[dkey]) { devInfo[dkey] = { code: it.dev, total: 0, valid: 0, rej: 0, first: it.s, last: it.e, n: 0 }; devOrder.push(dkey); }
      if (it.c) { var ca = e.centresAll[it.c] = e.centresAll[it.c] || { label: it.cl, tot: {} }; ca.tot[it.f] = (ca.tot[it.f] || 0) + 1; }
      if (!isNaN(it.s) && !isNaN(it.e)) {
        // L'analyse se fait téléphone par téléphone (un même compte peut être utilisé sur plusieurs téléphones, qui collectent en même temps) ;
        // le résultat est ensuite regroupé sur le compte (tableau de bord général). Sans identifiant de téléphone (anciennes réponses), on ne peut pas
        // savoir de quel téléphone il s'agit : l'intervalle n'est pas évalué (on ne mélange jamais deux téléphones inconnus).
        // Seuls les questionnaires USAGERS (zones sanitaires) et USAGERS / COMMUNAUTÉ (hôpitaux) peuvent être rejetés.
        // Tous les autres (COGECS…) ne sont jamais rejetés, quels que soient leur durée ou leur horaire, et ne comptent pas dans l'intervalle d'un téléphone.
        var rejectable = (it.f === 'usagers' || it.f === 'communaute');
        if (rejectable) {
          if (mins.duree > 0 && (it.e - it.s) / 60000 < mins.duree) { bad = true; e.rejDuree++; rj = { m: 'duree', dur: Math.round((it.e - it.s) / 1000), min: mins.duree }; }
          else if (mins.intervalle > 0 && it.dev && prevEnd[it.dev] !== undefined && (it.s - prevEnd[it.dev]) / 60000 < mins.intervalle) { bad = true; e.rejInterv++; rj = { m: 'intervalle', gap: Math.round((it.s - prevEnd[it.dev]) / 1000), prevEnd: new Date(prevEnd[it.dev]).toISOString(), min: mins.intervalle }; }
          if (it.dev) prevEnd[it.dev] = (prevEnd[it.dev] === undefined) ? it.e : Math.max(prevEnd[it.dev], it.e);
          // plafond : un même téléphone ne peut pas valider plus que le maximum par jour fixé par l'administrateur pour ce type de questionnaire, pour toute la journée (tous centres confondus)
          var capF = plafonds[it.f] || 0;
          if (!bad && capF > 0 && it.dev) {
            var pk = it.dev + '|' + Utilities.formatDate(new Date(it.s), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '|' + it.f; // tous centres confondus : un téléphone qui a atteint son maximum du jour ne peut plus compter pour un autre centre
            if ((perDay[pk] || 0) >= capF) { bad = true; e.rejPlafond++; rj = { m: 'plafond', cap: capF }; } else perDay[pk] = (perDay[pk] || 0) + 1;
          }
        }
      }
      var di = devInfo[dkey]; di.total++; if (bad) di.rej++; else di.valid++; if (!isNaN(it.s) && (isNaN(di.first) || it.s < di.first)) di.first = it.s; if (!isNaN(it.e) && (isNaN(di.last) || it.e > di.last)) di.last = it.e;
      if (bad) {
        e.rej[it.f] = (e.rej[it.f] || 0) + 1;
        if (e.rejets.length < 500) { rj.f = it.f; rj.dkey = dkey; rj.c = it.cl; rj.s = new Date(it.s).toISOString(); rj.e = new Date(it.e).toISOString(); e.rejets.push(rj); }
      } else {
        e.valid[it.f] = (e.valid[it.f] || 0) + 1;
        if (it.c) { var ce = e.centres[it.c] = e.centres[it.c] || { label: it.cl, valid: {} }; ce.valid[it.f] = (ce.valid[it.f] || 0) + 1; }
      }
    });
    // Numérotation des téléphones de ce compte (Téléphone 1, 2…) dans l'ordre de première utilisation ; 0 = identifiant inconnu (anciennes réponses)
    var known = devOrder.filter(function (k) { return k !== '?'; }).sort(function (x, y) { return (devInfo[x].first || 0) - (devInfo[y].first || 0); }), nn = 0;
    known.forEach(function (k) { devInfo[k].n = ++nn; });
    e.appareils = known.concat(devInfo['?'] ? ['?'] : []).map(function (k) {
      var d = devInfo[k]; return { n: d.n, code: d.code ? String(d.code).slice(-4) : '', total: d.total, valid: d.valid, rej: d.rej, first: isNaN(d.first) ? '' : new Date(d.first).toISOString(), last: isNaN(d.last) ? '' : new Date(d.last).toISOString() };
    });
    e.rejets.forEach(function (r) { r.tel = devInfo[r.dkey].n; r.code = devInfo[r.dkey].code ? String(devInfo[r.dkey].code).slice(-4) : ''; delete r.dkey; });
    e.nbAppareils = Object.keys(devs).length;
  });
  // période
  var dk = Object.keys(allDays).sort(), start = act.debut || dk[0] || today, endAct = act.fin || '';
  var end = endAct && endAct < today ? endAct : today; if (end < start) end = start;
  var nDays = Math.min(daysBetween_(start, end) + 1, SUIVI_MAXDAYS), days = [], i2;
  for (i2 = 0; i2 < nDays; i2++) days.push(i2 === 0 ? start : dayPlus_(start, i2));
  var totalDays = (act.debut && act.fin) ? Math.max(1, daysBetween_(act.debut, act.fin) + 1) : 0;
  var elapsed = act.debut ? Math.max(0, Math.min(daysBetween_(act.debut, today) + 1, totalDays || 99999)) : 0;
  var formsKeys = srcs.map(function (x) { return x[0]; });
  // Total des questionnaires validés par centre et par cible, tous enquêteurs confondus (plusieurs enquêteurs peuvent collecter dans un même centre)
  var centreTot = {};
  Object.keys(enq).forEach(function (k) { var cs = enq[k].centres || {}; Object.keys(cs).forEach(function (ck) { var t = centreTot[ck] = centreTot[ck] || {}; Object.keys(cs[ck].valid || {}).forEach(function (fk) { t[fk] = (t[fk] || 0) + cs[ck].valid[fk]; }); }); });
  var list = order.map(function (k) {
    var e = enq[k], total = 0, obj = 0, hasObj = false, f;
    // Objectif = quota d'un centre × nombre de centres réellement enquêtés par cet enquêteur (au moins 1 : le centre en cours)
    var qS = 0; formsKeys.forEach(function (fk) { qS += quotas[fk] || 0; });
    var cKeys = Object.keys(e.centresAll || {}), nc = cKeys.length, cd = [];
    if (qS > 0) {
      e.objByForm = {}; formsKeys.forEach(function (fk) { if (quotas[fk]) e.objByForm[fk] = quotas[fk] * Math.max(1, nc); });
      cKeys.forEach(function (ck) {
        var ca = e.centresAll[ck], cv = (e.centres[ck] && e.centres[ck].valid) || {}, all = centreTot[ck] || {}, got = 0, cgot = 0, det = {};
        // Part de l'enquêteur = son propre travail ; si le total du centre dépasse le quota, chacun est payé au prorata (le surplus n'est jamais payé).
        formsKeys.forEach(function (fk) { var q = quotas[fk] || 0, v = cv[fk] || 0, t = all[fk] || 0; if (q > 0) { got += t > q ? v * q / t : v; cgot += Math.min(t, q); det[fk] = { valides: v, total: ca.tot[fk] || 0, quota: q, centre: t }; } });
        var cfull = true; formsKeys.forEach(function (fk) { if ((quotas[fk] || 0) > 0 && (all[fk] || 0) < quotas[fk]) cfull = false; });
        cd.push({ centre: ca.label, cibles: det, complet: cfull, centrePct: Math.round(cgot / qS * 1000) / 10, pct: Math.round(got / qS * 1000) / 10, parts: got, montant: mins.montantCentre > 0 ? Math.round(mins.montantCentre * got / qS) : 0 });
      });
      cd.sort(function (a, b) { return String(a.centre).localeCompare(String(b.centre), 'fr'); });
    }
    formsKeys.forEach(function (fk) { total += e.byForm[fk] || 0; if (e.objByForm[fk] !== undefined) { obj += e.objByForm[fk]; hasObj = true; } });
    var run = 0, cumul = days.map(function (d) { run += e.daily[d] || 0; return run; });
    var before = 0; Object.keys(e.daily).forEach(function (d) { if (d < start) before += e.daily[d]; }); // réponses datées avant le début (comptées au 1er jour)
    if (before && cumul.length) cumul = cumul.map(function (v) { return v + before; });
    var attendu = (hasObj && totalDays) ? Math.round(obj * Math.min(elapsed, totalDays) / totalDays * 10) / 10 : null;
    var restant = hasObj ? Math.max(0, obj - total) : null, joursRest = totalDays ? Math.max(0, totalDays - elapsed) : null;
    var retenu = 0, nbValid = 0, nbRej = 0; formsKeys.forEach(function (fk) { retenu += Math.min(e.valid[fk] || 0, e.objByForm[fk] || 0); nbValid += e.valid[fk] || 0; nbRej += e.rej[fk] || 0; });
    return { nbCentres: nc, centresDetail: cd, valid: e.valid, rej: e.rej, nbValid: nbValid, nbRej: nbRej, rejDuree: e.rejDuree, rejInterv: e.rejInterv, rejPlafond: e.rejPlafond, nbAppareils: e.nbAppareils, appareils: e.appareils, rejets: e.rejets, id: e.id, numero: e.numero, nom: e.nom, inscrit: e.inscrit, retenu: retenu, name: e.name, byForm: e.byForm, objByForm: e.objByForm, total: total, objectif: hasObj ? obj : null, attendu: attendu,
      pct: hasObj && obj > 0 ? Math.round(total / obj * 1000) / 10 : null, ecart: attendu !== null ? Math.round((total - attendu) * 10) / 10 : null,
      rythmeReel: elapsed > 0 ? Math.round(total / elapsed * 10) / 10 : null, rythmeRequis: (restant !== null && joursRest) ? Math.round(restant / joursRest * 10) / 10 : null,
      last: Object.keys(e.daily).sort().pop() || '', cumul: cumul };
  });
  list.sort(function (a, b) { return b.total - a.total || a.name.localeCompare(b.name, 'fr'); });
  var withObj = list.filter(function (e) { return e.objectif !== null; }), avgObj = withObj.length ? withObj.reduce(function (s, e) { return s + e.objectif; }, 0) / withObj.length : null;
  var totals = { realise: 0, objectif: 0, attendu: 0, byForm: {}, nbEnq: list.length, nbInscrits: list.filter(function (e) { return e.inscrit; }).length, nbActifs: list.filter(function (e) { return e.total > 0; }).length };
  formsKeys.forEach(function (fk) { totals.byForm[fk] = { realise: 0, objectif: 0 }; });
  list.forEach(function (e) {
    totals.realise += e.total; if (e.objectif !== null) totals.objectif += e.objectif; if (e.attendu !== null) totals.attendu += e.attendu;
    formsKeys.forEach(function (fk) { totals.byForm[fk].realise += e.byForm[fk] || 0; if (e.objByForm[fk] !== undefined) totals.byForm[fk].objectif += e.objByForm[fk]; });
  });
  totals.attendu = Math.round(totals.attendu * 10) / 10; totals.valides = 0; totals.rejetes = 0; list.forEach(function (e) { totals.valides += e.nbValid; totals.rejetes += e.nbRej; });
  // Paiement (per diem) : montant total ÷ (enquêteurs inscrits × objectif) = tarif par questionnaire ; chacun reçoit son tarif × questionnaires réalisés (plafonnés à l'objectif de chaque cible)
  var montant = montantOf_(act.id), objEach = 0; formsKeys.forEach(function (fk) { objEach += glob[fk] || 0; });
  var nIns = totals.nbInscrits, paiement = { defini: false, montant: montant, nbInscrits: nIns, objectifParEnq: objEach };
  if (montant !== null && nIns > 0 && objEach > 0) {
    var tarif = montant / (nIns * objEach), dus = 0;
    list.forEach(function (e) { if (e.inscrit) { e.aPayer = Math.round(tarif * e.retenu); e.pctPaye = Math.round(e.retenu / objEach * 1000) / 10; dus += e.aPayer; } });
    paiement = { defini: true, montant: montant, nbInscrits: nIns, objectifParEnq: objEach, partParEnq: Math.round(montant / nIns), tarif: Math.round(tarif * 100) / 100, totalDu: dus, reste: Math.round(montant) - dus };
  }
  // Paiement par centre : chaque centre enquêté rapporte le montant du centre s'il atteint le quota de chaque cible, au prorata sinon ;
  // au-delà du quota d'une cible, le surplus n'est pas payé. L'enquêteur est payé pour tous les centres qu'il a faits.
  var qSum = 0; formsKeys.forEach(function (fk) { qSum += quotas[fk] || 0; });
  if (mins.montantCentre > 0 && qSum > 0) {
    var dus2 = 0;
    list.forEach(function (e) {
      var src = enq[e.id ? 'id:' + e.id : 'n:' + (strip_(e.name) || 'non renseigne')]; e.aPayer = 0; var tot = 0;
      (e.centresDetail || []).forEach(function (c) { tot += c.montant; });
      if (e.inscrit) { e.aPayer = tot; dus2 += tot; } e.retenu = 0; (e.centresDetail || []).forEach(function (c) { e.retenu += Math.round(c.parts || 0); });
    });
    var touches = {}, complets = {};
    list.forEach(function (e) { (e.centresDetail || []).forEach(function (c) { var k = strip_(c.centre); touches[k] = 1; if (c.complet) complets[k] = 1; }); });
    paiement = { defini: true, mode: 'centre', montantCentre: mins.montantCentre, quotas: quotas, nbInscrits: nIns, totalDu: dus2, montant: montant,
      nbCentresPrevus: mins.nbCentres, centresTouches: Object.keys(touches).length, centresComplets: Object.keys(complets).length, budgetPrevu: mins.nbCentres * mins.montantCentre };
  }
  var attenduMoyen = (avgObj !== null && totalDays) ? days.map(function (d, i) { var t = Math.min(daysBetween_(act.debut || start, d) + 1, totalDays); return Math.round(avgObj * Math.max(0, t) / totalDays * 10) / 10; }) : null;
  var cumulTotal = days.map(function (d, i) { var s = 0; list.forEach(function (e) { s += e.cumul[i]; }); return s; });
  return { start: start, end: end, today: today, days: days, totalDays: totalDays, elapsed: elapsed, forms: formsKeys, enqueteurs: list, totals: totals,
    avgObjectif: avgObj !== null ? Math.round(avgObj * 10) / 10 : null, mins: mins, quotas: quotas, paiement: paiement, attenduMoyen: attenduMoyen, cumulTotal: cumulTotal, nbObjectifs: withObj.length };
}
/** Point du personnel : effectif attendu (saisi par l'administrateur) et réponses du formulaire personnel pendant la période. */
function suiviPersonnel_(act, c) {
  var eff = suiviRows_(EFF_SHEET, EFF_HEADERS).filter(function (r) { return String(r.v[0]) === act.id; });
  var structs = [], globalEff = 0;
  eff.forEach(function (r) { var n = Number(r.v[2]) || 0, st = String(r.v[1] || '').trim(); if (!st) globalEff = n; else structs.push({ structure: st, effectif: n, repondu: 0, key: strip_(st) }); });
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(act.level === 'hopital' ? SHEET_PERSONNELS_H : SHEET_PERSONNELS), repondu = 0, autres = 0;
  if (sh && sh.getLastRow() > 1) {
    var hm = buildHeaderMap_(sh).map, scol = act.level === 'hopital' ? 'fs' : 'Structure', st = readColumn_(sh, hm, scol), days = rowDays_(sh, hm);
    var ck = act.level === 'hopital' ? strip_(act.cible) : '', hasCk = ck && st.some(function (v) { var k = strip_(v); return k && (k.indexOf(ck) > -1 || ck.indexOf(k) > -1); });
    var from = act.debut || '', to = act.fin || '9999-12-31';
    for (var i = 0; i < st.length; i++) {
      var k = strip_(st[i]); if (!k && !days[i]) continue;
      var d = days[i] || ''; if (from && (!d || d < from)) continue; if (d > to) continue;
      if (hasCk && !(k.indexOf(ck) > -1 || ck.indexOf(k) > -1)) continue;
      repondu++;
      var hit = false; structs.forEach(function (x) { if (!hit && k && x.key === k) { x.repondu++; hit = true; } });
      if (!hit) autres++;
    }
  }
  var sumEff = structs.reduce(function (s, x) { return s + x.effectif; }, 0), effTotal = globalEff || sumEff;
  structs.forEach(function (x) { x.taux = x.effectif > 0 ? Math.round(x.repondu / x.effectif * 1000) / 10 : null; delete x.key; });
  return { effectif: effTotal, repondu: repondu, taux: effTotal > 0 ? Math.round(repondu / effTotal * 1000) / 10 : null, structures: structs, autres: autres, defini: effTotal > 0 };
}
function handleSuiviGet(token, actId) {
  var c = suiviAccess_(token, actId); if (!c.ok) return c;
  var act = actById_(actId); if (!act) return { ok: false, error: 'Activité introuvable.' };
  var r = suiviCompute_(act); r.ok = true; r.activity = actSummary_(act); r.personnel = suiviPersonnel_(act, r); r.generatedAt = new Date().toISOString();
  var sup = []; suiviRows_(AFF_SHEET, AFF_HEADERS).forEach(function (x) { if (String(x.v[0]) === act.id) { var f = findAccountRow_(x.v[1]); if (f) sup.push(String(f.data[6] || f.data[0])); } });
  r.supervisors = sup; r.viewer = c.label;
  return r;
}
/** Configuration d'une activité : enquêteurs inscrits, objectif commun, montant total à payer (administrateur et superviseurs affectés) ;
 *  superviseurs, affectations et effectif du personnel : administrateur seulement. */
/** Téléphones des enquêteurs d'une activité : demandes en attente et téléphones autorisés (administrateur seulement). */
function handleSuiviDevices(token, actId) {
  var c = suiviAccess_(token, actId); if (!c.ok) return c; if (!c.admin) return { ok: false, error: "Réservé à l'administrateur." };
  var names = {}; enqList_(actId).forEach(function (e) { names[e.id] = e; });
  var list = suiviRows_(AP_SHEET, AP_HEADERS).filter(function (r) { return String(r.v[0]) === String(actId); }).map(function (r) {
    var e = names[String(r.v[1])] || {}, dv = String(r.v[2]);
    return { enqId: String(r.v[1]), nom: e.nom || '', prenom: e.prenom || '', numero: e.numero || '', dev: dv, code: dv.slice(-4), statut: String(r.v[3]), demande: String(r.v[4]), decide: String(r.v[5]), par: String(r.v[6]) };
  });
  return { ok: true, max: MAX_APPAREILS, appareils: list };
}
/** L'administrateur autorise, refuse ou retire un téléphone d'un enquêteur. */
function handleSuiviDeviceDecision(token, actId, enqId, dev, decision) {
  var c = suiviAccess_(token, actId); if (!c.ok) return c; if (!c.admin) return { ok: false, error: "Réservé à l'administrateur." };
  var st = decision === 'AUTORISE' ? 'AUTORISE' : decision === 'REFUSE' ? 'REFUSE' : ''; if (!st) return { ok: false, error: 'Décision invalide.' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = tsheet_(AP_SHEET, AP_HEADERS), done = false;
    suiviRows_(AP_SHEET, AP_HEADERS).forEach(function (r) {
      if (String(r.v[0]) === String(actId) && String(r.v[1]) === String(enqId) && String(r.v[2]) === String(dev)) {
        sh.getRange(r.row, 4, 1, 4).setNumberFormat('@').setValues([[st, String(r.v[4]), new Date().toISOString(), c.label || c.user]]); done = true;
      }
    });
    if (!done) return { ok: false, error: 'Téléphone introuvable.' };
    SpreadsheetApp.flush(); return { ok: true };
  } finally { lock.releaseLock(); }
}
function handleSuiviConfig(token, actId) {
  var c = suiviAccess_(token, actId); if (!c.ok) return c;
  var out = { ok: true, admin: c.admin, activities: [], supervisors: [], accountants: [] };
  if (c.admin) {
    var v = accountsSheet_().getDataRange().getValues();
    for (var i = 1; i < v.length; i++) if (String(v[i][2]) === SUP_LEVEL) out.supervisors.push({ username: String(v[i][0]), label: String(v[i][6] || v[i][0]), active: v[i][8] === undefined || v[i][8] === '' || v[i][8] === 'OUI' });
    out.accountants = []; for (var k = 1; k < v.length; k++) if (String(v[k][2]) === COMPTA_LEVEL) out.accountants.push({ username: String(v[k][0]), label: String(v[k][6] || v[k][0]) });
    out.activities = activitiesRows_().map(actSummary_);
  }
  if (actId) {
    var act = actById_(actId); if (!act) return { ok: false, error: 'Activité introuvable.' };
    out.forms = actForms_(act);
    out.enqueteurs = enqList_(act.id).map(function (x) { return { id: x.id, nom: x.nom, prenom: x.prenom, numero: x.numero }; });
    out.objGlobal = objGlobal_(act.id); out.montant = montantOf_(act.id); var mn_ = minsOf_(act.id); out.dureeMin = mn_.duree; out.intervalleMin = mn_.intervalle; out.tailleCentre = mn_.centre; out.plafonds = plafondsOf_(act.id); out.montantCentre = mn_.montantCentre; out.nbCentres = mn_.nbCentres; out.quotas = quotasOf_(act.id);
    if (c.admin) {
      out.assigned = suiviRows_(AFF_SHEET, AFF_HEADERS).filter(function (r) { return String(r.v[0]) === act.id; }).map(function (r) { return String(r.v[1]); });
      out.effectifs = suiviRows_(EFF_SHEET, EFF_HEADERS).filter(function (r) { return String(r.v[0]) === act.id; }).map(function (r) { return { structure: String(r.v[1]), effectif: Number(r.v[2]) || 0 }; });
    }
  }
  return out;
}
function handleSuiviSave(token, actId, cfg) {
  var c = suiviAccess_(token, actId); if (!c.ok) return c;
  var act = actById_(actId); if (!act) return { ok: false, error: 'Activité introuvable.' };
  cfg = cfg || {};
  var num = function (x) { var n = Math.round(Number(String(x).replace(/\s/g, '').replace(',', '.'))); return (n >= 0 && n <= 1000000) ? n : null; };
  var forms = actForms_(act), obj = [], g = cfg.objGlobal || {};
  forms.forEach(function (f) { if (g[f] === undefined || g[f] === null || String(g[f]).trim() === '') return; var n = num(g[f]); if (n !== null) obj.push([act.id, OBJ_ALL, f, String(n)]); });
  var par = [], mt = cfg.montant === undefined || cfg.montant === null ? '' : String(cfg.montant).trim(), m = '';
  if (mt !== '') {
    m = Math.round(Number(mt.replace(/\s/g, '').replace(',', '.')));
    if (!(m >= 0 && m <= 10000000000)) return { ok: false, error: 'Montant total à payer invalide.' };
  }
  var mnum = function (x) { var t = String(x === undefined || x === null ? '' : x).trim().replace(',', '.'); if (t === '') return 0; var n = Number(t); return (n >= 0 && n <= 600) ? Math.round(n * 10) / 10 : null; };
  var dMin = mnum(cfg.dureeMin), iMin = mnum(cfg.intervalleMin);
  if (dMin === null || iMin === null) return { ok: false, error: 'Les durées minimales doivent être comprises entre 0 et 600 minutes.' };
  var tC = Math.round(Number(String(cfg.tailleCentre === undefined || cfg.tailleCentre === null ? '' : cfg.tailleCentre).trim().replace(',', '.')) || 0), pP = 0;
  if (!(tC >= 0 && tC <= 100000) ) return { ok: false, error: 'Taille du centre (0 à 100000) invalide.' };
  var mc = Math.round(Number(String(cfg.montantCentre === undefined || cfg.montantCentre === null ? '' : cfg.montantCentre).trim().replace(/\s/g, '').replace(',', '.')) || 0);
  if (!(mc >= 0 && mc <= 1000000000)) return { ok: false, error: 'Montant par centre invalide.' };
  var qrows = []; forms.forEach(function (f) { var qn = Math.round(Number(String((cfg.quotas || {})[f] === undefined || (cfg.quotas || {})[f] === null ? '' : cfg.quotas[f]).trim().replace(',', '.')) || 0); if (qn > 0 && qn <= 100000) qrows.push([act.id, f, String(qn)]); });
  var prows = []; forms.forEach(function (f) { if (f !== 'usagers' && f !== 'communaute') return; var pn = Math.round(Number(String((cfg.plafonds || {})[f] === undefined || (cfg.plafonds || {})[f] === null ? '' : cfg.plafonds[f]).trim().replace(',', '.')) || 0); if (pn > 100000) pn = 0; if (pn > 0) prows.push([act.id, f, String(pn)]); });
  var nc = Math.round(Number(String(cfg.nbCentres === undefined || cfg.nbCentres === null ? '' : cfg.nbCentres).trim().replace(',', '.')) || 0);
  if (!(nc >= 0 && nc <= 100000)) return { ok: false, error: 'Nombre de centres invalide.' };
  if (m !== '' || dMin > 0 || iMin > 0 || tC > 0 || mc > 0 || nc > 0) par.push([act.id, m === '' ? '' : String(m), new Date().toISOString(), c.user, String(dMin), String(iMin), String(tC), String(pP), String(mc), String(nc)]);
  var aff = null, eff = null;
  if (c.admin) {
    var okSup = {}; var v = accountsSheet_().getDataRange().getValues(); for (var i = 1; i < v.length; i++) if (String(v[i][2]) === SUP_LEVEL) okSup[String(v[i][0])] = 1;
    aff = []; var seen = {}; (cfg.supervisors || []).forEach(function (u) { u = String(u); if (okSup[u] && !seen[u]) { seen[u] = 1; aff.push([act.id, u]); } });
    eff = []; var seenE = {};
    (cfg.effectifs || []).forEach(function (e) {
      var st = cleanName_(e.structure), n = num(e.effectif); if (n === null) return; var k = strip_(st); if (seenE[k]) return; seenE[k] = 1; eff.push([act.id, st, String(n)]);
    });
  }
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (aff) suiviReplace_(AFF_SHEET, AFF_HEADERS, act.id, aff);
    suiviReplace_(OBJ_SHEET, OBJ_HEADERS, act.id, obj); suiviReplace_(PAR_SHEET, PAR_HEADERS, act.id, par); suiviReplace_(QC_SHEET, QC_HEADERS, act.id, qrows); suiviReplace_(PJ_SHEET, PJ_HEADERS, act.id, prows);
    if (eff) suiviReplace_(EFF_SHEET, EFF_HEADERS, act.id, eff);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return { ok: true, superviseurs: aff ? aff.length : null, objectifs: obj.length, effectifs: eff ? eff.length : null, montant: (par.length && par[0][1] !== '') ? Number(par[0][1]) : null, dureeMin: dMin, intervalleMin: iMin, tailleCentre: tC, plafonds: cfg.plafonds || {}, montantCentre: mc, nbCentres: nc };
}
function handleCreateSupervisor(token, username, label, password) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  username = String(username || '').trim(); label = cleanName_(label) || username;
  if (!/^[A-Za-z0-9._-]{3,30}$/.test(username)) return { ok: false, error: "Identifiant : 3 à 30 caractères (lettres, chiffres, point, tiret, souligné), sans espace." };
  if (String(password || '').length < 8) return { ok: false, error: 'Le mot de passe doit contenir au moins 8 caractères.' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (findAccountRow_(username)) return { ok: false, error: 'Cet identifiant existe déjà.' };
    var sh = accountsSheet_(), row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 9).setNumberFormat('@').setValues([[username, sha256Hex_(username + ':' + password), SUP_LEVEL, '', '', '', label, new Date().toISOString(), 'OUI']]);
    setPlainPwd_(row, password);
    return { ok: true, username: username, label: label };
  } finally { lock.releaseLock(); }
}
var COMPTA_LEVEL = 'comptable';
function handleCreateAccountant(token, username, label, password) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  username = String(username || '').trim(); label = cleanName_(label) || username;
  if (!/^[A-Za-z0-9._-]{3,30}$/.test(username)) return { ok: false, error: "Identifiant : 3 à 30 caractères (lettres, chiffres, point, tiret, souligné), sans espace." };
  if (String(password || '').length < 8) return { ok: false, error: 'Le mot de passe doit contenir au moins 8 caractères.' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (findAccountRow_(username)) return { ok: false, error: 'Cet identifiant existe déjà.' };
    var sh = accountsSheet_(), row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, 9).setNumberFormat('@').setValues([[username, sha256Hex_(username + ':' + password), COMPTA_LEVEL, '', '', '', label, new Date().toISOString(), 'OUI']]);
    setPlainPwd_(row, password);
    return { ok: true, username: username, label: label };
  } finally { lock.releaseLock(); }
}
/** Accès comptabilité : le comptable (lecture et paiement) ; l'administrateur peut consulter. */
function comptaAccess_(token) {
  var s = token ? findSession_(token) : null; if (!s) return { ok: false, error: 'Session invalide ou expirée, reconnectez-vous.' };
  var a = findAccountRow_(s.username); if (!a || a.data[8] === 'NON') return { ok: false, error: 'Compte désactivé ou introuvable.' };
  var lvl = String(a.data[2]);
  if (lvl === COMPTA_LEVEL) return { ok: true, canPay: true, user: s.username, label: String(a.data[6] || s.username) };
  if (lvl === 'admin') return { ok: true, canPay: false, user: s.username, label: String(a.data[6] || s.username) };
  return { ok: false, error: 'Accès réservé au comptable.' };
}
/** Liste des activités et, pour l'activité choisie, des enquêteurs avec le montant qui leur revient, l'état de la facture et du paiement. */
function handleComptaGet(token, actId) {
  var c = comptaAccess_(token); if (!c.ok) return c;
  var acts = activitiesRows_().map(function (a) { return { id: a.id, name: a.name, level: a.level, cible: a.cible, debut: a.debut, fin: a.fin }; });
  acts.sort(function (x, y) { return (x.debut || '') < (y.debut || '') ? 1 : (x.debut || '') > (y.debut || '') ? -1 : 0; });
  var out = { ok: true, canPay: c.canPay, label: c.label, activities: acts, enqueteurs: [] };
  if (!actId) return out;
  var act = actById_(actId); if (!act) return { ok: false, error: 'Activité introuvable.' };
  var enqById = {}; enqList_(actId).forEach(function (e) { enqById[e.id] = e; });
  var comp = suiviCompute_(act), pays = {};
  suiviRows_(PAY_SHEET, PAY_HEADERS).forEach(function (r) { if (String(r.v[0]) === String(actId) && String(r.v[3]) === 'PAYE') pays[String(r.v[1])] = { montant: Number(r.v[2]) || 0, date: String(r.v[4]), par: String(r.v[5]), momo: String(r.v[6]) }; });
  comp.enqueteurs.forEach(function (e) {
    if (!e.id || !enqById[e.id]) return;
    var en = enqById[e.id], p = pays[e.id] || null;
    out.enqueteurs.push({ id: en.id, nom: en.nom, prenom: en.prenom, numero: en.numero, momo: en.momo, montant: Math.round(Number(e.aPayer) || 0), valide: en.valide, paye: !!p, paiement: p });
  });
  out.enqueteurs.sort(function (x, y) { return (x.nom + ' ' + x.prenom).localeCompare(y.nom + ' ' + y.prenom); });
  return out;
}
/** Le comptable coche « paiement effectué » (ou annule) après le transfert MTN MoMo. Le montant est recalculé par le serveur. */
function handleComptaPay(token, actId, enqId, paid) {
  var c = comptaAccess_(token); if (!c.ok) return c; if (!c.canPay) return { ok: false, error: 'Seul le comptable peut enregistrer un paiement.' };
  var en = enqById_(actId, enqId); if (!en) return { ok: false, error: 'Enquêteur introuvable.' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ex = paymentOf_(actId, enqId), sh = tsheet_(PAY_SHEET, PAY_HEADERS);
    if (!paid) { if (ex) sh.getRange(ex.row, 4, 1, 3).setNumberFormat('@').setValues([['ANNULE', new Date().toISOString(), c.label]]); SpreadsheetApp.flush(); return { ok: true }; }
    if (ex) return { ok: true, already: true };
    if (!en.valide) return { ok: false, error: "La facture de cet enquêteur n'a pas encore été validée par l'administrateur." };
    if (!cleanMomo_(en.momo)) return { ok: false, error: "Cet enquêteur n'a pas renseigné de numéro MTN MoMo valide (8 chiffres)." };
    var act = actById_(actId); if (!act) return { ok: false, error: 'Activité introuvable.' };
    var me = null; suiviCompute_(act).enqueteurs.forEach(function (e) { if (String(e.id) === String(enqId)) me = e; });
    var montant = Math.round(Number(me && me.aPayer) || 0); if (montant <= 0) return { ok: false, error: 'Aucun montant à payer pour cet enquêteur.' };
    sh.getRange(sh.getLastRow() + 1, 1, 1, PAY_HEADERS.length).setNumberFormat('@').setValues([[String(actId), String(enqId), String(montant), 'PAYE', new Date().toISOString(), c.label, en.momo]]);
    SpreadsheetApp.flush(); return { ok: true, montant: montant };
  } finally { lock.releaseLock(); }
}
function handleSetSupervisorPassword(token, username, password) {
  var c = requireAdminSession_(token); if (!c.ok) return c;
  if (String(password || '').length < 8) return { ok: false, error: 'Le mot de passe doit contenir au moins 8 caractères.' };
  var f = findAccountRow_(username); if (!f || (String(f.data[2]) !== SUP_LEVEL && String(f.data[2]) !== COMPTA_LEVEL)) return { ok: false, error: 'Compte introuvable.' };
  var sh = accountsSheet_(); sh.getRange(f.row, 2).setValue(sha256Hex_(String(f.data[0]).trim() + ':' + password)); setPlainPwd_(f.row, password); sh.getRange(f.row, 8).setValue(new Date().toISOString());
  invalidateSessionsForUser_(String(f.data[0]).trim());
  return { ok: true };
}
