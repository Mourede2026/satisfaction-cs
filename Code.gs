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
      delete d[ACT_COL]; delete d[ACT_ID_COL];
      var act = need ? activityByCode_(body.qr) : null;
      if (act) {
        var fk = FORM_KEY_[action];
        if (act.forms.length && fk && act.forms.indexOf(fk) === -1) return json_({ ok: false, error: 'FORM_NOT_ALLOWED' });
        d[ACT_COL] = act.name; d[ACT_ID_COL] = act.id;
      }
    }
    if (isSubmit && dejaRecu_(cid)) return json_({ ok: true, duplicate: true });
    if (action === "submitUsagers") out = appendRecord_(SHEET_USAGERS, d);
    else if (action === "submitPersonnels") out = appendRecord_(SHEET_PERSONNELS, d);
    else if (action === "submitCogecs") out = appendRecord_(SHEET_COGECS, d);
    else if (action === "submitUsagersHopital") out = appendRecord_(SHEET_USAGERS_H, d, 'dept');
    else if (action === "submitPersonnelHopital") out = appendRecord_(SHEET_PERSONNELS_H, d, 'dept');
    else if (action === "submitCommunauteHopital") out = appendRecord_(SHEET_COMMUNAUTE_H, d, 'dept');
    else if (action === "checkQr") out = { ok: true, valid: qrTokenValid_(d.token, d.level === 'hopital' ? ['enqueteur_hopital'] : d.level === 'zone' ? ['enqueteur'] : QR_ROLES) };
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
    var cols_ = [DATE_COLUMN]; if (data[ACT_COL]) cols_.push(ACT_COL, ACT_ID_COL);
    cols_.forEach(function (nm) {
      if (hm.map[nm]) return;
      sheet.getRange(1, hm.lastCol + 1).setValue(nm);
      try { sheet.getRange(2, hm.lastCol + 1, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@'); } catch (e) {}
      hm = buildHeaderMap_(sheet);
    });
    data[DATE_COLUMN] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
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
  // vérifié une fois, converti en empreinte (hash) puis effacé de la feuille.
  var pc = pwdColIndex_(), plain = pc < 0 || f.data[pc] === undefined || f.data[pc] === null ? '' : String(f.data[pc]).trim();
  if (plain) {
    if (String(password).trim() !== plain) { loginEchec_(username); return { ok: false, error: 'Identifiant ou mot de passe incorrect.' }; }
    var sh = accountsSheet_();
    sh.getRange(f.row, 2).setValue(sha256Hex_(String(f.data[0]).trim() + ':' + String(password).trim()));
    sh.getRange(f.row, pc + 1).clearContent();
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
  if (sha256Hex_(s.username + ':' + oldPassword) !== f.data[1]) return { ok: false, error: 'Ancien mot de passe incorrect.' };
  var sh = accountsSheet_();
  sh.getRange(f.row, 2).setValue(sha256Hex_(s.username + ':' + newPassword));
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
  if (sha256Hex_(s.username + ':' + currentPassword) !== f.data[1]) return { ok: false, error: 'Mot de passe incorrect.' };
  newUsername = (newUsername || '').toString().trim();
  if (!newUsername) return { ok: false, error: 'Identifiant vide.' };
  if (newUsername === s.username) return { ok: false, error: "C'est déjà votre identifiant actuel." };
  if (findAccountRow_(newUsername)) return { ok: false, error: 'Cet identifiant est déjà utilisé par un autre compte.' };
  var sh = accountsSheet_();
  sh.getRange(f.row, 1).setValue(newUsername);
  sh.getRange(f.row, 2).setValue(sha256Hex_(newUsername + ':' + currentPassword));
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
    out.push({ row: i + 1, id: String(v[i][0]), name: String(v[i][1]), level: String(v[i][2]), cible: String(v[i][3] || ''), forms: csvList_(v[i][4]), communes: csvList_(v[i][5]), debut: day_(v[i][6]), fin: day_(v[i][7]), code: String(v[i][8]).trim(), active: v[i][9] !== 'NON' });
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
  var communes = level === 'zone' ? (p.communes || []).map(function (x) { return String(x).replace(/,/g, ' ').trim(); }).filter(Boolean) : [];
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
  if (!f) sh.appendRow(['admin', sha256Hex_('admin:' + pwd), 'admin', '', '', '', 'Administrateur', '', 'OUI']);
  else { sh.getRange(f.row, 2).setValue(sha256Hex_('admin:' + pwd)); sh.getRange(f.row, 3).setValue('admin'); sh.getRange(f.row, 9).setValue('OUI'); }
  Logger.log('Identifiant : admin — Nouveau mot de passe : ' + pwd);
}

/* ====================================================================
 * PÉRIODES DE COLLECTE, DONNÉES HÔPITAUX, ÉVOLUTION
 * ==================================================================== */
var PERIODS_SHEET_NAME = 'Periodes', PERIODS_HEADERS = ['Id', 'Libellé', 'Début', 'Fin', 'Actif'];
var DATE_COLUMN = 'DateEnvoi';
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
