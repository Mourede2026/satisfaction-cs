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
  submitUsagers: ['enqueteur'], submitPersonnels: ['enqueteur'], submitCogecs: ['enqueteur'],
  submitUsagersHopital: ['enqueteur_hopital'], submitPersonnelHopital: ['enqueteur_hopital'],
  submitCommunauteHopital: ['enqueteur_hopital'], submitActeursComHopital: ['enqueteur_hopital']
};

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(e) { return json_({ ok: true, message: "API Satisfaction Zones Sanitaires en ligne. Utilisez POST." }); }

function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents), action = body.action, d = body.data || {};
    var need = QR_GUARD[action], cid = d._cid;
    if (need) {
      if (!qrTokenValid_(body.qr, need)) return json_({ ok: false, error: 'QR_INVALID' });
      if (dejaRecu_(cid)) return json_({ ok: true, duplicate: true });
    }
    if (action === "submitUsagers") out = appendRecord_(SHEET_USAGERS, d);
    else if (action === "submitPersonnels") out = appendRecord_(SHEET_PERSONNELS, d);
    else if (action === "submitCogecs") out = appendRecord_(SHEET_COGECS, d);
    else if (action === "submitUsagersHopital") out = appendRecord_(SHEET_USAGERS_H, d, 'dept');
    else if (action === "submitPersonnelHopital") out = appendRecord_(SHEET_PERSONNELS_H, d, 'dept');
    else if (action === "submitCommunauteHopital") out = appendRecord_(SHEET_COMMUNAUTE_H, d, 'dept');
    else if (action === "submitActeursComHopital") out = appendRecord_(SHEET_ACTEURS_H, d, 'dept');
    else if (action === "checkQr") out = { ok: true, valid: qrTokenValid_(d.token, QR_ROLES) };
    else if (action === "getDashboardData") out = getDashboardData(d.token);
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
    if (need && out && out.ok) noterRecu_(cid);
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
var SHEET_USAGERS_H = "🏥 USAGERS_HOPITAL", SHEET_PERSONNELS_H = "🏥 PERSONNEL_HOPITAL", SHEET_COMMUNAUTE_H = "🏥 COMMUNAUTE_HOPITAL", SHEET_ACTEURS_H = "🏥 ACTEURS_COM_HOPITAL";
var HEADERS_USAGERS_H = ["dept", "zone", "fs", "Veuillez_enregistrer_sition_au_moins_90", "Enqueteur", "Sujet_enqu_t", "Hospitalisation", "A1a", "Comment_avez_vous_tr_acc_der_l_h_pital_", "Avez_vous_rencontr_re_v_hicule_ou_moto_", "L_acc_s_aux_b_timent_mobilit_r_duite_", "Une_fois_l_int_rie_neaux_d_orientation_", "La_propret_et_le_co_ils_au_rendez_vous_", "Avez_vous_des_suggestions_part", "Si_oui_lesquelles", "Avez_vous_des_diffic_au_sein_de_l_h_pital", "A1", "_3_Qu_est_ce_que_vous_n_avez_pas_aim", "A2", "A3", "_6_Qu_est_ce_que_vous_n_avez_pas_aim", "A3a", "A4", "_9_Qu_est_ce_que_vous_n_avez_pas_aim", "A4a", "A5", "_12_Qu_est_ce_que_vous_n_avez_pas_aim", "A5a", "A6", "_15_Qu_est_ce_que_vous_n_avez_pas_aim", "A6a", "C1", "_2_Dites_nous_un_peu_comment_a_a_t", "C2", "D1", "_5_Dites_nous_un_peu_comment_a_a_t", "D2", "E1", "B1", "B2", "B3", "G1", "G", "G2", "F1", "F2", "F3", "C3", "E2", "E3", "E4", "A7", "A8", "D2_Precise_001", "couleur_tenue", "C4", "C5", "H1", "H2", "Avant_qu_une_d_cisio_vous_t_consult_e_", "Avez_vous_pu_exprime_les_soins_propos_s_", "Vous_est_il_d_j_arriv_de_ref", "Si_oui_votre_d_cisi_ression_ni_jugement_", "Lors_de_votre_prise_en_charge_", "Avez_vous_effectivement_d_sign", "Si_oui_cette_person_essionnels_de_sant_", "Lors_de_votre_prise_en_charge", "Lors_de_votre_prise_en_charge_001", "Lors_de_votre_prise_en_charge__001", "H3", "H4", "H5", "H6", "H7", "C6", "H8", "AA1", "D1_Precise", "AA2", "D2_Precise", "AA3", "D3_Precise", "AA4", "D4_Precise", "I1", "Si_non_pourquoi", "Si_oui_pourquoi", "I2", "Si_non_pourquoi_001", "Si_oui_pourquoi_001", "I3", "I4", "I5", "Age", "Sexe", "departement", "commune", "Profession", "Instruction", "Raison_Choix", "Avez_vous_autre_chose_dire", "today", "_start-geopoint_latitude", "_start-geopoint_longitude", "_start-geopoint_altitude", "_start-geopoint_precision"];
var HEADERS_PERSONNELS_H = ["dept", "zs", "fs", "Formulaire_renseign_par", "Choisir_l_identifiant", "age", "Sexe", "Q1", "Q2", "Q3", "Avez_vous_re_u_une_f_evoirs_des_patients_", "Vous_sentez_vous_sou_evoirs_des_patients_", "Avez_vous_connaissance_des_dro", "Ces_droits_et_devoir_votre_tablissement_", "Pensez_vous_que_vos_onnement_de_travail_", "Avez_vous_des_difficult_s_dans", "Quelles_sont_les_pri_cice_de_vos_devoirs_", "Q4", "Q5", "Q6", "Q7", "Q8", "Recevez_vous_les_inf_d_lais_raisonnables_", "Quels_sont_les_canau_ut_ce_qui_s_applique", "Q9", "Les_coll_gues_de_vot_protocoles_de_soins_", "Les_coll_gues_assure_esoins_des_patients_", "Les_coll_gues_collab_ontinuit_des_soins_", "Q10", "Q11", "Q12", "Q13", "Q14", "Q15", "Q16", "Q17", "Q18", "Q19", "Si_oui_de_quelle_mani_re", "Quelle_autre_mani_re", "Quelles_sont_les_dif_cc_s_l_information", "Q20", "Q21", "Q22", "Vous_pensez_que_l_av_chnique_de_mes_soins", "Vous_pensez_que_l_im_d_erreurs_m_dicales", "Je_re_ois_r_guli_rem_ction_de_mon_service", "Les_plaintes_des_pat_difier_nos_pratiques", "Le_manque_de_temps_e_des_retours_patients", "Les_retours_des_pati_utiles_cliniquement", "Pouvez_vous_citer_un_exemple_r", "Si_oui_pr_ciser", "Selon_vous_quels_so_l_h_pital_de_Pob_", "_9_Quels_sont_les_pri_antage_les_patients_", "Q23", "Q24"];
var HEADERS_COMMUNAUTE_H = ["dept", "zone", "Formation_Sanitaire", "Enqueteur", "Comment_avez_vous_tr_acc_der_l_h_pital_", "Avez_vous_rencontr_re_v_hicule_ou_moto_", "L_acc_s_aux_b_timent_mobilit_r_duite_", "Une_fois_l_int_rie_neaux_d_orientation_", "La_propret_et_le_co_ils_au_rendez_vous_", "Avez_vous_des_suggestions_part", "Si_oui_lesquelles", "Comment_jugez_vous_e_accueil_au_portail_", "Comment_jugez_vous_e_l_dans_les_services_", "Comment_jugez_vous_e_au_service_de_sortie", "Comment_trouvez_vous_les_dirigeants_", "Comment_appr_ciez_vo_es_agents_examiner", "Comment_appr_ciez_vo_diguer_de_bons_soins", "Comment_appr_ciez_vo_e_bonne_prescription", "Comment_appr_ciez_vo_niquer_sur_les_actes", "Pensez_vous_que_les_la_prise_en_charge_", "Comment_trouvez_vous_rapport_avec_le_mal", "Pensez_vous_que_les_ous_servir_au_besoin", "Les_m_dicaments_pres_ibles_dans_le_centre", "Les_m_dicaments_non_pharmacies_proches_", "Comment_trouvez_vous_ifs_des_m_dicaments_", "Comment_trouvez_vous_ents_hospitalisation", "Comment_trouvez_vous_la_cour_de_l_h_pital", "Comment_trouvez_vous_alle_de_consultation", "Comment_trouvez_vous_les_toilettes", "Comment_trouvez_vous_la_buanderie", "Pensez_vous_que_les_onn_es_sur_les_soins", "Pensez_vous_qu_on_a_est_pas_souvent_seul", "Avez_vous_acc_s_l_eau", "Vous_a_t_on_une_fois_des_soins_domicile", "Vous_est_il_jamais_a_r_des_agents_de_sant", "Des_agents_vous_ont_eux_directement_vers", "Pensez_vous_que_les_h_pital_g_rent_bien", "Pensez_vous_que_dans_u_sein_de_l_h_pital_", "Si_vous_devrez_reven_s_motivants_ce_choix", "Age", "Sexe", "Profession", "Instruction", "Avez_vous_autre_chose_dire"];
var HEADERS_ACTEURS_H = ["dept", "zone", "fs", "Enqueteur", "Que_pensez_vous_de_l_ette_structure_fs_", "Comment_appr_ciez_vo_ette_structure_fs_", "Selon_vous_les_soignant_respe", "Si_Oui_en_quoi", "Si_Non_en_quoi", "Les_usagers_de_cette_des_services_offerts", "Que_proposez_vous_ette_structure_fs_"];
function setupFeuillesHopitaux() {
  [[SHEET_USAGERS_H, HEADERS_USAGERS_H], [SHEET_PERSONNELS_H, HEADERS_PERSONNELS_H], [SHEET_COMMUNAUTE_H, HEADERS_COMMUNAUTE_H], [SHEET_ACTEURS_H, HEADERS_ACTEURS_H]]
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
function aggReport_(sheetName, tName, cats, keyHeader, withDemo, withCommune) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sheet = ss.getSheetByName(sheetName), tSheet = ss.getSheetByName(tName);
  var hm = buildHeaderMap_(sheet).map, keys = readColumn_(sheet, hm, keyHeader), n = keys.length;
  var cs = computeCategoryScores_(tSheet, cats, n), names = cats.map(function (c) { return c.name; });
  var ages = withDemo ? readColumn_(sheet, hm, "Âge") : [], sexes = withDemo ? readColumn_(sheet, hm, "Sexe") : [];
  var coms = withCommune ? readColumn_(sheet, hm, "Commune") : [], by = {};
  for (var i = 0; i < n; i++) {
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
function buildCommuneReport_(s, t, c) { var r = aggReport_(s, t, c, "Commune", true, false); return { categories: r.categories, byCommune: r.res }; }
function buildStructureReport_(s, t, c) { var r = aggReport_(s, t, c, "Structure", false, true); return { categories: r.categories, byStructure: r.res }; }
function buildEnqueteurReport_(s, t, c) { var r = aggReport_(s, t, c, "Identification de l'enquêteur", false, false); return { categories: r.categories, byEnqueteur: r.res }; }

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

/** Données du tableau de bord : réservé aux comptes connectés (jeton de session). */
function getDashboardData(sessionToken) {
  var session = sessionToken ? findSession_(sessionToken) : null, acc = session ? findAccountRow_(session.username) : null;
  var account = acc ? accountPublicView_(acc.data) : null;
  if (!account || !account.active) return { ok: false, error: "Authentification requise (connectez-vous avec votre compte)." };
  var errors = [];
  function safe_(label, fn, fb) { try { return fn(); } catch (e) { errors.push(label + " : " + e.message); return fb; } }
  var ss = SpreadsheetApp.getActiveSpreadsheet(), U = ss.getSheetByName(SHEET_USAGERS), C = ss.getSheetByName(SHEET_COGECS), P = ss.getSheetByName(SHEET_PERSONNELS);
  var tU = ss.getSheetByName(T_USAGERS), tC = ss.getSheetByName(T_COGECS), tP = ss.getSheetByName(T_PERSONNELS);
  var uH = U ? buildHeaderMap_(U).map : {}, cH = C ? buildHeaderMap_(C).map : {}, pH = P ? buildHeaderMap_(P).map : {};
  var uCom = safe_("uCom", function () { return readColumn_(U, uH, "Commune"); }, []);
  var uStr = safe_("uStr", function () { return readColumn_(U, uH, "Structure"); }, []);
  var uEnq = safe_("uEnq", function () { return readColumn_(U, uH, "Identification de l'enquêteur"); }, []);
  var uSc = safe_("uSc", function () { return computeRowScoresForCategories_(tU, CATEGORY_MAP.usagers, uCom.length); }, []);
  var cCom = safe_("cCom", function () { return readColumn_(C, cH, "Commune"); }, []);
  var cStr = safe_("cStr", function () { return readColumn_(C, cH, "Structure"); }, []);
  var cSc = safe_("cSc", function () { return computeRowScoresForCategories_(tC, CATEGORY_MAP.cogecs, cCom.length); }, []);
  var pCom = safe_("pCom", function () { return readColumn_(P, pH, "Commune"); }, []);
  var pStr = safe_("pStr", function () { return readColumn_(P, pH, "Structure"); }, []);
  var pSc = safe_("pSc", function () { return computeRowScoresForCategories_(tP, CATEGORY_MAP.personnels, pCom.length); }, []);
  var nz = function (a) { return a.filter(function (v) { return v !== ""; }).length; };
  var uN = nz(uCom), cN = nz(cCom), pN = nz(pCom), emptyC = { categories: [], byCommune: {} };
  var result = {
    ok: true, counts: { usagers: uN, cogecs: cN, personnels: pN, total: uN + cN + pN },
    usagers: { byCommune: safe_("u.com", function () { return groupWithScore_(uCom, uSc); }, []), byStructure: safe_("u.str", function () { return groupWithScore_(uStr, uSc); }, []), byEnqueteur: safe_("u.enq", function () { return groupWithScore_(uEnq, uSc); }, []) },
    cogecs: { byCommune: safe_("c.com", function () { return groupWithScore_(cCom, cSc); }, []), byStructure: safe_("c.str", function () { return groupWithScore_(cStr, cSc); }, []) },
    personnels: { byCommune: safe_("p.com", function () { return groupWithScore_(pCom, pSc); }, []), byStructure: safe_("p.str", function () { return groupWithScore_(pStr, pSc); }, []) },
    global: {
      byCible: [{ label: "Usagers", count: uN, avgScore: safe_("avgU", function () { return avgOf_(uSc); }, null) }, { label: "Personnel soignant", count: pN, avgScore: safe_("avgP", function () { return avgOf_(pSc); }, null) }, { label: "COGECS", count: cN, avgScore: safe_("avgC", function () { return avgOf_(cSc); }, null) }],
      byCommune: safe_("g.com", function () { return groupWithScore_(uCom.concat(cCom, pCom), uSc.concat(cSc, pSc)); }, [])
    },
    communeReport: {
      usagers: safe_("cr.u", function () { return buildCommuneReport_(SHEET_USAGERS, T_USAGERS, CATEGORY_MAP.usagers); }, emptyC),
      personnels: safe_("cr.p", function () { return buildCommuneReport_(SHEET_PERSONNELS, T_PERSONNELS, CATEGORY_MAP.personnels); }, emptyC),
      cogecs: safe_("cr.c", function () { return buildCommuneReport_(SHEET_COGECS, T_COGECS, CATEGORY_MAP.cogecs); }, emptyC)
    },
    structureReport: {
      usagers: safe_("sr.u", function () { return buildStructureReport_(SHEET_USAGERS, T_USAGERS, CATEGORY_MAP.usagers); }, { categories: [], byStructure: {} }),
      personnels: safe_("sr.p", function () { return buildStructureReport_(SHEET_PERSONNELS, T_PERSONNELS, CATEGORY_MAP.personnels); }, { categories: [], byStructure: {} }),
      cogecs: safe_("sr.c", function () { return buildStructureReport_(SHEET_COGECS, T_COGECS, CATEGORY_MAP.cogecs); }, { categories: [], byStructure: {} })
    },
    enqueteurReport: {
      usagers: safe_("er.u", function () { return buildEnqueteurReport_(SHEET_USAGERS, T_USAGERS, CATEGORY_MAP.usagers); }, { categories: [], byEnqueteur: {} }),
      cogecs: safe_("er.c", function () { return buildEnqueteurReport_(SHEET_COGECS, T_COGECS, CATEGORY_MAP.cogecs); }, { categories: [], byEnqueteur: {} })
    },
    gpsByDay: safe_("gps", function () { return getEnqueteurPositions_(); }, {}),
    errors: errors, generatedAt: new Date().toISOString(), account: account
  };
  return applyScopeToResult_(result, account);
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
function findAccountRow_(username) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME);
  if (!sh) return null;
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === username) return { row: i + 1, data: v[i] };
  return null;
}
function accountPublicView_(d) { return { username: d[0], level: d[2], departement: d[3], zone: d[4], commune: d[5], label: d[6], active: (d[8] === undefined || d[8] === '' || d[8] === 'OUI') }; }
function createSession_(username) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET_NAME), v = sh.getDataRange().getValues(), now = Date.now();
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
  if (sha256Hex_(username + ':' + password) !== f.data[1]) { loginEchec_(username); return { ok: false, error: 'Identifiant ou mot de passe incorrect.' }; }
  return { ok: true, token: createSession_(username), account: accountPublicView_(f.data) };
}
function requireAdminSession_(token) {
  var s = token ? findSession_(token) : null;
  if (!s) return { ok: false, error: 'Session invalide ou expirée, reconnectez-vous.' };
  var a = findAccountRow_(s.username);
  if (!a || a.data[2] !== 'admin' || a.data[8] === 'NON') return { ok: false, error: 'Réservé au compte administrateur.' };
  return { ok: true, acc: a };
}
function accountsSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACCOUNTS_SHEET_NAME); }
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
  for (var i = 0; i < roles.length; i++) { var r = findQrRow_(roles[i]); if (r && String(r.token).trim() === token) return true; }
  return false;
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
  for (var i = 0; i < QR_ROLES.length; i++) { var r = findQrRow_(QR_ROLES[i]); if (r && String(r.token).trim() === cand) return { ok: true, role: QR_ROLES[i] }; }
  return { ok: false, error: "Ce code ne correspond à aucune cible active." };
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
