// ===== PATCH POUR Code.gs — accès QR enquêteurs =====
// 1) COLLER ce bloc n'importe où dans Code.gs.
var QR_ROLES = ['enqueteur', 'personnel', 'enqueteur_hopital', 'personnel_hopital'];
var QR_GUARD = {
  submitUsagers: ['enqueteur'], submitPersonnels: ['enqueteur'], submitCogecs: ['enqueteur'],
  submitUsagersHopital: ['enqueteur_hopital'], submitPersonnelHopital: ['enqueteur_hopital'],
  submitCommunauteHopital: ['enqueteur_hopital'], submitActeursComHopital: ['enqueteur_hopital']
};
function qrTokenValid_(token, roles) {
  token = (token || '').toString().trim();
  if (!token) return false;
  for (var i = 0; i < roles.length; i++) {
    var row = findQrRow_(roles[i]);
    if (row && String(row.token).trim() === token) return true;
  }
  return false;
}
/** À exécuter UNE FOIS (▶) : crée les jetons des 4 rôles s'ils manquent. */
function ensureQrRoles() {
  var sh = getOrCreateSheet_(QRTOKENS_SHEET_NAME, ['Rôle', 'Token', 'GénéréLe']);
  QR_ROLES.forEach(function (r) {
    if (!findQrRow_(r)) sh.appendRow([r, Utilities.getUuid(), new Date().toISOString()]);
  });
}
// 2) REMPLACER la fonction handleGetQrToken existante par celle-ci (désormais réservée à l'admin).
function handleGetQrToken(sessionToken, role) {
  var check = requireAdminSession_(sessionToken);
  if (!check.ok) return check;
  if (QR_ROLES.indexOf(role) === -1) return { ok: false, error: 'Rôle inconnu.' };
  var row = findQrRow_(role);
  if (row) return { ok: true, token: row.token };
  var t = Utilities.getUuid();
  getOrCreateSheet_(QRTOKENS_SHEET_NAME, ['Rôle', 'Token', 'GénéréLe']).appendRow([role, t, new Date().toISOString()]);
  return { ok: true, token: t };
}
// 3) Dans doPost :
//   a) juste APRÈS la ligne  var action = body.action;  AJOUTER :
//        var need = QR_GUARD[action];
//        if (need && !qrTokenValid_(body.qr, need)) {
//          return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'QR_INVALID' }))
//                 .setMimeType(ContentService.MimeType.JSON);
//        }
//   b) REMPLACER la ligne getQrToken par :
//        else if (action === "getQrToken") out = handleGetQrToken(body.data.token, body.data.role);
//   c) AJOUTER avant le dernier  else out = { ok: false, error: "Action inconnue...
//        else if (action === "checkQr") out = { ok: true, valid: qrTokenValid_(body.data && body.data.token, QR_ROLES) };
// 4) Dans handleRefreshQrToken ET handleSetQrCode, remplacer
//      if (role !== 'enqueteur' && role !== 'personnel')
//    par
//      if (QR_ROLES.indexOf(role) === -1)
// 5) Déployer > Gérer les déploiements > Nouvelle version, puis exécuter ensureQrRoles une fois.
