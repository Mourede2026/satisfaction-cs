// Configuration + accès enquêteur par code (un code par niveau : zone sanitaire / hôpital)
const API_URL_DEFAULT = "https://script.google.com/macros/s/AKfycbwbFgMu71sQhLWfLZSo7_UKLkCzy1sNnfju0Z6o0q9EFubPmXC101fxpY125wP8NKu18w/exec";
// L'adresse du serveur peut être corrigée depuis la page Administration (stockée sur l'appareil)
// Nouveau déploiement : si l'adresse d'origine a changé, on abandonne l'ancienne adresse corrigée à la main sur l'appareil
(function () {
  try {
    if (localStorage.getItem("api_default_seen") !== API_URL_DEFAULT) {
      localStorage.removeItem("api_url");
      localStorage.setItem("api_default_seen", API_URL_DEFAULT);
    }
  } catch (e) {}
})();
function apiUrl() { return localStorage.getItem("api_url") || API_URL_DEFAULT; }
const API_URL = API_URL_DEFAULT;
const QR_KEYS = { zone: "qr_token_zone", hopital: "qr_token_hopital" };
const QR_KEY = QR_KEYS.zone; // compatibilité

// Migration de l'ancien code unique (avant la séparation zone / hôpital)
(function () {
  try {
    var t = localStorage.getItem("qr_token");
    if (t) {
      var lv = localStorage.getItem("qr_role") === "enqueteur_hopital" ? "hopital" : "zone";
      if (!localStorage.getItem(QR_KEYS[lv])) localStorage.setItem(QR_KEYS[lv], t);
      localStorage.removeItem("qr_token"); localStorage.removeItem("qr_role");
    }
  } catch (e) {}
})();

// Niveau de la page courante : les formulaires hôpital ont "Hopital" dans leur nom
function pageLevel() { return /Hopital/i.test(location.pathname) ? "hopital" : "zone"; }
function getQr(level) { return localStorage.getItem(QR_KEYS[level || pageLevel()]) || ""; }
function getQrLabel(level) { return localStorage.getItem(QR_KEYS[level || pageLevel()] + "_label") || ""; }
function setQr(level, token, label) {
  localStorage.setItem(QR_KEYS[level], token);
  localStorage.setItem(QR_KEYS[level] + "_label", label || "");
}
function clearQr(level) {
  localStorage.removeItem(QR_KEYS[level]); localStorage.removeItem(QR_KEYS[level] + "_label");
}

function callApi(action, data, password) {
  if (!apiUrl() || apiUrl().indexOf("COLLEZ_ICI") !== -1) {
    return Promise.reject(new Error("API_URL non configurée : modifiez config.js"));
  }
  if (/^submit/.test(action) && data && !data._cid) data._cid = (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : Date.now() + '-' + Math.random().toString(16).slice(2);
  return fetch(apiUrl(), {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: action, data: data, password: password, qr: getQr() })
  }).catch(function () {
    throw new Error("Serveur injoignable : vérifiez la connexion Internet et le déploiement Apps Script (accès « Tout le monde »).");
  }).then(function (r) {
    if (!r.ok) throw new Error("Réponse HTTP " + r.status + " du serveur.");
    return r.text();
  }).then(function (t) {
    var j;
    try { j = JSON.parse(t); } catch (e) {
      throw new Error("Réponse invalide : dans Apps Script, déployez en « Exécuter en tant que : Moi » et « Accès : Tout le monde », puis vérifiez API_URL dans config.js.");
    }
    if (j && j.error === "QR_INVALID" && window.qrExit) window.qrExit();
    return j;
  });
}

// Les formulaires d'enquête (hors Personnel, en accès libre) exigent le code du niveau.
// Il est saisi une seule fois, revérifié toutes les 60 s et à chaque envoi :
// si l'administrateur le change ou le désactive, l'accès est coupé (les réponses en attente sont conservées).
(function () {
  if (!/Form(?!Personnel)[A-Za-z]*\.html$/.test(location.pathname)) return;
  var level = pageLevel(), here = location.pathname.split("/").pop();
  var to = function (extra) { return "Acces.html?level=" + level + "&next=" + encodeURIComponent(here) + (extra || ""); };
  window.qrExit = function () { clearQr(level); location.replace(to("&expired=1")); };
  if (!getQr(level)) { location.replace(to()); return; }
  function verify() {
    if (!navigator.onLine) return;
    callApi("checkQr", { token: getQr(level), level: level }).then(function (r) {
      if (r && r.ok && !r.valid) window.qrExit();
    }).catch(function () {});
  }
  verify();
  setInterval(verify, 60000);
})();

if ('serviceWorker' in navigator) { window.addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () {}); }); }
