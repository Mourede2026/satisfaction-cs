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
  localStorage.removeItem(QR_KEYS[level]); localStorage.removeItem(QR_KEYS[level] + "_label"); localStorage.removeItem(QR_KEYS[level] + "_act");
}
// Activité scannée (nom, cible, formulaires autorisés, communes, dates) : mémorisée avec le code
function getActivity(level) { try { return JSON.parse(localStorage.getItem(QR_KEYS[level || pageLevel()] + "_act") || "null"); } catch (e) { return null; } }
function setActivity(level, a) { if (a) localStorage.setItem(QR_KEYS[level] + "_act", JSON.stringify(a)); else localStorage.removeItem(QR_KEYS[level] + "_act"); }
// Clé de formulaire (pour restreindre ce qu'une activité autorise)
function formKeyOf(file) { return /Cogecs/i.test(file) ? "cogecs" : /Communaute/i.test(file) ? "communaute" : /Usagers/i.test(file) ? "usagers" : ""; }
function activityAllows(a, key) { return !a || !a.forms || !a.forms.length || a.forms.indexOf(key) > -1; }

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
  // Une activité qui n'autorise pas ce formulaire : retour à la liste des formulaires de l'activité
  var act0 = getActivity(level);
  if (act0 && !activityAllows(act0, formKeyOf(here))) { location.replace("index.html#" + level); return; }
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

// Bandeau « Activité » en haut des formulaires d'enquête + restriction de la commune quand l'activité en désigne
document.addEventListener("DOMContentLoaded", function () {
  if (!/Form(?!Personnel)[A-Za-z]*\.html$/.test(location.pathname)) return;
  var a = getActivity(); if (!a) return;
  var d = document.createElement("div");
  d.style.cssText = "background:#e6f6ee;color:#1a7f4b;padding:8px 14px;font-size:.85rem;text-align:center";
  d.textContent = "📅 Activité : " + a.name + (a.cible ? " — " + a.cible : "");
  var anchor = document.getElementById("status-banner") || document.querySelector("main");
  if (anchor && anchor.parentNode) anchor.parentNode.insertBefore(d, anchor);
});
// À appeler après chaque rendu du formulaire (zone sanitaire) : limite la liste des communes à celles de l'activité ; une seule = présélectionnée
function applyActivityToForm() {
  var a = getActivity("zone"); if (!a || !a.communes || !a.communes.length) return;
  var sel = document.getElementById("sel_commune"); if (!sel) return;
  for (var i = sel.options.length - 1; i >= 1; i--) if (a.communes.indexOf(sel.options[i].value) === -1) sel.remove(i);
  if (a.communes.length === 1 && sel.options.length === 2) { sel.selectedIndex = 1; if (typeof onCommuneChange === "function") onCommuneChange(sel.value); }
}

if ('serviceWorker' in navigator) { window.addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () {}); }); }

// Lien direct « solo » (?solo=1) : le formulaire du personnel s'ouvre seul, sans lien de retour vers l'accueil
(function () {
  if (!/[?&]solo=1/.test(location.search) || !/FormPersonnel/.test(location.pathname)) return;
  document.addEventListener("DOMContentLoaded", function () {
    var a = document.querySelectorAll("a.back");
    for (var i = 0; i < a.length; i++) a[i].parentNode.removeChild(a[i]);
  });
})();
