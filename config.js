// Configuration + accès par code QR (session enquêteur)
const API_URL = "https://script.google.com/macros/s/AKfycbwjB6aW_ksTbTJgf2n7-VXh7d2FPuTVQov8m6ZijLN5OPPSj4kEFqU1w08HowBs-GcWAw/exec";
const QR_KEY = "qr_token";

function callApi(action, data, password) {
  if (!API_URL || API_URL.indexOf("COLLEZ_ICI") !== -1) {
    return Promise.reject(new Error("API_URL non configurée : modifiez config.js"));
  }
  if (/^submit/.test(action) && data && !data._cid) data._cid = (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : Date.now() + '-' + Math.random().toString(16).slice(2);
  return fetch(API_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: action, data: data, password: password, qr: localStorage.getItem(QR_KEY) || "" })
  }).then(function (r) {
    if (!r.ok) throw new Error("Réponse HTTP " + r.status);
    return r.json();
  }).then(function (j) {
    if (j && j.error === "QR_INVALID" && window.qrExit) window.qrExit();
    return j;
  });
}

// Les formulaires exigent un code scanné ; il est revérifié toutes les 60 s
// et à chaque envoi : si l'administrateur change le code, l'accès est coupé.
(function () {
  if (!/Form(?!Personnels)[A-Za-z]*\.html$/.test(location.pathname)) return; // le formulaire Personnel soignant reste libre d'accès
  window.qrExit = function () {
    localStorage.removeItem(QR_KEY);
    location.replace("Acces.html?expired=1");
  };
  if (!localStorage.getItem(QR_KEY)) { location.replace("Acces.html"); return; }
  function verify() {
    if (!navigator.onLine) return;
    callApi("checkQr", { token: localStorage.getItem(QR_KEY) }).then(function (r) {
      if (r && r.ok && !r.valid) window.qrExit();
    }).catch(function () {});
  }
  verify();
  setInterval(verify, 60000);
})();

if ('serviceWorker' in navigator) { window.addEventListener('load', function () { navigator.serviceWorker.register('sw.js').catch(function () {}); }); }
