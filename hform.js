/* Moteur des formulaires hôpital : construit le questionnaire à partir de window.HFORM
   (schéma issu des fichiers XLSForm), gère conditions d'affichage, listes en cascade,
   envoi, file d'attente hors connexion et position GPS. */
(function () {
  var S = window.HFORM, STORE_LABELS = true; // true : le libellé du choix est enregistré ; false : son code
  var QUEUE_KEY = 'queue_' + S.key, ans = {}, byName = {}, lastPos = null;
  S.fields.forEach(function (f) { byName[f.name] = f; });
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

  /* ---------- logique ---------- */
  function condTrue(c) {
    var v = ans[c.f], has = Array.isArray(v) ? v.indexOf(c.v) > -1 : v === c.v;
    return c.op === '=' ? has : !has;
  }
  function relOk(rel) { return !rel || rel.some(function (g) { return g.every(condTrue); }); }
  function visible(f) { var g = f.g != null ? S.groups[f.g] : null; return (!g || relOk(g.rel)) && relOk(f.rel); }
  function opts(f) {
    var l = S.lists[f.list] || [];
    if (!f.flt) return l;
    var p = ans[f.flt.field];
    return p ? l.filter(function (o) { return o[f.flt.col] === p; }) : [];
  }
  function labelOf(f, n) { var o = (S.lists[f.list] || []).filter(function (x) { return x.n === n; })[0]; return o ? o.l : n; }
  function empty(v) { return v === undefined || v === null || v === '' || (Array.isArray(v) && !v.length); }

  /* ---------- rendu ---------- */
  function ctl(f) {
    var n = f.name, v = ans[n], h = '';
    if (f.type === 'select_one' || f.type === 'select_multiple') {
      var o = opts(f), multi = f.type === 'select_multiple';
      if (!multi && (f.flt || o.length > 7)) {
        h = '<select data-n="' + n + '"><option value="">' + (f.flt && !ans[f.flt.field] ? '-- Choisir d\'abord la question précédente --' : '-- Choisir --') + '</option>' +
          o.map(function (x) { return '<option value="' + esc(x.n) + '"' + (v === x.n ? ' selected' : '') + '>' + esc(x.l) + '</option>'; }).join('') + '</select>';
      } else {
        h = '<div class="options' + (f.app === 'horizontal' ? ' row' : '') + '">' + o.map(function (x, i) {
          var on = multi ? (v || []).indexOf(x.n) > -1 : v === x.n;
          return '<label class="opt' + (on ? ' selected' : '') + '"><input type="' + (multi ? 'checkbox' : 'radio') + '" name="r_' + n + '" data-n="' + n + '" value="' + esc(x.n) + '"' + (on ? ' checked' : '') + '><span>' + esc(x.l) + '</span></label>';
        }).join('') + '</div>';
      }
    } else if (f.type === 'text') {
      h = f.app === 'multiline' ? '<textarea data-n="' + n + '">' + esc(v) + '</textarea>' : '<input type="text" data-n="' + n + '" value="' + esc(v) + '">';
    } else if (f.type === 'integer') {
      var c = f.con || {};
      h = '<input type="number" inputmode="numeric" data-n="' + n + '" value="' + esc(v) + '"' + (c.min != null ? ' min="' + c.min + '"' : '') + (c.max != null ? ' max="' + c.max + '"' : '') + '>';
      if (c.min != null && c.max != null) h += '<div class="hint">Entre ' + c.min + ' et ' + c.max + '</div>';
      else if (c.max != null) h += '<div class="hint">Maximum : ' + c.max + '</div>';
    } else if (f.type === 'geopoint') {
      h = '<button type="button" class="secondary" data-gps="' + n + '">📍 Enregistrer la position GPS</button><div class="hint" id="gps_' + n + '">' + (v ? '✅ Position enregistrée (précision ' + Math.round(+v.split(' ')[3]) + ' m)' : 'Aucune position enregistrée') + '</div>';
    }
    return h;
  }
  function renderCtl(f) { var c = $('c_' + f.name); if (c && f.type !== 'note') c.innerHTML = ctl(f); }
  function field(f) {
    var star = f.req ? '<span class="req"> *</span>' : '';
    if (f.type === 'note') return '<div class="field note" id="q_' + f.name + '"><p>' + esc(f.label) + '</p></div>';
    return '<div class="field" id="q_' + f.name + '"><label class="qlabel">' + esc(f.label) + star + '</label>' +
      (f.hint && f.type !== 'geopoint' ? '<div class="hint" style="margin:-4px 0 6px">' + esc(f.hint) + '</div>' : '') +
      (f.hint && f.type === 'geopoint' ? '<div class="hint" style="margin:-4px 0 6px">' + esc(f.hint) + '</div>' : '') +
      '<div id="c_' + f.name + '">' + ctl(f) + '</div><div class="ferr" id="e_' + f.name + '"></div></div>';
  }
  function render() {
    var used = {}, html = '';
    S.fields.forEach(function (f) { if (f.g != null) used[f.g] = 1; });
    S.groups.forEach(function (g, i) { if (!used[i] && g.title) html += '<div class="group intro" id="g_' + i + '"><p>' + esc(g.title) + '</p></div>'; });
    var cur = -2, open = false;
    S.fields.forEach(function (f) {
      var g = f.g == null ? -1 : f.g;
      if (g !== cur) {
        if (open) html += '</div>';
        html += '<div class="group" id="g_' + (g < 0 ? 'x' + S.fields.indexOf(f) : g) + '"><h2>' + esc(g < 0 ? '📍 Identification' : S.groups[g].title) + '</h2>';
        cur = g; open = true;
      }
      html += field(f);
    });
    if (open) html += '</div>';
    $('fieldsRoot').innerHTML = html;
    recompute();
  }
  function cascade(n) { S.fields.forEach(function (g) { if (g.flt && g.flt.field === n) { delete ans[g.name]; renderCtl(g); cascade(g.name); } }); }
  function recompute() {
    for (var i = 0; i < 8; i++) {
      var changed = false;
      S.fields.forEach(function (f) {
        var vis = visible(f), w = $('q_' + f.name);
        if (w) w.classList.toggle('hidden', !vis);
        if (!vis && ans[f.name] !== undefined) { delete ans[f.name]; renderCtl(f); cascade(f.name); changed = true; }
      });
      if (!changed) break;
    }
    S.groups.forEach(function (g, k) { var e = $('g_' + k); if (e) e.classList.toggle('hidden', !relOk(g.rel)); });
  }

  /* ---------- événements ---------- */
  var root = $('fieldsRoot');
  root.addEventListener('change', function (e) {
    var t = e.target, n = t.getAttribute('data-n'); if (!n) return; var f = byName[n];
    if (t.type === 'radio') ans[n] = t.value;
    else if (t.type === 'checkbox') {
      var set = (ans[n] || []).slice(), i = set.indexOf(t.value);
      if (t.checked && i < 0) set.push(t.value); if (!t.checked && i > -1) set.splice(i, 1);
      var order = opts(f).map(function (o) { return o.n; });
      set.sort(function (a, b) { return order.indexOf(a) - order.indexOf(b); });
      if (set.length) ans[n] = set; else delete ans[n];
    } else if (t.tagName === 'SELECT') { if (t.value) ans[n] = t.value; else delete ans[n]; }
    else return;
    if (t.type === 'radio' || t.type === 'checkbox') renderCtl(f);
    cascade(n); recompute();
  });
  root.addEventListener('input', function (e) {
    var t = e.target, n = t.getAttribute('data-n'); if (!n || (t.tagName !== 'INPUT' && t.tagName !== 'TEXTAREA') || t.type === 'radio' || t.type === 'checkbox') return;
    var f = byName[n]; if (t.value === '') delete ans[n]; else ans[n] = t.value;
    if (f.type === 'integer') t.style.borderColor = badInt(f) ? '#c0392b' : '#dfe4ea';
    recompute();
  });
  root.addEventListener('click', function (e) {
    var b = e.target.closest('[data-gps]'); if (!b) return;
    var n = b.getAttribute('data-gps'), info = $('gps_' + n);
    if (!navigator.geolocation) { info.textContent = '⚠️ GPS non disponible sur cet appareil.'; return; }
    info.textContent = 'Recherche de la position…'; b.disabled = true;
    navigator.geolocation.getCurrentPosition(function (p) {
      b.disabled = false; lastPos = p;
      ans[n] = p.coords.latitude + ' ' + p.coords.longitude + ' ' + (p.coords.altitude || 0) + ' ' + (p.coords.accuracy || 0);
      info.textContent = '✅ Position enregistrée (précision ' + Math.round(p.coords.accuracy || 0) + ' m)';
    }, function () { b.disabled = false; info.textContent = '⚠️ Position impossible. Activez la localisation et réessayez.'; }, { enableHighAccuracy: true, timeout: 20000, maximumAge: 0 });
  });

  /* ---------- validation et envoi ---------- */
  function badInt(f) {
    var v = ans[f.name]; if (empty(v)) return false; var n = Number(v), c = f.con || {};
    return isNaN(n) || (c.min != null && n < c.min) || (c.max != null && n > c.max);
  }
  function validate() {
    var first = null;
    S.fields.forEach(function (f) {
      var w = $('q_' + f.name), er = $('e_' + f.name); if (!w || f.type === 'note') return;
      w.style.outline = ''; if (er) er.textContent = '';
      if (w.classList.contains('hidden')) return;
      var msg = '';
      if (f.req && empty(ans[f.name])) msg = 'Réponse obligatoire.';
      else if (f.type === 'integer' && badInt(f)) msg = f.cmsg || 'Valeur hors limites.';
      if (msg) { w.style.outline = '2px solid #c0392b'; if (er) er.textContent = msg; if (!first) first = w; }
    });
    return first;
  }
  function out(f, v) {
    if (f.type === 'select_one') return STORE_LABELS ? labelOf(f, v) : v;
    if (f.type === 'select_multiple') return STORE_LABELS ? v.map(function (x) { return labelOf(f, x); }).join(', ') : v.join(' ');
    if (f.type === 'integer') return Number(v);
    return v;
  }
  function payload() {
    var d = {};
    S.fields.forEach(function (f) { if (f.type !== 'note' && !empty(ans[f.name])) d[f.name] = out(f, ans[f.name]); });
    d.today = new Date().toISOString().slice(0, 10);
    var g = S.fields.filter(function (f) { return f.type === 'geopoint' && ans[f.name]; })[0], p;
    if (g) { p = ans[g.name].split(' '); d['_start-geopoint_latitude'] = +p[0]; d['_start-geopoint_longitude'] = +p[1]; d['_start-geopoint_altitude'] = +p[2]; d['_start-geopoint_precision'] = +p[3]; }
    else if (lastPos) { d['_start-geopoint_latitude'] = lastPos.coords.latitude; d['_start-geopoint_longitude'] = lastPos.coords.longitude; d['_start-geopoint_altitude'] = lastPos.coords.altitude || ''; d['_start-geopoint_precision'] = lastPos.coords.accuracy || ''; }
    return d;
  }
  window.onSubmit = function () {
    $('errorMsg').style.display = 'none';
    var bad = validate();
    if (bad) { $('errorMsg').style.display = 'block'; bad.scrollIntoView({ behavior: 'smooth', block: 'center' }); return; }
    var btn = $('submitBtn'); btn.disabled = true; btn.textContent = 'Enregistrement...';
    var data = payload();
    function done() { btn.disabled = false; btn.textContent = '✅ Enregistrer la réponse'; }
    if (!navigator.onLine) { queue(data); toast('Hors connexion : réponse sauvegardée localement.'); reset(); done(); return; }
    callApi(S.fn, data).then(function (res) {
      if (res && res.ok) { toast('✅ Réponse enregistrée avec succès.'); reset(); done(); badge(); }
      else if (res && res.error === 'QR_INVALID') { queue(data); done(); }
      else { queue(data); toast('Connexion instable : réponse sauvegardée localement.'); reset(); done(); }
    }).catch(function () { queue(data); toast('Connexion instable : réponse sauvegardée localement.'); reset(); done(); });
  };
  function getQ() { try { return JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]'); } catch (e) { return []; } }
  function setQ(q) { localStorage.setItem(QUEUE_KEY, JSON.stringify(q)); }
  function queue(d) { var q = getQ(); q.push(d); setQ(q); badge(); }
  function badge() {
    var q = getQ(), el = $('qCount'), b = $('status-banner');
    if (q.length) { el.style.display = 'inline-block'; el.textContent = q.length; } else el.style.display = 'none';
    if (!navigator.onLine) { b.className = 'offline'; b.textContent = '⚠️ Hors connexion — les réponses seront sauvegardées localement.'; }
    else if (q.length) { b.className = 'queued'; b.textContent = q.length + ' réponse(s) en attente de synchronisation.'; }
    else { b.className = ''; b.textContent = ''; }
  }
  window.syncQueue = function (manual) {
    var q = getQ();
    if (!q.length) { if (manual) toast('Aucune réponse en attente.'); return; }
    if (!navigator.onLine) { if (manual) toast('Toujours hors connexion.'); return; }
    (function next() {
      if (!q.length) { setQ(q); badge(); if (manual) toast('✅ Synchronisation terminée.'); return; }
      callApi(S.fn, q[0]).then(function (r) {
        if (r && r.ok) { q.shift(); setQ(q); badge(); next(); } else { setQ(q); badge(); if (manual) toast('Synchronisation interrompue.'); }
      }).catch(function () { setQ(q); badge(); if (manual) toast('Synchronisation interrompue (connexion).'); });
    })();
  };
  function reset() { ans = {}; render(); window.scrollTo({ top: 0, behavior: 'smooth' }); }
  function toast(m) { var t = $('toast'); t.textContent = m; t.style.display = 'block'; setTimeout(function () { t.style.display = 'none'; }, 3000); }
  window.addEventListener('online', function () { badge(); window.syncQueue(false); });
  window.addEventListener('offline', badge);
  document.title = S.title; $('ftitle').textContent = S.emoji + ' ' + S.title;
  render(); badge(); window.syncQueue(false);
  if (navigator.geolocation) navigator.geolocation.getCurrentPosition(function (p) { lastPos = p; }, function () {}, { timeout: 8000, maximumAge: 60000 });
})();
