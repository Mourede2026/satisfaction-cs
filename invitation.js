/* invitation.js — Message + image d'invitation pour l'enquête de satisfaction du PERSONNEL
   Utilisé par AdminQR.html (section Activités). Le message et l'image dépendent du niveau
   (zone sanitaire / hôpital), de la structure, de la période et du lien du formulaire personnel. */
(function () {
  var JOURS = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];
  var MOIS = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
  var MOIS_C = ['Janv', 'Févr', 'Mars', 'Avr', 'Mai', 'Juin', 'Juil', 'Août', 'Sept', 'Oct', 'Nov', 'Déc'];
  var SLOGAN = '« Un agent satisfait, c’est un service renforcé et des patients mieux pris en charge. »';
  var FAM = '"Arial Narrow","Liberation Sans Narrow",Arial,Helvetica,sans-serif';
  var st = { src: { level: 'zone', cible: '', debut: '', fin: '', name: '' }, cfg: {}, edited: false, timer: null, ready: false };

  function $(i) { return document.getElementById(i); }
  function two(n) { return ('0' + n).slice(-2); }
  function pd(s) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(s || '')) return null;
    var p = s.split('-'); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2], 12));
  }
  function one(d, mn, withMonth, withYear) {
    return JOURS[d.getUTCDay()] + ' ' + two(d.getUTCDate()) + (withMonth ? ' ' + mn[d.getUTCMonth()] : '') + (withYear ? ' ' + d.getUTCFullYear() : '');
  }
  // Période : « Lundi 18 au Vendredi 29 Mai 2026 » (mois et année affichés une seule fois si identiques)
  function rangeText(debut, fin, mn) {
    var a = pd(debut), b = pd(fin);
    if (a && b) {
      if (+a === +b) return one(a, mn, 1, 1);
      if (a.getUTCFullYear() !== b.getUTCFullYear()) return one(a, mn, 1, 1) + ' au ' + one(b, mn, 1, 1);
      if (a.getUTCMonth() !== b.getUTCMonth()) return one(a, mn, 1, 0) + ' au ' + one(b, mn, 1, 1);
      return one(a, mn, 0, 0) + ' au ' + one(b, mn, 1, 1);
    }
    if (a) return 'à partir du ' + one(a, mn, 1, 1);
    if (b) return 'jusqu’au ' + one(b, mn, 1, 1);
    return '';
  }
  function dateSentence(s) {
    var a = pd(s.debut), b = pd(s.fin);
    if (a && b) return 'Elle se déroulera du *' + rangeText(s.debut, s.fin, MOIS) + '.*';
    if (a) return 'Elle débutera le *' + one(a, MOIS, 1, 1) + '.*';
    if (b) return 'Elle se déroulera jusqu’au *' + one(b, MOIS, 1, 1) + '.*';
    return '';
  }
  function year() {
    var d = pd(st.src.fin) || pd(st.src.debut);
    return d ? d.getUTCFullYear() : new Date().getFullYear();
  }
  function zoneName(c) { return (c || '').trim().replace(/^(la\s+)?(zone\s+sanitaire|z\.?\s?s\.?)\s+(de\s+la\s+|de\s+|du\s+|d[’']\s?)?/i, ''); }
  function hospName(c) {
    c = (c || '').trim(); if (!c) return '';
    if (/^hz\b/i.test(c)) return 'Hôpital de zone de ' + c.replace(/^hz\b\.?\s*(de\s+|d[’']\s?)?/i, '');
    if (/^h[oô]pital/i.test(c)) return c.replace(/^hopital/i, 'Hôpital');
    return 'Hôpital ' + c;
  }
  function structLabel() {
    var s = st.src;
    if (s.level === 'hopital') return (hospName(s.cible) || '[nom de l’hôpital]');
    return 'Zone Sanitaire ' + (zoneName(s.cible) || '[nom de la zone sanitaire]');
  }
  function formUrl() {
    var k = st.src.level === 'hopital' ? 'hop' : 'cs';
    try { return window.lkUrl(k); } catch (e) { return ''; }
  }

  /* ---------- Message ---------- */
  function buildMsg() {
    var s = st.src, y = year(), url = ($('inv-url').value || '').trim() || formUrl();
    var intro = s.level === 'hopital'
      ? 'de l’' + (hospName(s.cible) || '[nom de l’hôpital]') + ' au titre de l’année ' + y
      : 'de la Zone Sanitaire ' + (zoneName(s.cible) || '[nom de la zone sanitaire]');
    var ds = dateSentence(s);
    return 'Bonjour chers tous.\n' +
      'Dans le cadre de l’enquête de satisfaction des Usagers ' + intro + ', l’Etablissement *MOUREDE & Associés* vous invite à prendre part à une enquête destinée au personnel. ' +
      'Celle-ci vise à recueillir votre avis sur les conditions de travail et l’organisation des services au cours de l’année ' + y + '. ' +
      '_L’enquête est entièrement anonyme : aucune information permettant d’identifier un agent ne sera collectée._' + (ds ? ' ' + ds : '') + '\n' +
      '👉 Pour y participer, veuillez cliquer sur le lien suivant : *' + url + '* ou scanner le *Qr code* dans l’image d’invitation.\n' +
      'Merci pour votre participation.';
  }

  /* ---------- Image d'invitation (canevas) ---------- */
  function loadImg(src) {
    return new Promise(function (res) {
      if (!src) { res(null); return; }
      var i = new Image(); i.onload = function () { res(i); }; i.onerror = function () { res(null); }; i.src = src;
    });
  }
  function fitImg(cx, img, x, y, w, h, al) {
    if (!img) return;
    var r = Math.min(w / img.width, h / img.height), iw = img.width * r, ih = img.height * r;
    var dx = al === 'r' ? x + w - iw : al === 'l' ? x : x + (w - iw) / 2;
    cx.drawImage(img, dx, y + (h - ih) / 2, iw, ih);
  }
  function font(size, bold, italic) { return (italic ? 'italic ' : '') + (bold ? 'bold ' : '') + size + 'px ' + FAM; }
  function fitFont(cx, text, maxW, start, min, bold, italic) {
    var s = start; cx.font = font(s, bold, italic);
    while (s > min && cx.measureText(text).width > maxW) { s -= 2; cx.font = font(s, bold, italic); }
    return s;
  }
  function wrap(cx, t, max) {
    var w = String(t).split(' '), l = [], cur = '';
    w.forEach(function (x) { var tt = cur ? cur + ' ' + x : x; if (cx.measureText(tt).width > max && cur) { l.push(cur); cur = x; } else cur = tt; });
    if (cur) l.push(cur); return l;
  }
  function rrect(cx, x, y, w, h, r) {
    cx.beginPath(); cx.moveTo(x + r, y); cx.arcTo(x + w, y, x + w, y + h, r); cx.arcTo(x + w, y + h, x, y + h, r);
    cx.arcTo(x, y + h, x, y, r); cx.arcTo(x, y, x + w, y, r); cx.closePath();
  }
  function partnerSrc() {
    var c = st.cfg;
    if (c.partner === 'msv') return 'inv/msv.png';
    if (c.partner === 'custom' && c.partnerLogo) return c.partnerLogo;
    return '';
  }
  function drawBanner() {
    var c = st.cfg, pSrc = partnerSrc(), hasP = !!pSrc, sc = (c.logoScale || 100) / 100, ps = (c.pScale || 100) / 100, showQr = c.qr !== false;
    var slogan = $('inv-slogan-on').checked ? (($('inv-slogan').value || '').trim() || SLOGAN) : '';
    var W = 1600, H = slogan ? 590 : 520;
    return Promise.all([loadImg('inv/mourede.png'), loadImg(c.logo), loadImg(pSrc)]).then(function (im) {
      var cv = $('inv-cv'); cv.width = W; cv.height = H;
      var cx = cv.getContext('2d');
      cx.fillStyle = '#fff'; cx.fillRect(0, 0, W, H);
      // cadre
      cx.save(); cx.shadowColor = 'rgba(0,0,0,.28)'; cx.shadowBlur = 14; cx.shadowOffsetY = 5;
      cx.fillStyle = '#fff'; rrect(cx, 18, 18, W - 36, H - 36, 60); cx.fill(); cx.restore();
      cx.lineWidth = 10; cx.strokeStyle = '#1b2d55'; rrect(cx, 18, 18, W - 36, H - 36, 60); cx.stroke();
      // en-tête
      fitImg(cx, im[0], 62, 52, 1030, 125, 'l');
      if (hasP) fitImg(cx, im[2], 1540 - 370 * ps, 46, 370 * ps, 160 * ps, 'r');   // partenaire en haut à droite
      else if (im[1]) fitImg(cx, im[1], 1540 - 280 * sc, 40, 280 * sc, 180 * sc, 'r'); // sinon : logo de la structure
      var left = (hasP && im[1]) ? 300 : 62, right = showQr ? 1255 : 1540;
      if (hasP && im[1]) fitImg(cx, im[1], 62, 232, 220 * sc, 220 * sc, 'c');     // logo de la structure à gauche
      var mid = (left + right) / 2, maxW = right - left - 10;
      cx.textAlign = 'center'; cx.textBaseline = 'alphabetic';
      // titre
      var title = ($('inv-title').value || '').trim() || 'EVALUATION DE LA SATISFACTION DU PERSONNEL';
      fitFont(cx, title, maxW, 58, 28, true); cx.fillStyle = c.cTitle || '#c55a11'; cx.fillText(title, mid, 292);
      // structure
      var sl = (($('inv-structline').value || '').trim() || structLabel()).toLocaleUpperCase('fr');
      cx.font = font(46, true); var lines = wrap(cx, sl, maxW).slice(0, 3), y = 358;
      cx.fillStyle = c.cStruct || '#0b6bb8'; lines.forEach(function (t) { cx.fillText(t, mid, y); y += 54; });
      // dates
      var dt = rangeText(st.src.debut, st.src.fin, MOIS_C);
      if (dt) { var line = '📅 ' + dt; fitFont(cx, line, maxW, 44, 24, false); cx.fillStyle = '#111'; cx.fillText(line, mid, y + 14); }
      // QR code du formulaire personnel
      var url = ($('inv-url').value || '').trim() || formUrl(), QS = 240, qx = 1305, qy = 228;
      if (showQr && url && window.QRCode) {
        var h = document.createElement('div'); h.style.cssText = 'position:fixed;left:-9999px;top:0'; document.body.appendChild(h);
        try {
          new QRCode(h, { text: url, width: QS, height: QS, correctLevel: QRCode.CorrectLevel.M });
          var src = h.querySelector('canvas');
          if (src) { cx.fillStyle = '#fff'; cx.fillRect(qx - 6, qy - 6, QS + 12, QS + 12); cx.imageSmoothingEnabled = false; cx.drawImage(src, qx, qy, QS, QS); cx.imageSmoothingEnabled = true; }
        } catch (e) {}
        h.remove();
      }
      // slogan
      if (slogan) { fitFont(cx, slogan, W - 200, 40, 22, false, true); cx.fillStyle = '#111'; cx.fillText(slogan, W / 2, H - 52); }
      return cv;
    });
  }

  /* ---------- Configuration mémorisée par structure (logos, partenaire, slogan) ---------- */
  function slugS(t) { return String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 60); }
  function keyOf(s) { return 'inv_cfg_' + (s.level || 'zone') + '_' + slugS(s.cible || s.name); }
  function loadCfg(s) { try { return JSON.parse(localStorage.getItem(keyOf(s)) || 'null'); } catch (e) { return null; } }
  function persist() {
    if (!(st.src.cible || st.src.name)) return;
    try { localStorage.setItem(keyOf(st.src), JSON.stringify(st.cfg)); }
    catch (e) { $('inv-logo-st').textContent = 'Logo non mémorisé (espace de stockage plein).'; }
  }
  function defaults() { return { logo: '', partner: 'none', partnerLogo: '', sloganOn: st.src.level !== 'hopital', slogan: SLOGAN, qr: true, cTitle: '#c55a11', cStruct: '#0b6bb8', logoScale: 100, pScale: 100 }; }
  function uiFromCfg() {
    var c = st.cfg;
    $('inv-partner').value = c.partner || 'none';
    $('inv-plogo-wrap').style.display = c.partner === 'custom' ? 'block' : 'none';
    $('inv-slogan-on').checked = !!c.sloganOn;
    $('inv-slogan').value = c.slogan || SLOGAN;
    $('inv-qr').checked = c.qr !== false; $('inv-ctitle').value = c.cTitle || '#c55a11'; $('inv-cstruct').value = c.cStruct || '#0b6bb8';
    $('inv-lscale').value = c.logoScale || 100; $('inv-pscale').value = c.pScale || 100;
    $('inv-logo-st').textContent = c.logo ? '✅ Logo importé' : 'Aucun logo importé';
    $('inv-plogo-st').textContent = c.partnerLogo ? '✅ Logo importé' : 'Aucun logo importé';
  }
  function cfgFromUi() {
    var c = st.cfg;
    c.partner = $('inv-partner').value; c.sloganOn = $('inv-slogan-on').checked; c.slogan = $('inv-slogan').value;
    c.qr = $('inv-qr').checked; c.cTitle = $('inv-ctitle').value; c.cStruct = $('inv-cstruct').value; c.logoScale = +$('inv-lscale').value; c.pScale = +$('inv-pscale').value;
    $('inv-plogo-wrap').style.display = c.partner === 'custom' ? 'block' : 'none';
  }
  function readLogo(file, cb) {
    var fr = new FileReader();
    fr.onload = function () {
      var i = new Image();
      i.onload = function () {
        var r = Math.min(1, 500 / Math.max(i.width, i.height)), cv = document.createElement('canvas');
        cv.width = Math.round(i.width * r); cv.height = Math.round(i.height * r);
        cv.getContext('2d').drawImage(i, 0, 0, cv.width, cv.height); cb(cv.toDataURL('image/png'));
      };
      i.onerror = function () { cb(''); }; i.src = fr.result;
    };
    fr.onerror = function () { cb(''); }; fr.readAsDataURL(file);
  }

  /* ---------- Rafraîchissement ---------- */
  function refresh() {
    if (!st.ready) return;
    $('inv-lvl').textContent = st.src.level === 'hopital' ? 'Hôpital' : 'Zone sanitaire';
    if (!$('inv-url').dataset.touched) $('inv-url').value = formUrl();
    if (!st.edited) $('inv-msg').value = buildMsg();
    $('inv-edited').style.display = st.edited ? 'block' : 'none';
    drawBanner();
  }
  function later() { clearTimeout(st.timer); st.timer = setTimeout(refresh, 150); }

  function readSrc() {
    return { level: $('inv-level').value, cible: $('inv-struct').value.trim(), debut: $('inv-debut').value, fin: $('inv-fin').value, name: st.src.name || '' };
  }
  function writeSrc(src) {
    $('inv-level').value = src.level; $('inv-struct').value = src.cible || ''; $('inv-debut').value = src.debut || ''; $('inv-fin').value = src.fin || '';
  }
  function setSrc(src, explicit, label) {
    var levelChanged = src.level !== st.src.level;
    st.src = src; writeSrc(src);
    var stored = loadCfg(src);
    if (stored) { st.cfg = stored; uiFromCfg(); }
    else if (explicit || levelChanged) { st.cfg = defaults(); uiFromCfg(); }
    if (explicit) { st.edited = false; delete $('inv-url').dataset.touched; }
    if (levelChanged) delete $('inv-url').dataset.touched;
    $('inv-src').textContent = label;
    later();
  }
  function srcFromForm() {
    return { level: $('a-level').value, cible: $('a-cible').value.trim(), debut: $('a-debut').value, fin: $('a-fin').value, name: $('a-name').value.trim() };
  }
  function fromForm(explicit) {
    if (!st.ready) return;
    setSrc(srcFromForm(), !!explicit, 'Les champs de droite sont préremplis depuis le formulaire de création ci-dessus ; vous pouvez tout modifier ici, l’aperçu change en direct.');
  }
  function fromActivity(a) {
    setSrc({ level: a.level, cible: a.cible || '', debut: a.debut || '', fin: a.fin || '', name: a.name || '' }, true, 'Invitation de l’activité « ' + a.name + ' ».');
    $('invBox').scrollIntoView({ behavior: 'smooth' });
  }

  /* ---------- Actions ---------- */
  function fname() { return 'Invitation-personnel-' + slugS(structLabel() || 'enquete') + '.png'; }
  function toBlob() { return drawBanner().then(function (cv) { return new Promise(function (res, rej) { cv.toBlob(function (b) { b ? res(b) : rej(new Error('Image impossible')); }, 'image/png'); }); }); }
  function copyText(t, ok) {
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(function () { toast(ok); }, function () { prompt('Copiez ce message :', t); });
    else prompt('Copiez ce message :', t);
  }
  function actDownload() { toBlob().then(function (b) { saveBlob(b, fname()); toast('Image téléchargée'); }).catch(function (e) { err(e.message); }); }
  function actShare() {
    var msg = $('inv-msg').value;
    toBlob().then(function (b) {
      var f = new File([b], fname(), { type: 'image/png' });
      if (navigator.canShare && navigator.canShare({ files: [f] })) {
        try { if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(msg); } catch (e) {}
        navigator.share({ files: [f], text: msg, title: 'Enquête de satisfaction — Personnel' }).then(function () {}, function () {});
        toast('Message copié : collez-le comme légende si besoin');
      } else { saveBlob(b, f.name); copyText(msg, 'Image téléchargée et message copié'); }
    }).catch(function (e) { err(e.message); });
  }
  function actWhatsapp() { window.open('https://wa.me/?text=' + encodeURIComponent($('inv-msg').value), '_blank'); }

  /* ---------- Panneau (éditeur : aperçu à gauche, champs à droite) ---------- */
  function init() {
    var anchor = $('actForm'); if (!anchor || $('invBox')) return;
    var css = document.createElement('style');
    css.textContent = '#invBox .iw{display:grid;gap:18px}@media(min-width:860px){#invBox .iw{grid-template-columns:minmax(0,1.5fr) minmax(290px,1fr);align-items:start}}' +
      '#invBox .side{background:#f7f9fb;border:1px solid #e3e8ee;border-radius:12px;padding:6px 12px 12px}#invBox .side h4{margin:14px 0 2px;font-size:.82rem;color:#0f4c81;text-transform:uppercase;letter-spacing:.03em}' +
      '#invBox .side input,#invBox .side select{margin-top:4px;padding:9px 10px;font-size:.92rem}#invBox .side input[type=color]{padding:2px;height:38px;width:100%}#invBox .side input[type=range]{padding:0;border:0}' +
      '#invBox .two{display:grid;grid-template-columns:1fr 1fr;gap:8px}#invBox .two>div{min-width:0}#invBox .two input{min-width:0;padding:9px 6px}';
    document.head.appendChild(css);
    var box = document.createElement('div'); box.className = 'box'; box.id = 'invBox';
    box.innerHTML =
      '<h3>📣 Éditeur : message et invitation du personnel <span class="badge" id="inv-lvl"></span></h3>' +
      '<p class="hint" id="inv-src" style="margin-top:6px"></p>' +
      '<div class="iw">' +
      ' <div>' +
      '  <canvas id="inv-cv" width="1600" height="520" style="width:100%;height:auto;border:1px solid #dfe4ea;border-radius:10px;background:#fff"></canvas>' +
      '  <label style="margin-top:12px">Message à envoyer</label><textarea id="inv-msg" rows="12" style="width:100%;padding:11px 12px;border:1px solid #dfe4ea;border-radius:10px;font-size:.95rem;font-family:inherit;margin-top:6px"></textarea>' +
      '  <p class="hint" id="inv-edited" style="display:none;margin-top:6px">Message modifié à la main : il ne suit plus les champs. « Régénérer » le reconstruit.</p>' +
      '  <div class="row"><button type="button" class="pri" id="inv-share">📤 Envoyer image + message</button><button type="button" class="alt" id="inv-dl">⬇️ Image</button><button type="button" class="alt" id="inv-copy">📋 Message</button><button type="button" class="alt" id="inv-wa">💬 WhatsApp</button><button type="button" class="alt" id="inv-regen">↺ Régénérer</button></div>' +
      ' </div>' +
      ' <div class="side">' +
      '  <h4>Structure</h4>' +
      '  <label>Niveau</label><select id="inv-level"><option value="zone">Zone sanitaire (bureau de zone)</option><option value="hopital">Hôpital</option></select>' +
      '  <label style="margin-top:6px">Nom de la zone sanitaire / de l’hôpital</label><input id="inv-struct" list="dl-zone" placeholder="Ex. Covè-Zagnanado-Ouinhi">' +
      '  <label style="margin-top:6px">Ligne affichée sur l’image (vide = automatique)</label><input id="inv-structline" placeholder="Ex. HÔPITAL DE ZONE DE DJIDJA">' +
      '  <label style="margin-top:6px">Titre de l’image</label><input id="inv-title" value="EVALUATION DE LA SATISFACTION DU PERSONNEL">' +
      '  <h4>Période</h4><div class="two"><div><label>Début</label><input id="inv-debut" type="date"></div><div><label>Fin</label><input id="inv-fin" type="date"></div></div>' +
      '  <h4>Lien et QR code</h4><label>Lien du formulaire du personnel</label><input id="inv-url" placeholder="https://…">' +
      '  <div class="chk" style="margin-top:6px"><label><input type="checkbox" id="inv-qr">Afficher le QR code sur l’image</label></div>' +
      '  <h4>Logos</h4>' +
      '  <label>Logo de la structure</label><input type="file" id="inv-logo" accept="image/*"><div class="row" style="align-items:center;margin-top:4px"><span class="meta" id="inv-logo-st"></span><button type="button" class="alt" id="inv-logo-rm">Retirer</button></div>' +
      '  <label style="margin-top:6px">Taille du logo de la structure</label><input type="range" id="inv-lscale" min="50" max="110" step="5" value="100">' +
      '  <label style="margin-top:6px">Partenaire / appui financier</label><select id="inv-partner"><option value="none">Aucun partenaire</option><option value="msv">Médecins Sans Vacances (MSV)</option><option value="custom">Autre partenaire (importer son logo)</option></select>' +
      '  <div id="inv-plogo-wrap" style="display:none"><input type="file" id="inv-plogo" accept="image/*"><div class="row" style="align-items:center;margin-top:4px"><span class="meta" id="inv-plogo-st"></span><button type="button" class="alt" id="inv-plogo-rm">Retirer</button></div></div>' +
      '  <label style="margin-top:6px">Taille du logo du partenaire</label><input type="range" id="inv-pscale" min="50" max="110" step="5" value="100">' +
      '  <h4>Texte et couleurs</h4>' +
      '  <div class="chk"><label><input type="checkbox" id="inv-slogan-on">Phrase d’accroche en bas</label></div><input id="inv-slogan" placeholder="Phrase d’accroche">' +
      '  <div class="two" style="margin-top:6px"><div><label>Couleur du titre</label><input type="color" id="inv-ctitle" value="#c55a11"></div><div><label>Couleur de la structure</label><input type="color" id="inv-cstruct" value="#0b6bb8"></div></div>' +
      ' </div>' +
      '</div>';
    anchor.parentNode.insertBefore(box, anchor.nextSibling);
    st.ready = true; st.cfg = defaults(); uiFromCfg();

    $('inv-logo').addEventListener('change', function () {
      var f = this.files && this.files[0]; if (!f) return;
      readLogo(f, function (d) { if (!d) { err('Image illisible.'); return; } st.cfg.logo = d; $('inv-logo-st').textContent = '✅ Logo importé'; persist(); later(); });
    });
    $('inv-logo-rm').addEventListener('click', function () { st.cfg.logo = ''; $('inv-logo').value = ''; $('inv-logo-st').textContent = 'Aucun logo importé'; persist(); later(); });
    $('inv-plogo').addEventListener('change', function () {
      var f = this.files && this.files[0]; if (!f) return;
      readLogo(f, function (d) { if (!d) { err('Image illisible.'); return; } st.cfg.partnerLogo = d; $('inv-plogo-st').textContent = '✅ Logo importé'; persist(); later(); });
    });
    $('inv-plogo-rm').addEventListener('click', function () { st.cfg.partnerLogo = ''; $('inv-plogo').value = ''; $('inv-plogo-st').textContent = 'Aucun logo importé'; persist(); later(); });
    ['inv-partner', 'inv-slogan-on', 'inv-qr', 'inv-ctitle', 'inv-cstruct', 'inv-lscale', 'inv-pscale', 'inv-slogan'].forEach(function (id) {
      var h = function () { cfgFromUi(); persist(); later(); };
      $(id).addEventListener('input', h); $(id).addEventListener('change', h);
    });
    ['inv-struct', 'inv-debut', 'inv-fin', 'inv-structline', 'inv-title'].forEach(function (id) { $(id).addEventListener('input', function () { st.src = readSrc(); later(); }); });
    $('inv-struct').addEventListener('change', function () { var sc = loadCfg(readSrc()); st.src = readSrc(); if (sc) { st.cfg = sc; uiFromCfg(); } later(); });
    $('inv-level').addEventListener('change', function () {
      var lv = $('inv-level').value, sc = loadCfg(readSrc());
      st.src = readSrc(); delete $('inv-url').dataset.touched;
      if (sc) st.cfg = sc; else st.cfg.sloganOn = lv !== 'hopital';
      uiFromCfg(); later();
    });
    $('inv-url').addEventListener('input', function () { this.dataset.touched = '1'; later(); });
    $('inv-msg').addEventListener('input', function () { st.edited = true; $('inv-edited').style.display = 'block'; });
    $('inv-regen').addEventListener('click', function () { st.edited = false; delete $('inv-url').dataset.touched; refresh(); toast('Message régénéré'); });
    $('inv-copy').addEventListener('click', function () { copyText($('inv-msg').value, 'Message copié'); });
    $('inv-dl').addEventListener('click', actDownload);
    $('inv-share').addEventListener('click', actShare);
    $('inv-wa').addEventListener('click', actWhatsapp);
    ['a-name', 'a-level', 'a-cible', 'a-debut', 'a-fin'].forEach(function (id) {
      var el = $(id); if (!el) return;
      el.addEventListener('input', function () { fromForm(false); }); el.addEventListener('change', function () { fromForm(false); });
    });
    fromForm(true);
  }

  window.Inv = { init: init, fromForm: fromForm, fromActivity: fromActivity, persist: persist, buildMsg: buildMsg };
})();
