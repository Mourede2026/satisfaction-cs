/* Attestation de bonne fin d'exécution (enquêteur de terrain) — génération d'un PDF d'une page, non modifiable.
   Reproduit le canevas Word fourni (zone sanitaire / hôpital). Aucune donnée n'est envoyée : tout se fait dans le navigateur.
   Dépend de jsPDF (window.jspdf) pour l'assemblage du PDF. Les polices et images sont dans le dossier att/. */
(function (g) {
  var PW = 841.89, PH = 595.28;            // A4 paysage en points
  var DPI = 300, S = DPI / 72;              // résolution de rendu (300 dpi)
  var MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  var SIGNATAIRE = { nom: 'KABA M. M. Sandra M.', fonction: 'La Directrice' };
  var QUALITE_DEFAUT = 'Enquêteur (trice) de Terrain';
  var assets = null;

  function frDate(iso) {
    var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? m[3] + ' ' + MOIS[+m[2] - 1] + ' ' + m[1] : '';
  }
  function loadImg(src) {
    return new Promise(function (res, rej) { var i = new Image(); i.onload = function () { res(i); }; i.onerror = function () { rej(new Error('Image introuvable : ' + src)); }; i.src = src; });
  }
  function loadFont(family, url, weight, style) {
    var f = new FontFace(family, 'url(' + url + ')', { weight: weight, style: style || 'normal' });
    return f.load().then(function (ff) { document.fonts.add(ff); return ff; });
  }
  function loadAssets() {
    if (assets) return assets;
    assets = Promise.all([
      loadImg('att/cadre.jpg'), loadImg('att/logo.png'), loadImg('att/tampon.png'),
      loadFont('AttNarrow', 'att/narrow.woff', '400'), loadFont('AttNarrow', 'att/narrow-bold.woff', '700'),
      loadFont('AttGeorgia', 'att/gelasio-bold.woff', '700'), loadFont('AttGeorgia', 'att/gelasio-bolditalic.woff', '700', 'italic')
    ]).then(function (r) { return { cadre: r[0], logo: r[1], tampon: r[2] }; }).catch(function (e) { assets = null; throw e; });
    return assets;
  }

  var NARROW = function (pt, bold) { return (bold ? '700 ' : '400 ') + pt + 'px AttNarrow, "Arial Narrow", "Liberation Sans Narrow", Arial, sans-serif'; };
  var GEORGE = function (pt, italic) { return '700 ' + (italic ? 'italic ' : '') + pt + 'px AttGeorgia, Georgia, serif'; };

  function runsWidth(cx, runs, pt) {
    var w = 0; runs.forEach(function (r) { cx.font = NARROW(pt, r.b); w += cx.measureText(r.t).width; }); return w;
  }
  function drawRuns(cx, runs, x, y, pt, extraSpace) {
    runs.forEach(function (r) {
      cx.font = NARROW(pt, r.b);
      if (!extraSpace) { cx.fillText(r.t, x, y); x += cx.measureText(r.t).width; return; }
      // justification : l'espace excédentaire est réparti sur les espaces
      r.t.split(/( )/).forEach(function (part) {
        if (part === ' ') { x += cx.measureText(' ').width + extraSpace; }
        else if (part) { cx.fillText(part, x, y); x += cx.measureText(part).width; }
      });
    });
  }
  /** Découpe des segments (texte + gras) en lignes, aux espaces, sans dépasser maxW. */
  function wrap(cx, runs, pt, maxW) {
    var toks = [];
    runs.forEach(function (r) { r.t.split(/( )/).forEach(function (p) { if (p !== '') toks.push({ t: p, b: r.b }); }); });
    var lines = [], cur = [], w = 0;
    toks.forEach(function (tk) {
      cx.font = NARROW(pt, tk.b); var tw = cx.measureText(tk.t).width;
      if (tk.t === ' ') { if (cur.length) { cur.push(tk); w += tw; } return; }
      if (w + tw > maxW && cur.length) {
        while (cur.length && cur[cur.length - 1].t === ' ') cur.pop();
        lines.push(cur); cur = []; w = 0;
      }
      cur.push(tk); w += tw;
    });
    while (cur.length && cur[cur.length - 1].t === ' ') cur.pop();
    if (cur.length) lines.push(cur);
    return lines.map(function (l) { // fusionne les jetons voisins de même style
      var out = []; l.forEach(function (t) { var p = out[out.length - 1]; if (p && p.b === t.b) p.t += t.t; else out.push({ t: t.t, b: t.b }); }); return out;
    });
  }

  /** Données : {nom, prenom, level:'zone'|'hopital', libelle, debut, fin, annee, lieu, dateSignature, qualite} */
  function render(d) {
    return loadAssets().then(function (A) {
      var cv = document.createElement('canvas'); cv.width = Math.round(PW * S); cv.height = Math.round(PH * S);
      var cx = cv.getContext('2d');
      cx.fillStyle = '#fff'; cx.fillRect(0, 0, cv.width, cv.height);
      cx.setTransform(S, 0, 0, S, 0, 0);
      cx.imageSmoothingQuality = 'high'; cx.textBaseline = 'alphabetic';
      // fond (cadre bleu) et logo
      cx.drawImage(A.cadre, 0, -15, 842.2, 610.5);
      cx.drawImage(A.logo, 244.5, 21, 353.1, 47.5);
      // titre
      cx.textAlign = 'center'; cx.fillStyle = '#1A1A2E'; cx.font = GEORGE(28); cx.fillText('ATTESTATION', PW / 2, 93.7);
      cx.fillStyle = '#1E4696'; cx.font = NARROW(12, true); cx.fillText("DE BONNE FIN D'EXÉCUTION", PW / 2, 118.8);
      // « La soussignée … »
      cx.textAlign = 'left'; cx.fillStyle = '#000';
      drawRuns(cx, [{ t: 'La soussignée Madame ', b: 0 }, { t: SIGNATAIRE.nom, b: 1 }, { t: ', ', b: 0 }, { t: "Directrice de l'Etablissement MOUREDE & ASSOCIES", b: 1 }, { t: ', atteste que le/la nommé(e) :', b: 0 }], 90.5, 144.3, 13);
      // nom complet (adapté si très long)
      var full = String(d.nom || '').toUpperCase() + ' ' + String(d.prenom || ''), sz = 22;
      cx.font = GEORGE(sz, true); while (cx.measureText(full).width > 640 && sz > 12) { sz -= 1; cx.font = GEORGE(sz, true); }
      cx.textAlign = 'center'; cx.fillStyle = '#142A5A'; cx.fillText(full, PW / 2, 179.8);
      cx.fillStyle = '#000'; cx.font = NARROW(13); cx.fillText('a effectivement participé en tant que', PW / 2, 206.8);
      var q = String(d.qualite || QUALITE_DEFAUT), qs = 20;
      cx.font = GEORGE(qs, true); while (cx.measureText(q).width > 640 && qs > 12) { qs -= 1; cx.font = GEORGE(qs, true); }
      cx.fillStyle = '#142A5A'; cx.fillText(q, PW / 2, 240.6);
      // corps
      cx.textAlign = 'left'; cx.fillStyle = '#000';
      var X0 = 55.1, W = 732.3, PITCH = 20.8, GAP = 6.1, y = 273.1;
      var lieuTxt = d.level === 'hopital' ? "de l'Hôpital de zone de " : 'de la Zone Sanitaire ';
      var A1 = [{ t: 'Dans le cadre de l’enquête de satisfaction des usagers menée au sein ' + lieuTxt, b: 0 }, { t: String(d.libelle || ''), b: 1 },
        { t: " au titre de l'année ", b: 0 }, { t: String(d.annee || ''), b: 1 }, { t: ' du ', b: 0 }, { t: frDate(d.debut), b: 1 }, { t: ' au ', b: 0 }, { t: frDate(d.fin), b: 1 }, { t: '.', b: 0 }];
      var B1 = [{ t: 'A ce titre, il/elle a réalisé les interviews auprès des usagers selon le questionnaire; collecté fidèlement les réponses et noté les observations; respecté les règles d’éthique et de confidentialité; remonté les difficultés rencontrées au superviseur; assuré la qualité de la saisie initiale via tablette ou smartphone. Il/elle a accompli ces tâches avec sérieux, rigueur et professionnalisme, conformément aux outils méthodologiques et numériques mis à disposition.', b: 0 }];
      [A1, B1].forEach(function (para) {
        var lines = wrap(cx, para, 13, W);
        lines.forEach(function (ln, i) {
          var last = i === lines.length - 1, extra = 0;
          if (!last) { var sp = 0; ln.forEach(function (r) { sp += (r.t.match(/ /g) || []).length; }); extra = sp ? (W - runsWidth(cx, ln, 13)) / sp : 0; }
          drawRuns(cx, ln, X0, y, 13, extra);
          if (!last) y += PITCH;
        });
        y += PITCH + GAP;
      });
      cx.font = NARROW(13); cx.fillText('La présente attestation est délivrée pour servir et valoir ce que de droit.', X0, y);
      y += 14.3 + 0;                                  // « Fait à … » (aligné à droite)
      cx.textAlign = 'right'; cx.fillText('Fait à ' + String(d.lieu || '') + ', le ' + frDate(d.dateSignature), 752, y);
      // signature : titre, tampon + signature, nom
      cx.textAlign = 'center';
      cx.font = NARROW(13, true); cx.fillText(SIGNATAIRE.fonction, 448, 434.0);
      cx.drawImage(A.tampon, 314.4, 421.6, 148.3, 97.75);
      cx.fillText(SIGNATAIRE.nom, 450, 510.4);
      // pied de page
      cx.textAlign = 'left';
      drawRuns(cx, [{ t: 'RCCM :', b: 1 }, { t: ' RB/ODH/10-A-028   ', b: 0 }, { t: 'IFU :', b: 1 }, { t: ' 3 2010 0030 3210   ', b: 0 }, { t: 'Tél :', b: 1 }, { t: ' 00229 01 97 51 81 22 – 01 97 34 73 26', b: 0 }], 35.2, 550.0, 10);
      return cv;
    });
  }

  function jsPDFClass() { return g.jspdf && g.jspdf.jsPDF; }
  function randomPwd() { var a = new Uint8Array(16); (g.crypto || g.msCrypto).getRandomValues(a); return Array.prototype.map.call(a, function (b) { return ('0' + b.toString(16)).slice(-2); }).join(''); }

  /** Blob PDF d'une page. L'image est aplatie (le texte n'est plus modifiable) et le PDF est protégé contre la modification. */
  function toPdfBlob(d) {
    var J = jsPDFClass(); if (!J) return Promise.reject(new Error('Bibliothèque PDF non chargée (connexion Internet requise).'));
    return render(d).then(function (cv) {
      var doc = new J({ orientation: 'landscape', unit: 'pt', format: 'a4', compress: true, encryption: { userPermissions: ['print'], ownerPassword: randomPwd() } });
      doc.setProperties({ title: 'Attestation de bonne fin d\'exécution — ' + d.nom + ' ' + d.prenom, subject: 'Enquête de satisfaction', author: 'MOUREDE & ASSOCIES', creator: 'MOUREDE & ASSOCIES' });
      doc.addImage(cv.toDataURL('image/jpeg', 0.93), 'JPEG', 0, 0, PW, PH, undefined, 'FAST');
      cv.width = cv.height = 1;
      return doc.output('blob');
    });
  }
  function slug(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]+/g, '_').replace(/^_|_$/g, ''); }
  function fileName(d) { return 'Attestation_' + (d.level === 'hopital' ? 'Hopital' : 'Zone_Sanitaire') + '_' + slug(String(d.nom || '').toUpperCase()) + '_' + slug(d.prenom) + '.pdf'; }

  g.Attestation = { render: render, toPdfBlob: toPdfBlob, fileName: fileName, frDate: frDate, loadAssets: loadAssets, QUALITE_DEFAUT: QUALITE_DEFAUT, MOIS: MOIS };
})(window);
