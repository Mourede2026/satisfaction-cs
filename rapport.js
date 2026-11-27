/* rapport.js — contenu automatique des rapports et présentations (v28)
 * Construit, à partir des données renvoyées par le serveur (action reportData) et des informations saisies par l'administrateur,
 * un « modèle » de rapport (titres, textes, tableaux, graphiques) puis le transforme en Word (.docx) et en PowerPoint (.pptx).
 * Zone sanitaire : format ZoBoZa 2026. Hôpital : canevas 2026 des hôpitaux de zone. */
(function (root) {
  'use strict';
  var BRAND = '0F4C81', GOOD = '2E9E5B', BAD = 'C0392B', GREY = '6B7280', LIGHT = 'E8F0F8', AMBER = 'D98E04';
  var CIBLE_LABEL = { usagers: 'Usagers des services de santé', personnels: 'Personnel soignant', cogecs: 'Membres des COGECS' };
  var CIBLE_COURT = { usagers: 'Usagers', personnels: 'Personnel soignant', cogecs: 'COGECS' };

  /* ---------- utilitaires ---------- */
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function f1(v) { return isNum(v) ? (Math.round(v * 10) / 10).toFixed(1).replace('.', ',') : '—'; }
  function pc(v) { return isNum(v) ? f1(v) + ' %' : '—'; }
  function fN(v) { return String(v).replace(/\B(?=(\d{3})+(?!\d))/g, '\u00A0'); }
  function parseNum(s) { var d = String(s === null || s === undefined ? '' : s).replace(/[\s\u00A0\u202F]/g, '').replace(/[^0-9,.]/g, '').replace(',', '.'); var v = parseFloat(d); return isFinite(v) ? v : 0; }
  /** Densité (hab/km²) calculée à partir de la population et de la superficie saisies. */
  function density(P) { P = P || {}; var pop = parseNum(P.population), sup = parseNum(P.superficie); return { pop: pop, sup: sup, dens: pop > 0 && sup > 0 ? Math.round(pop / sup) : 0 }; }
  /** Phrase « chiffres clés » : superficie, population, densité calculée, population cible. */
  function chiffresCles(P) {
    P = P || {}; var d = density(P), t = [];
    if (P.superficie) t.push('une superficie de ' + P.superficie + (/km/i.test(P.superficie) ? '' : ' km²'));
    if (P.population) t.push('une population estimée à ' + P.population + ' habitants' + (d.dens ? ', soit une densité de ' + fN(d.dens) + ' habitants par km²' : ''));
    if (P.populationCible) t.push('une population cible de ' + P.populationCible + ' personnes');
    return t.length ? 'Elle présente ' + t.join(', ') + '.' : '';
  }
  function extraFigs(M, P, sec) { ((P && P.figures) || []).forEach(function (f) { if (f && f.src && f.section === sec) M.image(f.legende || 'Figure', f.src); }); }
  function b(v) { return '**' + v + '**'; }
  function mean(a) { var s = 0, n = 0; a.forEach(function (v) { if (isNum(v)) { s += v; n++; } }); return n ? s / n : null; }
  function roman(n) { var m = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']], r = ''; m.forEach(function (p) { while (n >= p[0]) { r += p[1]; n -= p[0]; } }); return r; }
  function qual(v) { if (!isNum(v)) return 'non évalué'; return v >= 85 ? 'très satisfaisant' : v >= 70 ? 'satisfaisant' : v >= 50 ? 'moyennement satisfaisant' : 'insatisfaisant'; }
  function delta(a, bb) { if (!isNum(a) || !isNum(bb)) return '—'; var d = Math.round((a - bb) * 10) / 10; return (d >= 0 ? '+' : '-') + f1(Math.abs(d)); }
  function cleanDom(n) { return String(n).replace(/^\d+\.\s*/, ''); }
  function cap(s) { s = String(s || ''); return s.charAt(0).toUpperCase() + s.slice(1); }
  function frDate(s) { return s ? String(s).split('-').reverse().join('/') : ''; }
  function plural(n, s, p) { return n + ' ' + (n > 1 ? (p || s + 's') : s); }
  function listFr(a) { a = a.filter(Boolean); if (a.length <= 1) return a.join(''); return a.slice(0, -1).join(', ') + ' et ' + a[a.length - 1]; }
  function sortDesc(arr, key) { return arr.slice().sort(function (x, y) { return (y[key] === null ? -1 : y[key]) - (x[key] === null ? -1 : x[key]); }); }
  function pctOf(n, t) { return t > 0 ? Math.round(n / t * 1000) / 10 : null; }
  function colorFor(v, m) { return isNum(m) ? (v >= m ? GOOD : BAD) : BRAND; }

  /* ---------- constructeur du modèle ---------- */
  function Model() { this.blocks = []; this.nt = 0; this.nf = 0; }
  Model.prototype.h = function (l, t, o) { this.blocks.push({ t: 'h' + l, text: t, noNum: o && o.noNum }); };
  Model.prototype.p = function (t) { if (t) this.blocks.push({ t: 'p', text: t }); };
  Model.prototype.bullets = function (a) { a = (a || []).filter(Boolean); if (a.length) this.blocks.push({ t: 'ul', items: a }); };
  Model.prototype.table = function (cap, head, rows, o) { o = o || {}; this.nt++; this.blocks.push({ t: 'table', no: this.nt, caption: cap, head: head, rows: rows, lastBold: !!o.lastBold, boldRows: o.boldRows || [], note: o.note || '', first: o.first || 'left' }); };
  Model.prototype.chart = function (cap, spec) { this.nf++; this.blocks.push({ t: 'chart', no: this.nf, caption: cap, spec: spec }); };
  Model.prototype.image = function (cap, src) { if (!src) return; this.nf++; this.blocks.push({ t: 'image', no: this.nf, caption: cap, src: src }); };
  Model.prototype.pb = function () { this.blocks.push({ t: 'pb' }); };
  Model.prototype.tables = function () { return this.blocks.filter(function (x) { return x.t === 'table'; }); };
  Model.prototype.figures = function () { return this.blocks.filter(function (x) { return x.t === 'chart' || x.t === 'image'; }); };

  /* ---------- phrases d'analyse ---------- */
  /** rows = [{name, score}] (formations sanitaires) ; renvoie la phrase « moyenne / mieux notée / moins bien notée ». */
  function domainSentence(scope, avg, rows, unit) {
    var ok = rows.filter(function (r) { return isNum(r.score); }), t = 'Dans ' + scope + ', ce domaine obtient une moyenne de ' + b(pc(avg)) + ' (' + b(qual(avg)) + ').';
    if (ok.length >= 2) {
      var s = sortDesc(ok, 'score'), hi = s[0], lo = s[s.length - 1];
      t += ' ' + cap(unit.best) + ' est ' + hi.name + ' (' + b(pc(hi.score)) + '), contre ' + lo.name + ' (' + b(pc(lo.score)) + ') pour ' + unit.worst + ', soit un écart de ' + f1(hi.score - lo.score) + ' points entre ' + unit.plural + '.';
    } else if (ok.length === 1) t += ' Un seul élément est évalué : ' + ok[0].name + ' (' + b(pc(ok[0].score)) + ').';
    return t;
  }
  var UNIT_FS = { best: 'la formation sanitaire la mieux notée', worst: 'la moins bien notée', plural: 'formations sanitaires' };
  var UNIT_COM = { best: 'la commune la mieux notée', worst: 'la moins bien notée', plural: 'communes' };
  /** Appréciation globale : domaines au-dessus / en dessous de la moyenne. */
  function globalAppraisal(scope, level, domains, scores, label) {
    var avg = mean(scores), up = [], down = [];
    domains.forEach(function (d, i) { if (!isNum(scores[i])) return; (scores[i] >= avg ? up : down).push({ n: cleanDom(d), s: scores[i] }); });
    var t = 'Sur les ' + domains.length + ' domaines évalués ' + scope + ', ' + up.length + ' ont obtenu un score supérieur ou égal à la moyenne ' + level + ', fixée à ' + b(pc(avg)) + (up.length ? ' : ' + up.map(function (x) { return '« ' + x.n + ' »'; }).join(', ') : '') + '.';
    if (down.length) t += '\n' + 'Les domaines en dessous de cette moyenne sont : ' + down.map(function (x) { return x.n + ' (' + b(pc(x.s)) + ')'; }).join(', ') + '.';
    return t;
  }

  /* ---------- recommandations types ---------- */
  var RECO = {
    'Accueil et orientation': "Améliorer l'accueil et l'orientation des usagers (signalétique, agent d'accueil, courtoisie) dans les formations sanitaires les moins bien notées.",
    'Information et communication': "Renforcer l'information et la communication avec les usagers (explications sur l'état de santé, le traitement et les coûts) dans les formations où ces items sont faibles.",
    'Qualité de l’examen et de l’écoute': "Consolider la qualité de l'examen clinique et de l'écoute des patients, en particulier dans les formations sanitaires les moins bien notées.",
    'Soins et continuité': "Renforcer la continuité des soins et le suivi des patients après la consultation.",
    'Respect, intimité et confidentialité': "Garantir le respect de l'intimité et de la confidentialité des patients dans toutes les salles de soins.",
    'Accessibilité financière et médicaments': "Renforcer la disponibilité et l'accessibilité financière des médicaments dans les formations sanitaires les moins bien notées sur cette dimension.",
    'Environnement et hygiène': "Améliorer l'hygiène et l'entretien des locaux (propreté, toilettes, points d'eau, salles d'attente).",
    'Équité, respect et compétence du personnel': "Renforcer l'équité, le respect et la compétence du personnel soignant (formation continue, supervision, prise en charge sans discrimination).",
    'Transparence et pratiques abusives': "Lutter contre les pratiques abusives et renforcer la transparence (affichage des tarifs, délivrance systématique de justificatifs, tolérance zéro pour les sollicitations indues).",
    'Rapidité et réactivité de la prise en charge': "Réduire les temps d'attente et améliorer la réactivité de la prise en charge, y compris aux heures creuses.",
    'Vaccination': "Améliorer l'organisation de la vaccination (temps d'attente, information des mères, disponibilité des vaccins).",
    'Ressources matérielles et infrastructures': "Doter les formations sanitaires des équipements, consommables et infrastructures nécessaires à des soins de qualité.",
    'Organisation du travail et charge professionnelle': "Rééquilibrer la répartition des tâches et alléger la charge de travail du personnel soignant.",
    'Valorisation, reconnaissance et perspectives professionnelles': "Mettre en place un plan de reconnaissance et de valorisation du personnel soignant (perspectives de carrière, motivation).",
    'Relations interpersonnelles et climat social': "Entretenir un climat social apaisé dans les équipes (gestion constructive des conflits, soutien entre collègues).",
    'Protocoles, directives et communication institutionnelle': "Rendre les protocoles et les directives du Bureau de Zone Sanitaire accessibles et mieux communiqués aux agents.",
    'Compétences techniques et fiabilité de l’équipe': "Renforcer les compétences techniques des équipes par des formations ciblées.",
    'Supervision, encadrement et formation': "Renforcer la supervision et l'encadrement technique du personnel, en particulier dans les formations sanitaires les moins bien notées.",
    'Information et charte des droits des patients': "Améliorer l'information des usagers sur leurs droits et devoirs (Charte, personne de confiance, refus de traitement).",
    'Pratiques éthiques et relation patient': "Veiller aux pratiques éthiques et à la qualité de la relation soignant-patient (consentement, confidentialité, dignité).",
    'Gouvernance et collaboration avec le COGECS': "Dynamiser la collaboration entre le personnel et le COGECS (réunions régulières, gestion des conflits, sensibilisation communautaire).",
    'Gouvernance et gestion du centre de santé': "Dynamiser la gouvernance et la transparence des COGECS, notamment sur la gestion des ressources et les rapports transmis à la coordination de zone.",
    'Collaboration interne avec le personnel soignant': "Renforcer la collaboration entre les COGECS et le personnel soignant.",
    'Relation avec la coordination de zone sanitaire': "Renforcer le dialogue entre les COGECS et la coordination de zone sanitaire.",
    'Capacités et moyens des membres du COGECS': "Former les membres des COGECS à leurs rôles et responsabilités et leur fournir des outils (manuels, guides, fiches)."
  };
  function recoFor(dom) { var k = cleanDom(dom); return RECO[k] || null; }

  root.RapportBase = { BRAND: BRAND, GOOD: GOOD, BAD: BAD, GREY: GREY, LIGHT: LIGHT, AMBER: AMBER, CIBLE_LABEL: CIBLE_LABEL, CIBLE_COURT: CIBLE_COURT, isNum: isNum, fN: fN, parseNum: parseNum, density: density, chiffresCles: chiffresCles, extraFigs: extraFigs, f1: f1, pc: pc, b: b, mean: mean, roman: roman, qual: qual, delta: delta, cleanDom: cleanDom, cap: cap, frDate: frDate, plural: plural, listFr: listFr, sortDesc: sortDesc, pctOf: pctOf, colorFor: colorFor, Model: Model, domainSentence: domainSentence, globalAppraisal: globalAppraisal, UNIT_FS: UNIT_FS, UNIT_COM: UNIT_COM, RECO: RECO, recoFor: recoFor };
})(typeof window !== 'undefined' ? window : globalThis);
/* ===== Zone sanitaire : modèle de rapport (format ZoBoZa 2026) ===== */
(function (root) {
  'use strict';
  var B = root.RapportBase, f1 = B.f1, pc = B.pc, b = B.b, mean = B.mean, isNum = B.isNum, delta = B.delta;
  var CL = B.CIBLE_LABEL, CC = B.CIBLE_COURT, KEYS = ['usagers', 'personnels', 'cogecs'];
  var DE = { usagers: 'des usagers', personnels: 'du personnel soignant', cogecs: 'des membres des COGECS' };
  var DEL = { usagers: 'des usagers des services de santé', personnels: 'du personnel soignant', cogecs: 'des membres des COGECS' };
  var PAR = { usagers: 'par les usagers des services de santé', personnels: 'par le personnel soignant', cogecs: 'par les membres des COGECS' };

  function ctx(data, P) {
    var c = { data: data, P: P || {}, main: data.main, communes: data.communes || [], domains: data.domains };
    c.zone = (P && P.nomZone) || String(data.zone || '').replace(/^(zone sanitaire|zs)\s+/i, '') || 'la zone';
    c.ZONE = c.zone.toUpperCase();
    c.edition = (P && P.edition) || (data.end ? String(data.end).slice(0, 4) : String(new Date().getFullYear()));
    c.periode = (P && P.periodeTexte) || (data.periodLabels && data.periodLabels.length ? data.periodLabels.join(' + ') : c.edition);
    c.fs = Object.keys(c.main.byStructure || {}).sort(function (x, y) { return x.localeCompare(y, 'fr'); });
    c.fsOf = function (com) { return c.fs.filter(function (s) { return c.main.byStructure[s].commune === com; }); };
    c.n = c.main.n || {};
    c.total = (c.n.usagers || 0) + (c.n.personnels || 0) + (c.n.cogecs || 0);
    c.quota = Number((P && P.quotaUsagers) || 60) || 60;
    c.com = function (cn) { return /^[AEIOUYÉÈ]/i.test(cn) ? "d'" + cn : 'de ' + cn; };
    return c;
  }
  function fsRows(c, key, com, dom) {
    return c.fsOf(com).filter(function (s) { return c.main.byStructure[s][key]; }).map(function (s) { var g = c.main.byStructure[s][key]; return { name: s, n: g.n, score: g.d[dom] }; });
  }
  function crossTable(M, cap, communes, cats, getCounts, labelFirst) {
    var tot = {}, perCom = communes.map(function (c) { var o = getCounts(c) || {}; Object.keys(o).forEach(function (k) { tot[k] = (tot[k] || 0) + o[k]; }); return o; });
    var sums = perCom.map(function (o) { return cats.reduce(function (s, k) { return s + (o[k] || 0); }, 0); }), gt = sums.reduce(function (a, x) { return a + x; }, 0);
    var head = [labelFirst || '']; communes.forEach(function (c) { head.push(c, '%'); }); head.push('Total', '%');
    var rows = cats.map(function (k) {
      var r = [k]; communes.forEach(function (c, i) { var n = perCom[i][k] || 0; r.push(String(n), pc(B.pctOf(n, sums[i]))); });
      var t = cats.length ? (tot[k] || 0) : 0; r.push(String(t), pc(B.pctOf(t, gt))); return r;
    });
    var last = ['Total']; communes.forEach(function (c, i) { last.push(String(sums[i]), sums[i] ? '100,0 %' : '—'); }); last.push(String(gt), gt ? '100,0 %' : '—'); rows.push(last);
    M.table(cap, head, rows, { lastBold: true });
    return { tot: tot, gt: gt, perCom: perCom, sums: sums };
  }
  function topCat(tot, gt) { var k = Object.keys(tot).sort(function (a, z) { return tot[z] - tot[a]; })[0]; return k ? { k: k, p: B.pctOf(tot[k], gt) } : null; }

  /* ---------- texte fixe adapté du protocole ZoBoZa ---------- */
  function intro(M, c) {
    var P = c.P, comT = B.listFr(c.communes), fsT = P.nbFsPub ? ' dotées de ' + P.nbFsPub + ' formations sanitaires publiques' + (P.nbFsPriv ? ' et ' + P.nbFsPriv + ' privées' : '') : '';
    M.h(1, 'Introduction', { noNum: true });
    M.p("La satisfaction en milieu de soins est une composante essentielle de la qualité des services de santé. Elle reflète la réaction des usagers face aux évènements marquant leur contact avec les formations sanitaires, influencée par leurs attentes, leurs expériences et le contexte dans lequel les soins sont délivrés.");
    M.p("Depuis 2021, les zones sanitaires du Bénin ont l'obligation réglementaire de réaliser au moins une enquête annuelle de satisfaction des usagers. Cet indicateur, inscrit dans le cadre de performance, vise à objectiver la perception des soins par les bénéficiaires et à orienter les actions d'amélioration.");
    M.p("La zone sanitaire " + c.zone + (P.departement ? ", située dans le département " + (/^[AEIOUYÉ]/i.test(P.departement) ? "de l'" : "du ") + P.departement : '') + ", regroupe les communes de " + (comT || '—') + fsT + '. Ces structures de premier niveau constituent le socle de l’accès aux soins pour les populations.');
    M.p("La présente enquête de satisfaction (" + c.periode + ") s'inscrit dans une logique de gouvernance participative et de gestion fondée sur les preuves. Elle vise à recueillir la voix des usagers, du personnel et des membres des COGECS, à identifier les domaines à améliorer et à orienter les décisions vers des actions concrètes.");
    if (P.introExtra) M.p(P.introExtra);
    M.h(1, 'Justification');
    M.p("La satisfaction des usagers est reconnue comme un indicateur fondamental de la qualité des soins et de la performance des systèmes de santé. Elle reflète la capacité des structures sanitaires à répondre aux besoins réels des populations.");
    M.p("Dans la zone sanitaire " + c.zone + ", les formations sanitaires de premier niveau constituent le principal point d'accès aux soins. Ces structures sont confrontées à des défis persistants : insuffisance de ressources, affluence, exigences croissantes des usagers. Le personnel soignant évolue dans un environnement marqué par une charge de travail élevée et une reconnaissance parfois limitée, ce qui affecte sa motivation et la qualité des soins.");
    M.p("Dans ce contexte, l'enquête de satisfaction apparaît comme une démarche essentielle pour :");
    M.bullets(['recueillir la voix des usagers et du personnel ;', 'identifier les domaines critiques à améliorer ;', 'renforcer la redevabilité des structures sanitaires ;', "orienter les décisions locales vers des actions concrètes et ciblées."]);
    M.h(1, 'Cadre conceptuel');
    M.p("La satisfaction des usagers en milieu de soins est une notion complexe, influencée par plusieurs facteurs qui interagissent tout au long du parcours de prise en charge. Elle n'est pas une donnée brute ou spontanée mais un résultat construit à partir des interactions entre le patient et ses caractéristiques, le service de santé et ses ressources, et les dimensions de la satisfaction (accueil, information, confidentialité, compétence, coût, environnement).");
    M.p("Ce cadre sert de base méthodologique à l'élaboration des outils de mesure, à l'interprétation des résultats et à l'orientation des actions d'amélioration.");
    M.h(1, 'Objectifs');
    M.h(2, 'Objectif général', { noNum: true });
    M.p("Évaluer le niveau de satisfaction des usagers, du personnel soignant et des membres des COGECS dans les formations sanitaires de la zone sanitaire " + c.zone + " (" + c.periode + ").");
    M.h(2, 'Objectifs spécifiques', { noNum: true });
    M.bullets(["Évaluer le niveau de satisfaction des usagers dans chacun des centres de santé de la zone sanitaire ;", "Identifier les domaines de soins jugés insatisfaisants par les usagers ;", "Recueillir les besoins exprimés par le personnel soignant en lien avec la qualité des soins et leur motivation professionnelle ;", "Apprécier le fonctionnement des Comités de Gestion des Centres de Santé (COGECS) et leur collaboration avec le personnel ;", "Comparer, lorsque cela est pertinent, les résultats obtenus à ceux des périodes ou éditions précédentes ;", "Proposer des recommandations concrètes pour améliorer la qualité des soins."]);
  }
  function cadre(M, c) {
    var P = c.P;
    M.h(1, "Cadre d'étude");
    if (P.cadreTexte) { String(P.cadreTexte).split(/\n+/).forEach(function (x) { M.p(x); }); if (!/densit/i.test(P.cadreTexte) && B.chiffresCles(P)) M.p(B.chiffresCles(P)); B.extraFigs(M, P, 'cadre'); return; }
    var t = 'La zone sanitaire ' + c.zone + ' regroupe les communes de ' + (B.listFr(c.communes) || '—') + '.';
    if (P.superficie) t += ' Elle couvre une superficie estimée à ' + P.superficie + '.';
    if (P.population) { var dn = B.density(P); t += ' Sa population est estimée à ' + P.population + ' habitants' + (dn.dens ? ', soit une densité de ' + B.fN(dn.dens) + ' habitants par km²' : '') + '.'; }
    if (P.nbFsPub) t += ' Les soins y sont assurés par un réseau de ' + P.nbFsPub + ' formations sanitaires publiques' + (P.nbFsPriv ? ' et ' + P.nbFsPriv + ' privées' : '') + '.';
    M.p(t);
    if (P.limites) M.p(P.limites);
    if (P.carte) { M.p('La figure ci-dessous présente la carte sanitaire de la zone et ses limites.'); M.image('Carte sanitaire de la zone sanitaire ' + c.ZONE, P.carte); }
    B.extraFigs(M, P, 'cadre');
    M.p("L'étude s'est déroulée dans " + (P.nbCentresZone ? 'l’ensemble des ' + P.nbCentresZone : 'les') + ' centres de santé couverts par la collecte, selon des critères de localisation, de fréquentation et de diversité fonctionnelle, ainsi que sur la base des décisions de couverture prises par l’équipe de coordination de la zone.');
  }
  function methodes(M, c) {
    var P = c.P, nCom = c.communes.length, perCom = c.communes.map(function (cn) { return c.fsOf(cn).length + ' pour ' + cn; });
    M.h(1, 'Méthodologie');
    M.h(2, "Type d'étude");
    M.p("Il s'agit d'une étude transversale descriptive, menée dans la zone sanitaire " + c.zone + ", sur " + b(c.fs.length + ' centres de santé') + " (" + B.listFr(perCom) + "), au moyen d'équipes déployées en parallèle par commune. Elle décrit le niveau de satisfaction des usagers communautaires, des membres COGECS et du personnel soignant vis-à-vis des services offerts dans les formations sanitaires.");
    M.h(2, 'Population cible');
    M.p("L'étude concerne trois groupes principaux :");
    M.bullets([b('Le personnel soignant') + ' exerçant dans les formations sanitaires de la zone.', b('Les usagers communautaires') + ', c’est-à-dire les personnes vivant dans les villages et quartiers environnants ayant consulté une formation sanitaire au cours des trois (03) derniers mois, ainsi que leurs accompagnateurs.', b('Les membres du COGECS') + ' : essentiellement le président, le secrétaire et le vérificateur des centres de santé.']);
    M.p("Les usagers présents dans les centres au moment de l'enquête n'ont pas été inclus.");
    if (P.populationCible) M.p('La population cible de l’enquête est estimée à ' + b(P.populationCible) + ' personnes.');
    M.h(2, 'Méthode de recrutement des participants');
    M.p(b('Usagers communautaires. ') + "Un quota de " + c.quota + " usagers par centre de santé a été retenu. Au total " + (c.n.usagers || 0) + " usagers ont été enquêtés dans " + c.fs.length + " centres de santé.");
    M.p(b('Personnel soignant. ') + "La participation s'est faite exclusivement sur une base volontaire, sans quota prédéfini, afin d'encourager une expression libre et spontanée. " + (c.n.personnels || 0) + " agents ont répondu.");
    M.p(b('Membres COGECS. ') + "Les trois membres clés (président, secrétaire, vérificateur) ont été sollicités dans chaque centre de santé ; seuls les membres effectivement disponibles ont été retenus. " + (c.n.cogecs || 0) + " membres ont répondu.");
    M.h(2, 'Outils et techniques de collecte de données');
    M.p("La collecte a été réalisée avec la plateforme numérique de collecte de l'étude (application web installable sur smartphone ou tablette, fonctionnant hors connexion avec synchronisation automatique). Les usagers et les membres des COGECS ont été interrogés en face-à-face par des enquêteurs formés ; le personnel soignant a rempli un formulaire numérique auto-administré, accessible par un lien ou un QR code, de façon volontaire et confidentielle. Chaque questionnaire porte un identifiant unique et l'enquêteur est identifié par son numéro, ce qui permet un suivi quotidien de la qualité des données.");
    M.h(2, "Variables de l'évaluation");
    M.p("Les variables de l'évaluation sont regroupées en domaines de satisfaction :");
    KEYS.forEach(function (k) { M.p(b(CL[k] + ' (' + c.domains[k].length + ' domaines)')); M.bullets(c.domains[k].map(B.cleanDom)); });
    M.p("Chaque domaine est décliné en plusieurs questions spécifiques formulées de manière à capter les perceptions, les expériences et les jugements des répondants. L'outil de collecte reprend le questionnaire validé, afin de préserver la comparabilité des résultats d'une période à l'autre.");
    M.h(2, 'Traitement et analyse des données');
    M.p("À la fin de la collecte (ou à chaque point d'étape pendant la collecte), les données sont vérifiées afin d'assurer leur complétude et leur cohérence : un questionnaire trop court, ou commencé trop tôt après le précédent, est rejeté. L'analyse descriptive utilise des mesures de fréquence ; les scores sont calculés automatiquement par la plateforme.");
    M.p("Chaque question associée à un domaine de satisfaction est ramenée à un score maximal de 1 point. Pour les questions à deux modalités, les réponses sont cotées 0 ou 1 ; pour trois modalités, 0 – 0,5 – 1 ; pour quatre modalités, 0 – 0,33 – 0,66 – 1 ; pour cinq modalités, 0 – 0,25 – 0,50 – 0,75 – 1. Le score d'un domaine est la moyenne des scores obtenus, et la satisfaction générale d'une cible est la moyenne de l'ensemble des domaines. Le score global combiné est la moyenne des scores des cibles.");
    M.p("Le niveau de satisfaction est apprécié selon l'échelle suivante : " + b('très satisfaisant') + ' (85 % et plus), ' + b('satisfaisant') + ' (70 à 84,9 %), ' + b('moyennement satisfaisant') + ' (50 à 69,9 %) et ' + b('insatisfaisant') + ' (moins de 50 %). L’analyse se fait à trois niveaux :');
    M.bullets([b('Au niveau des centres de santé,') + ' le score de chaque domaine est calculé pour l’ensemble des participants du centre ; les domaines inférieurs au niveau moyen du centre sont considérés comme non satisfaisants.', b('Au niveau communal,') + ' le score est la moyenne des scores des centres de santé de la commune, qui sert de seuil pour distinguer les centres performants des centres non performants.', b('Au niveau de la zone sanitaire,') + ' la moyenne des scores des communes est utilisée comme seuil de référence, afin d’identifier la commune la moins satisfaite.']);
    M.h(2, 'Considérations éthiques');
    M.bullets(["Respect strict des principes éthiques liés à la collecte de données en milieu sanitaire.", "Aucune information nominative recueillie ; réponses traitées de manière anonyme et confidentielle.", "Participation entièrement volontaire ; chaque enquêté est informé de l'objectif de l'étude et de son droit de refuser ou d'interrompre sa participation à tout moment, sans conséquence sur les soins reçus.", "Enquêteurs formés à l'éthique de l'enquête : respect de la vie privée, neutralité, non-influence des réponses.", "Aucune pression exercée sur les usagers et aucun avantage proposé en échange de leur participation.", "Données utilisées exclusivement à des fins d'analyse institutionnelle et d'amélioration de la qualité des soins ; diffusion des résultats de manière agrégée."]);
    B.extraFigs(M, P, 'methodes');
    M.h(2, 'Calendrier et organisation');
    M.p("La collecte s'est déroulée " + (c.data.start ? 'du ' + B.frDate(c.data.start) + ' au ' + B.frDate(c.data.end) : 'sur la période de l’enquête') + (P.calendrierExtra ? '. ' + P.calendrierExtra : '.') + " L'enquête s'est articulée autour des étapes suivantes :");
    M.bullets(['finalisation et validation du protocole ;', 'recrutement et formation des enquêteurs, avec pré-test du questionnaire ;', 'collecte des données sur le terrain avec des équipes déployées en parallèle par commune ;', 'supervision quotidienne de la qualité des données depuis le tableau de suivi ;', 'nettoyage, apurement et analyse des données ;', 'rédaction du rapport et restitution des résultats aux acteurs de la zone sanitaire.']);
  }

  /* ---------- résultats ---------- */
  function description(M, c) {
    var d = c.data.demo, P = c.P;
    M.h(1, 'Résultats de l’enquête ' + c.edition);
    M.h(2, 'Description de la population enquêtée');
    var rows = c.communes.map(function (cn) { var n = c.main.byCommune[cn]; var u = n.usagers ? n.usagers.n : 0, p = n.personnels ? n.personnels.n : 0, g = n.cogecs ? n.cogecs.n : 0, t = u + p + g; return [cn, String(u), pc(B.pctOf(u, t)), String(p), pc(B.pctOf(p, t)), String(g), pc(B.pctOf(g, t)), String(t)]; });
    var tu = c.n.usagers || 0, tp = c.n.personnels || 0, tg = c.n.cogecs || 0;
    rows.push([c.zone, String(tu), pc(B.pctOf(tu, c.total)), String(tp), pc(B.pctOf(tp, c.total)), String(tg), pc(B.pctOf(tg, c.total)), String(c.total)]);
    M.p('Au total, ' + b(c.total) + ' personnes ont été enquêtées dans la zone sanitaire : ' + b(tu) + ' usagers (' + b(pc(B.pctOf(tu, c.total))) + '), ' + b(tp) + ' agents du personnel soignant (' + b(pc(B.pctOf(tp, c.total))) + ') et ' + b(tg) + ' membres de COGECS (' + b(pc(B.pctOf(tg, c.total))) + ').');
    M.table('Répartition des personnes enquêtées par commune et par cible (usagers, personnel soignant, COGECS), zone sanitaire ' + c.ZONE + ', ' + c.edition, ['Commune', 'Usagers', '%', 'Personnel', '%', 'COGECS', '%', 'Total'], rows, { lastBold: true });
    M.chart('Répartition des enquêtés par cible et par commune, zone sanitaire ' + c.ZONE + ', ' + c.edition, { kind: 'cluster', labels: c.communes, series: [{ name: 'Usagers', values: c.communes.map(function (cn) { return (c.main.byCommune[cn].usagers || {}).n || 0; }) }, { name: 'Personnel', values: c.communes.map(function (cn) { return (c.main.byCommune[cn].personnels || {}).n || 0; }) }, { name: 'COGECS', values: c.communes.map(function (cn) { return (c.main.byCommune[cn].cogecs || {}).n || 0; }) }], unit: '', integer: true });
    // Usagers
    M.h(3, 'Caractéristiques sociodémographiques des usagers');
    var ageMoy = c.communes.map(function (cn) { return cn + ' ' + f1((d.usagers[cn] || {}).ageMoy) + ' ans'; });
    var allAges = mean(c.communes.map(function (cn) { return (d.usagers[cn] || {}).ageMoy; }));
    M.p("L'âge moyen était de " + b(f1(allAges) + ' ans') + ' (' + ageMoy.join(', ') + ').');
    var ageT = crossTable(M, "Répartition des usagers enquêtés selon la tranche d'âge, par commune, zone sanitaire " + c.ZONE + ', ' + c.edition, c.communes, ['Moins de 25 ans', 'Entre 25 et 50 ans', 'Plus de 50 ans'], function (cn) { return d.usagers[cn].age; }, "Tranche d'âge");
    var ta = topCat(ageT.tot, ageT.gt); if (ta) M.p('La catégorie « ' + ta.k + ' » est la plus représentée parmi les usagers enquêtés (' + b(pc(ta.p)) + ' sur l’ensemble de la zone).');
    crossTable(M, 'Répartition des usagers enquêtés selon le sexe, par commune, zone sanitaire ' + c.ZONE + ', ' + c.edition, c.communes, ['Féminin', 'Masculin'], function (cn) { return d.usagers[cn].sexe; }, 'Sexe');
    crossTable(M, "Répartition des usagers enquêtés selon le niveau d'instruction, par commune, zone sanitaire " + c.ZONE + ', ' + c.edition, c.communes, ['Non scolarisé', 'Alphabétisé', 'Primaire', 'Secondaire', 'Universitaire'], function (cn) { return d.usagers[cn].scolarite; }, "Niveau d'instruction");
    var pro = {}; c.communes.forEach(function (cn) { var o = d.usagers[cn].profession || {}; Object.keys(o).forEach(function (k) { pro[k] = (pro[k] || 0) + o[k]; }); });
    var proCats = Object.keys(pro).sort(function (x, y) { return pro[y] - pro[x]; });
    if (proCats.length) crossTable(M, 'Répartition des usagers enquêtés selon la profession, par commune, zone sanitaire ' + c.ZONE + ', ' + c.edition, c.communes, proCats, function (cn) { return d.usagers[cn].profession; }, 'Profession');
    // Personnel
    M.h(3, 'Caractéristiques sociodémographiques du personnel soignant');
    var ageP = mean(c.communes.map(function (cn) { return (d.personnels[cn] || {}).ageMoy; }));
    M.p("L'âge moyen du personnel soignant enquêté était de " + b(f1(ageP) + ' ans') + '.');
    crossTable(M, 'Répartition du personnel soignant enquêté selon le sexe, par commune, zone sanitaire ' + c.ZONE + ', ' + c.edition, c.communes, ['Féminin', 'Masculin'], function (cn) { return d.personnels[cn].sexe; }, 'Sexe');
    crossTable(M, "Répartition du personnel soignant enquêté selon le niveau d'instruction, par commune, zone sanitaire " + c.ZONE + ', ' + c.edition, c.communes, ['Non scolarisé', 'Alphabétisé', 'Primaire', 'Secondaire', 'Universitaire'], function (cn) { return d.personnels[cn].scolarite; }, "Niveau d'instruction");
  }
  function communeSection(M, c, cn, idx) {
    var n = '2.' + (idx + 1), cm = c.main.byCommune[cn], fs = c.fsOf(cn), sub = 0;
    M.h(2, 'Présentation des données de satisfaction dans la commune ' + c.com(cn));
    KEYS.forEach(function (k) {
      if (!cm[k] || !cm[k].n) return;
      M.h(3, CL[k]);
      M.p('Le tableau ci-après présente, domaine par domaine, le score de satisfaction obtenu par chaque formation sanitaire de la commune ' + c.com(cn) + ' auprès ' + DEL[k] + '. La moyenne de ces formations sanitaires constitue le niveau de satisfaction de la commune pour le domaine concerné.');
      c.domains[k].forEach(function (dom, i) {
        var rows = fsRows(c, k, cn, dom), avg = cm[k].d[dom];
        M.h(4, (i + 1) + '. ' + B.cleanDom(dom));
        M.p(B.domainSentence('la commune ' + c.com(cn), avg, rows, B.UNIT_FS));
        if ((c.P.commentaires || {})[k + '|' + cn + '|' + dom]) M.p(c.P.commentaires[k + '|' + cn + '|' + dom]);
        M.table('Score de satisfaction ' + DEL[k] + ' pour le domaine « ' + B.cleanDom(dom) + ' », par centre de santé, commune ' + c.com(cn) + ', ' + c.edition, ['Formation sanitaire', 'N', 'Score (%)'], rows.map(function (r) { return [r.name, String(r.n), f1(r.score)]; }).concat([['Moyenne commune ' + c.com(cn), '', f1(avg)]]), { lastBold: true });
      });
      var scores = c.domains[k].map(function (d) { return cm[k].d[d]; });
      M.h(3, 'Appréciation globale du niveau de satisfaction ' + DE[k] + ' dans la commune ' + c.com(cn));
      M.p(B.globalAppraisal('dans la commune ' + c.com(cn), 'communale', c.domains[k], scores));
      M.chart('Appréciation globale de la satisfaction ' + DEL[k] + ' par domaine, commune ' + c.com(cn) + ', ' + c.edition, { kind: 'bar', labels: c.domains[k].map(B.cleanDom), values: scores, mean: B.mean(scores), horizontal: true });
    });
    // global commune
    M.h(3, 'Niveau de satisfaction globale de la commune ' + c.com(cn));
    var rows = fs.map(function (s) { var g = c.main.byStructure[s]; return { name: s, u: g.usagers ? g.usagers.o : null, g: g.cogecs ? g.cogecs.o : null, p: g.personnels ? g.personnels.o : null, c: g.combined }; });
    M.p('Le score global de la commune ' + c.com(cn) + ' est de ' + b(pc(cm.combined)) + ' (' + b(B.qual(cm.combined)) + '), moyenne des scores des usagers (' + pc((cm.usagers || {}).o) + '), du personnel soignant (' + pc((cm.personnels || {}).o) + ') et des COGECS (' + pc((cm.cogecs || {}).o) + ').');
    M.table('Niveau de satisfaction globale, par centre de santé, commune ' + c.com(cn) + ', ' + c.edition, ['Formation sanitaire', 'Usagers', 'COGECS', 'Personnel', 'Score global (%)'], rows.map(function (r) { return [r.name, f1(r.u), f1(r.g), f1(r.p), f1(r.c)]; }).concat([['Commune ' + cn, f1((cm.usagers || {}).o), f1((cm.cogecs || {}).o), f1((cm.personnels || {}).o), f1(cm.combined)]]), { lastBold: true });
    var avgFs = B.mean(rows.map(function (r) { return r.c; }));
    M.chart('Niveau de satisfaction globale par centre de santé, commune ' + c.com(cn) + ', ' + c.edition, { kind: 'bar', labels: rows.map(function (r) { return r.name; }), values: rows.map(function (r) { return r.c; }), mean: avgFs, horizontal: true });
  }
  function zoneSection(M, c, no) {
    var zn = c.main.zone, cs = c.communes;
    M.h(2, 'Présentation des données de satisfaction dans la zone sanitaire');
    KEYS.forEach(function (k) {
      if (!zn[k] || !zn[k].n) return;
      M.h(3, CL[k]);
      M.p('Le tableau ci-après présente, domaine par domaine, les scores de satisfaction obtenus ' + PAR[k] + ' au niveau de la zone sanitaire. Le score de la zone est la moyenne des scores des communes.');
      var rows = c.domains[k].map(function (dom) { var line = [B.cleanDom(dom)]; cs.forEach(function (cn) { line.push(f1((c.main.byCommune[cn][k] || {}).d ? c.main.byCommune[cn][k].d[dom] : null)); }); line.push(f1(zn[k].d[dom])); return line; });
      rows.push(['Moyenne'].concat(cs.map(function (cn) { return f1((c.main.byCommune[cn][k] || {}).o); }), [f1(zn[k].o)]));
      M.table('Score de satisfaction ' + DEL[k] + ' par domaine et par commune, zone sanitaire ' + c.ZONE + ', ' + c.edition, ['Domaine'].concat(cs, ['Zone sanitaire']), rows, { lastBold: true });
      c.domains[k].forEach(function (dom, i) {
        var cr = cs.map(function (cn) { return { name: 'la commune ' + c.com(cn), score: (c.main.byCommune[cn][k] || { d: {} }).d[dom] }; });
        M.p(b((i + 1) + '. ' + B.cleanDom(dom)) + ' — ' + B.domainSentence('la zone sanitaire', zn[k].d[dom], cr, B.UNIT_COM));
      });
      var scores = c.domains[k].map(function (d) { return zn[k].d[d]; });
      M.h(3, 'Appréciation globale du niveau de satisfaction ' + DE[k] + ' dans la zone sanitaire');
      M.p(B.globalAppraisal('dans la zone sanitaire', 'de la zone', c.domains[k], scores));
      M.chart('Appréciation globale de la satisfaction ' + DEL[k] + ' par domaine, zone sanitaire ' + c.ZONE + ', ' + c.edition, { kind: 'bar', labels: c.domains[k].map(B.cleanDom), values: scores, mean: B.mean(scores), horizontal: true });
    });
  }
  function synthese(M, c) {
    M.h(2, 'Synthèse comparative et classement des formations sanitaires');
    var cs = c.communes, zn = c.main.zone, sc = cs.map(function (cn) { return { name: cn, s: c.main.byCommune[cn].combined }; }), srt = B.sortDesc(sc.map(function (x) { return { name: x.name, score: x.s }; }), 'score');
    M.h(3, 'Score global de satisfaction par commune');
    if (srt.length > 1) M.p('Le score global combiné va de ' + b(pc(srt[srt.length - 1].score)) + ' à ' + srt[srt.length - 1].name + ' à ' + b(pc(srt[0].score)) + ' à ' + srt[0].name + ', soit un écart de ' + f1(srt[0].score - srt[srt.length - 1].score) + ' points.');
    M.table('Score global de satisfaction combiné, par commune, zone sanitaire ' + c.ZONE + ', ' + c.edition, ['Commune', 'Score global (%)'], sc.map(function (x) { return [x.name, f1(x.s)]; }), {});
    M.chart('Score global de satisfaction par commune, zone sanitaire ' + c.ZONE + ', ' + c.edition, { kind: 'bar', labels: cs, values: sc.map(function (x) { return x.s; }), mean: zn.combined });
    M.h(3, 'Score global de satisfaction par cible pour la zone sanitaire');
    var cib = KEYS.filter(function (k) { return zn[k]; }).map(function (k) { return { k: k, s: zn[k].o }; }), cs2 = B.sortDesc(cib.map(function (x) { return { name: CC[x.k], score: x.s }; }), 'score');
    var t = 'Le score combiné de satisfaction de la zone sanitaire s’établit à ' + b(pc(zn.combined)) + ' (' + b(B.qual(zn.combined)) + ').';
    if (cs2.length > 1) t += ' La cible la plus satisfaite est « ' + cs2[0].name + ' » (' + b(pc(cs2[0].score)) + ') et la moins satisfaite « ' + cs2[cs2.length - 1].name + ' » (' + b(pc(cs2[cs2.length - 1].score)) + ').';
    M.p(t);
    M.table('Score global de satisfaction combiné et par cible, zone sanitaire ' + c.ZONE + ', ' + c.edition, ['Cible', 'Score global ZS (%)'], cib.map(function (x) { return [CC[x.k], f1(x.s)]; }).concat([['Score combiné', f1(zn.combined)]]), { lastBold: true });
    M.chart('Satisfaction globale par cible, zone sanitaire ' + c.ZONE + ', ' + c.edition, { kind: 'bar', labels: cib.map(function (x) { return CC[x.k]; }).concat(['Score combiné']), values: cib.map(function (x) { return x.s; }).concat([zn.combined]), mean: zn.combined });
    M.h(3, 'Classement des formations sanitaires');
    var rk = c.fs.map(function (s) { var g = c.main.byStructure[s]; return { name: s, com: g.commune, u: g.usagers ? g.usagers.o : null, g: g.cogecs ? g.cogecs.o : null, p: g.personnels ? g.personnels.o : null, score: g.combined }; });
    rk = B.sortDesc(rk, 'score');
    if (rk.length > 2) M.p('La formation sanitaire la mieux classée est ' + rk[0].name + ' (' + rk[0].com + ', ' + b(pc(rk[0].score)) + ') ; la moins bien classée est ' + rk[rk.length - 1].name + ' (' + rk[rk.length - 1].com + ', ' + b(pc(rk[rk.length - 1].score)) + '). ' + rk.filter(function (r) { return isNum(r.score) && r.score < zn.combined; }).length + ' formations sanitaires sur ' + rk.length + ' sont en dessous de la moyenne de la zone (' + pc(zn.combined) + ').');
    M.table('Classement des formations sanitaires selon le score global de satisfaction, zone sanitaire ' + c.ZONE + ', ' + c.edition, ['Rang', 'Formation sanitaire', 'Commune', 'Usagers', 'COGECS', 'Personnel', 'Score global (%)'], rk.map(function (r, i) { return [String(i + 1), r.name, r.com, f1(r.u), f1(r.g), f1(r.p), f1(r.score)]; }), {});
  }
  /** Évolution : comparaison à une édition précédente (data.compare) et/ou entre périodes du rapport (data.byPeriod). */
  function evolution(M, c) {
    var d = c.data, cmp = d.compare, bp = d.byPeriod || [], cs = c.communes, zn = c.main.zone;
    var cols = [];
    if (cmp && cmp.sum && cmp.sum.zone && isNum(cmp.sum.zone.combined)) cols.push({ label: cmp.labels.join(' + '), sum: cmp.sum });
    bp.forEach(function (p) { cols.push({ label: p.label, sum: p.sum }); });
    if (!cols.length) return;
    M.h(2, cmp ? 'Comparaison des niveaux de satisfaction' : 'Évolution du niveau de satisfaction entre les périodes');
    var heads = cols.map(function (x) { return x.label + ' (%)'; });
    M.p('Cette section compare les niveaux de satisfaction obtenus ' + (cmp ? 'à ceux de ' + cmp.labels.join(' + ') : 'entre les périodes du rapport (' + bp.map(function (p) { return p.label; }).join(', ') + ')') + ', par commune et pour la zone sanitaire, puis par cible enquêtée.');
    M.h(3, 'Évolution par commune et pour la zone sanitaire');
    var rows = cs.map(function (cn) { var line = [cn]; cols.forEach(function (x) { var bc = x.sum.byCommune[cn]; line.push(f1(bc ? bc.combined : null)); }); line.push(f1(c.main.byCommune[cn].combined)); line.push(cols.length ? delta(c.main.byCommune[cn].combined, cols[0].sum.byCommune[cn] ? cols[0].sum.byCommune[cn].combined : null) : '—'); return line; });
    rows.push(['Zone sanitaire ' + c.zone].concat(cols.map(function (x) { return f1(x.sum.zone.combined); }), [f1(zn.combined), delta(zn.combined, cols[0].sum.zone.combined)]));
    var first = cols[0], dz = isNum(zn.combined) && isNum(first.sum.zone.combined) ? zn.combined - first.sum.zone.combined : null;
    if (isNum(dz)) M.p('À l’échelle de la zone sanitaire, le niveau de satisfaction passe de ' + b(pc(first.sum.zone.combined)) + ' (' + first.label + ') à ' + b(pc(zn.combined)) + ' (' + (bp.length ? 'ensemble de la période' : c.periode) + '), soit une ' + (dz >= 0 ? 'progression' : 'baisse') + ' de ' + b(f1(Math.abs(dz)) + ' points') + '.');
    M.table('Évolution du niveau de satisfaction globale par commune et pour la zone sanitaire ' + c.ZONE, ['Commune / Zone'].concat(heads, [(bp.length ? 'Ensemble' : c.edition) + ' (%)', 'Évolution (points)']), rows, { lastBold: true });
    M.chart('Évolution du niveau de satisfaction par commune et zone sanitaire', { kind: 'cluster', labels: cs.concat(['Zone']), series: cols.map(function (x) { return { name: x.label, values: cs.map(function (cn) { var bc = x.sum.byCommune[cn]; return bc ? bc.combined : null; }).concat([x.sum.zone.combined]) }; }).concat([{ name: bp.length ? 'Ensemble' : c.edition, values: cs.map(function (cn) { return c.main.byCommune[cn].combined; }).concat([zn.combined]) }]), unit: ' %', max: 100 });
    M.h(3, 'Évolution par cible enquêtée');
    var rows2 = KEYS.filter(function (k) { return zn[k]; }).map(function (k) { var line = [CC[k]]; cols.forEach(function (x) { line.push(f1(x.sum.zone[k] ? x.sum.zone[k].o : null)); }); line.push(f1(zn[k].o)); line.push(delta(zn[k].o, cols[0].sum.zone[k] ? cols[0].sum.zone[k].o : null)); return line; });
    M.table('Évolution du niveau de satisfaction globale par cible enquêtée, zone sanitaire ' + c.ZONE, ['Cible'].concat(heads, [(bp.length ? 'Ensemble' : c.edition) + ' (%)', 'Évolution (points)']), rows2, {});
    var best = KEYS.filter(function (k) { return zn[k] && cols[0].sum.zone[k]; }).map(function (k) { return { name: CC[k], score: zn[k].o - cols[0].sum.zone[k].o }; });
    if (best.length > 1) { best = B.sortDesc(best, 'score'); M.p('La plus forte évolution concerne « ' + best[0].name + ' » (' + (best[0].score >= 0 ? '+' : '') + f1(best[0].score) + ' points) ; la plus faible concerne « ' + best[best.length - 1].name + ' » (' + (best[best.length - 1].score >= 0 ? '+' : '') + f1(best[best.length - 1].score) + ' points).'); }
    M.chart('Évolution du niveau de satisfaction par cible enquêtée', { kind: 'cluster', labels: KEYS.filter(function (k) { return zn[k]; }).map(function (k) { return CC[k]; }), series: cols.map(function (x) { return { name: x.label, values: KEYS.filter(function (k) { return zn[k]; }).map(function (k) { return x.sum.zone[k] ? x.sum.zone[k].o : null; }) }; }).concat([{ name: bp.length ? 'Ensemble' : c.edition, values: KEYS.filter(function (k) { return zn[k]; }).map(function (k) { return zn[k].o; }) }]), unit: ' %', max: 100 });
  }
  function recoList(c) {
    if (c.P.recos && c.P.recos.length) return c.P.recos;
    var zn = c.main.zone, out = [], seen = {};
    KEYS.forEach(function (k) {
      if (!zn[k]) return;
      var low = c.domains[k].map(function (d) { return { d: d, s: zn[k].d[d] }; }).filter(function (x) { return isNum(x.s); }).sort(function (a, z) { return a.s - z.s; }).slice(0, 2);
      low.forEach(function (x) { var r = B.recoFor(x.d); if (r && !seen[r]) { seen[r] = 1; out.push(r); } });
    });
    out.push('Partager les résultats de ce rapport avec chaque formation sanitaire pour l’élaboration d’un plan d’amélioration local.');
    out.push('Reconduire cette évaluation régulièrement afin de suivre l’évolution de la satisfaction dans le temps, par commune et par formation sanitaire.');
    return out;
  }
  function conclusion(M, c) {
    var zn = c.main.zone, parts = KEYS.filter(function (k) { return zn[k]; }).map(function (k) { return ({ usagers: 'les usagers', personnels: 'le personnel soignant', cogecs: 'les membres des COGECS' })[k] + ' (' + b(pc(zn[k].o)) + ')'; });
    M.h(1, 'Conclusion', { noNum: true });
    if (c.P.conclusionTexte) { String(c.P.conclusionTexte).split(/\n+/).forEach(function (x) { M.p(x); }); }
    else {
      var rk = B.sortDesc(c.fs.map(function (s) { return { name: s, score: c.main.byStructure[s].combined }; }), 'score');
      M.p("L'évaluation conduite dans la zone sanitaire " + c.zone + " (" + c.periode + ") révèle un niveau de satisfaction global " + b(B.qual(zn.combined)) + ' (' + b(pc(zn.combined)) + '), avec les résultats suivants par cible : ' + B.listFr(parts) + '.');
      if (rk.length > 2) M.p("L'analyse par commune et par formation sanitaire met en évidence des disparités : les scores vont de " + pc(rk[rk.length - 1].score) + ' (' + rk[rk.length - 1].name + ') à ' + pc(rk[0].score) + ' (' + rk[0].name + '). Les efforts doivent porter prioritairement sur les domaines les moins bien notés et sur les formations sanitaires en dessous de la moyenne de la zone.');
    }
    B.extraFigs(M, c.P, 'resultats');
    M.h(1, 'Recommandations', { noNum: true });
    M.bullets(recoList(c));
    M.h(1, 'Annexes', { noNum: true });
    B.extraFigs(M, c.P, 'annexes');
    M.p('Domaines de satisfaction évalués :');
    KEYS.forEach(function (k) { M.p(b(CL[k])); M.bullets(c.domains[k].map(B.cleanDom)); });
  }


  var MOIS = ['Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin', 'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'];
  function moisAnnee(iso) { var d = iso ? new Date(iso) : new Date(); if (isNaN(d)) d = new Date(); return MOIS[d.getMonth()] + ' ' + d.getFullYear(); }
  var NOMBRE = { 1: 'un', 2: 'deux', 3: 'trois', 4: 'quatre', 5: 'cinq', 6: 'six', 7: 'sept', 8: 'huit', 9: 'neuf', 10: 'dix', 11: 'onze', 12: 'douze', 13: 'treize' };
  function resumeZone(c) {
    var zn = c.main.zone, d = c.data, P = c.P, cmp = d.compare && d.compare.sum && d.compare.sum.zone && isNum(d.compare.sum.zone.combined) ? d.compare : null;
    var tot = c.total, cs = c.communes;
    var cib = KEYS.filter(function (k) { return zn[k] && isNum(zn[k].o); }).map(function (k) { return { k: k, o: zn[k].o }; }).sort(function (a, z) { return z.o - a.o; });
    var NOM = { usagers: 'les usagers', personnels: 'le personnel soignant', cogecs: 'les membres des COGECS' };
    var low = cib.length ? cib[cib.length - 1] : null, lowDom = null;
    if (low) { var ds = c.domains[low.k].map(function (x) { return { n: B.cleanDom(x), s: zn[low.k].d[x] }; }).filter(function (x) { return isNum(x.s); }).sort(function (a, z) { return a.s - z.s; }); lowDom = ds[0]; }
    var prog = cmp ? (zn.combined >= cmp.sum.zone.combined ? 'en progression' : 'en baisse') + ' par rapport à ' + cmp.labels.join(' + ') + ' (' + b(pc(cmp.sum.zone.combined)) + '), ' : '';
    var cpart = cs.map(function (cn) { return b(pc(c.main.byCommune[cn].combined)) + ' à ' + cn; });
    var r = [];
    r.push({ label: 'Introduction', text: "La satisfaction des usagers et du personnel de santé est un indicateur clé de performance et de qualité des services. Dans la zone sanitaire " + cs.join('–') + ' (' + c.ZONE + "), l'évaluation de cette satisfaction constitue un outil stratégique pour orienter les décisions, renforcer la confiance des populations et améliorer la gouvernance locale." });
    r.push({ label: 'Objectifs', text: "L'étude " + c.edition + " vise à mesurer le niveau de satisfaction des usagers communautaires, du personnel soignant et des membres des COGECS, à identifier les domaines prioritaires d'amélioration et à comparer les résultats entre communes, cibles" + (cmp ? ' et avec ' + cmp.labels.join(' + ') : '') + '.' });
    r.push({ label: 'Méthodes', text: "Une enquête transversale a été menée en " + c.edition + " dans " + c.fs.length + " centres de santé de la zone, auprès de " + b(B.fN(tot)) + " participants, répartis entre les usagers (" + B.fN(c.n.usagers || 0) + ', soit ' + pc(B.pctOf(c.n.usagers || 0, tot)) + '), le personnel soignant (' + B.fN(c.n.personnels || 0) + ', soit ' + pc(B.pctOf(c.n.personnels || 0, tot)) + ') et les membres des COGECS (' + (c.n.cogecs || 0) + ', soit ' + pc(B.pctOf(c.n.cogecs || 0, tot)) + '). Les données ont été collectées par questionnaires structurés couvrant ' + (NOMBRE[c.domains.usagers.length] || c.domains.usagers.length) + ' domaines de satisfaction pour les usagers, ' + (NOMBRE[c.domains.personnels.length] || c.domains.personnels.length) + ' pour le personnel soignant et ' + (NOMBRE[c.domains.cogecs.length] || c.domains.cogecs.length) + ' pour les membres des COGECS.' });
    r.push({ label: 'Résultats', text: "L'enquête de satisfaction menée en " + c.edition + ' dans la zone sanitaire ' + c.ZONE + ' révèle un niveau global de satisfaction combiné de ' + b(pc(zn.combined)) + ', ' + prog + 'avec ' + (cs.length > 1 ? 'des disparités entre communes : ' + B.listFr(cpart) : 'un résultat de ' + cpart.join('')) + '.' + (cib.length > 1 ? ' Par cible, ' + B.listFr(cib.slice(0, cib.length - 1).map(function (x) { return NOM[x.k] + ' (' + b(pc(x.o)) + ')'; })) + ' affichent les niveaux de satisfaction les plus élevés, tandis que ' + NOM[low.k] + ' reste la cible la moins satisfaite (' + b(pc(low.o)) + ')' + (lowDom ? ', notamment sur le domaine « ' + lowDom.n + ' » (' + pc(lowDom.s) + ')' : '') + '.' : '') });
    r.push({ label: 'Conclusion', text: "L'édition " + c.edition + ' met en évidence un niveau de satisfaction ' + b(B.qual(zn.combined)) + ' dans la zone sanitaire ' + c.ZONE + (cmp ? ', ' + prog.replace(/, $/, '') : '') + '. ' + (low ? cap1(NOM[low.k]) + ' demeure la cible prioritaire pour les actions d’amélioration' + (lowDom ? ', en particulier sur le domaine « ' + lowDom.n + ' »' : '') + '.' : '') });
    return r;
  }
  function cap1(s) { return s.charAt(0).toUpperCase() + s.slice(1); }
  function buildZone(data, P) {
    var c = ctx(data, P), M = new B.Model();
    intro(M, c); cadre(M, c); methodes(M, c); description(M, c);
    c.communes.forEach(function (cn, i) { communeSection(M, c, cn, i); });
    zoneSection(M, c); synthese(M, c); evolution(M, c); conclusion(M, c);
    return { model: M, ctx: c, meta: { kind: 'zone', title: 'ÉVALUATION DE LA SATISFACTION DES USAGERS DE LA ZONE SANITAIRE ' + c.ZONE, sub: 'Édition ' + c.edition + ' — ' + c.periode, edition: c.edition, consultant: (P && P.consultant) || 'Ets MOUREDE & Associés', commanditaire: (P && P.commanditaire) || 'Administration de la Zone Sanitaire', presentateur: (P && P.presentateur) || 'Équipe de Consultants', draft: !!(P && P.statut !== 'final'), generatedAt: data.generatedAt, structure: c.zone, scope: 'zone', logo: (P && P.logo) || '', cover: { direction: 'DIRECTION DEPARTEMENTALE DE LA SANTE DU ' + String((P && P.departement) || '…').toUpperCase(), ligne2: 'ZONE SANITAIRE ' + String((P && P.nomCouverture) || c.communes.join(' – ') || c.zone).toUpperCase(), titre: 'RAPPORT D’ENQUETE SUR LA SATISFACTION DES USAGERS ET DU PERSONNEL DANS LES FORMATIONS SANITAIRES', date: (P && P.dateRapport) || moisAnnee(data.generatedAt) }, fontSz: 24 }, recos: recoList(c), resume: resumeZone(c), motsCles: 'Évaluation, satisfaction, performance sanitaire, ZS-' + c.zone };
  }
  root.RapportZone = { build: buildZone, ctx: ctx, fsRows: fsRows, recoList: recoList, KEYS: KEYS };
})(typeof window !== 'undefined' ? window : globalThis);
/* ===== Hôpital de zone : modèle de rapport au canevas 2026 (plan des rapports HZ Kandi / Malanville / Pobè / Bassila 2026) ===== */
(function (root) {
  'use strict';
  var B = root.RapportBase, f1 = B.f1, pc = B.pc, b = B.b, mean = B.mean, isNum = B.isNum, delta = B.delta;
  var HK = ['usagers', 'personnel', 'communaute'];
  var HL = { usagers: 'Usagers de l’hôpital', personnel: 'Personnel de l’hôpital', communaute: 'Communauté' };
  var HDE = { usagers: 'des usagers', personnel: 'du personnel', communaute: 'de la communauté' };
  var MOIS = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  var NW = { 0: 'zéro', 1: 'un', 2: 'deux', 3: 'trois', 4: 'quatre', 5: 'cinq', 6: 'six', 7: 'sept', 8: 'huit', 9: 'neuf', 10: 'dix', 11: 'onze', 12: 'douze', 13: 'treize', 14: 'quatorze', 15: 'quinze' };
  function nw(n) { return (NW[n] !== undefined ? NW[n] : String(n)) + ' (' + (n < 10 ? '0' : '') + n + ')'; }
  function lab(c) { var L = root.HLABELS || {}; return (L[c] || c).replace(/\s+/g, ' ').replace(/\?$/, ' ?'); }
  function periodTxt(d) {
    if (!d.start) return '';
    var a = String(d.start).split('-'), z = String(d.end || d.start).split('-');
    if (a[0] === z[0] && a[1] === z[1]) return 'du ' + Number(a[2]) + ' au ' + Number(z[2]) + ' ' + MOIS[Number(z[1]) - 1] + ' ' + z[0];
    return 'du ' + Number(a[2]) + ' ' + MOIS[Number(a[1]) - 1] + ' ' + a[0] + ' au ' + Number(z[2]) + ' ' + MOIS[Number(z[1]) - 1] + ' ' + z[0];
  }
  function ctx(data, P) {
    P = P || {};
    var c = { data: data, P: P, ci: data.cibles };
    c.name = P.nomAffiche || P.nomHopital || data.hopital || 'l’hôpital';
    c.nm = c.name.replace(/^(hz|h\.z\.|hôpital de zone|hôpital)[\s-]*(de\s+|d['’]\s*)?/i, '').trim() || c.name;
    c.HZ = 'HZ-' + c.nm; c.NAME = c.nm.toUpperCase();
    c.zone = (P.nomZoneSanitaire || P.zone || '').replace(/^(zone sanitaire|zs)\s+/i, '').trim() || c.nm;
    c.edition = P.edition || (data.end ? String(data.end).slice(0, 4) : String(new Date().getFullYear()));
    c.periode = P.periodeTexte || periodTxt(data) || (data.periodLabels && data.periodLabels.length ? data.periodLabels.join(' + ') : c.edition);
    c.n = {}; HK.forEach(function (k) { c.n[k] = c.ci[k] ? c.ci[k].n : 0; });
    c.total = c.n.usagers + c.n.personnel + c.n.communaute;
    c.sub = function (t) { return String(t).replace(/\{HZ\}/g, c.HZ).replace(/\{NOM\}/g, c.nm).replace(/\{ZONE\}/g, c.zone).replace(/\{PERIODE\}/g, c.periode).replace(/\{ANNEE\}/g, c.edition); };
    return c;
  }
  function emit(M, c, arr) { (arr || []).forEach(function (x) { if (x.t === 'ul') M.bullets(x.items.map(c.sub)); else M.p(c.sub(x.text)); }); }
  function cnt(o, k) { return o && o[k] ? o[k] : 0; }
  function sumVals(o) { return Object.keys(o || {}).reduce(function (s, k) { return s + o[k]; }, 0); }
  function top(o) { var k = Object.keys(o || {}).sort(function (a, z) { return o[z] - o[a]; })[0]; return k; }

  /* tableau « profil » : groupes Ambulatoire / Hospitalisé */
  function profilTable(M, c, cap, A, H, groups) {
    var head = ['', 'Ambulatoire (n= ' + A.n + ') n', '%', 'Hospitalisé (n= ' + H.n + ') n', '%'], rows = [], bold = [];
    groups.forEach(function (g) {
      var cats = g.cats || Array.from(new Set(Object.keys(A[g.key] || {}).concat(Object.keys(H[g.key] || {})))).sort(function (x, z) { return ((H[g.key] || {})[z] || 0) + ((A[g.key] || {})[z] || 0) - (((H[g.key] || {})[x] || 0) + ((A[g.key] || {})[x] || 0)); });
      if (!cats.length) return;
      bold.push(rows.length); rows.push([g.title, '', '', '', '']);
      var ta = sumVals(A[g.key]), th = sumVals(H[g.key]);
      cats.forEach(function (k) { var a = cnt(A[g.key], k), h = cnt(H[g.key], k); rows.push([k, String(a), f1(B.pctOf(a, ta)), String(h), f1(B.pctOf(h, th))]); });
    });
    if (rows.length) M.table(cap, head, rows, { boldRows: bold });
  }
  function simpleTable(M, cap, first, obj, order) {
    var cats = order || Object.keys(obj || {}).sort(function (x, z) { return obj[z] - obj[x]; }), t = sumVals(obj);
    if (!t) return null;
    M.table(cap, [first, 'Effectif', '%'], cats.filter(function (k) { return obj[k]; }).map(function (k) { return [k, String(obj[k]), f1(B.pctOf(obj[k], t))]; }).concat([['Total', String(t), '100,0']]), { lastBold: true });
    return { top: top(obj), p: B.pctOf(obj[top(obj)], t), t: t };
  }

  function buildHosp(data, P) {
    P = P || {};
    var c = ctx(data, P), M = new B.Model(), ci = c.ci, T = root.HZT || {}, HZ = c.HZ, Y = c.edition;
    var cmp = data.compare && isNum(data.compare.combined) ? data.compare : null;
    var bp = data.byPeriod || [];
    var u = ci.usagers, split = u && u.hospitalise && u.hospitalise.n && u.ambulatoire && u.ambulatoire.n;
    /* --- Introduction --- */
    M.h(1, 'Introduction', { noNum: true });
    if (T.intro1) M.p(c.sub(T.intro1));
    M.p("A cet effet, l'Hôpital de zone de " + c.nm + " a initié pour le compte de l'année " + Y + ", la réalisation d'une enquête de satisfaction des clients afin de prouver scientifiquement les faiblesses et les domaines de soins à améliorer pour le bien-être des populations. La présente enquête permettra aussi de voir non seulement le niveau de satisfaction chez les usagers mais également d'apprécier le niveau de motivation du personnel.");
    if (P.introExtra) M.p(P.introExtra);
    M.h(1, 'Cadre conceptuel de la satisfaction des usagers en milieu de soin');
    emit(M, c, T.cadre);
    /* --- Objectifs --- */
    M.h(1, 'Objectifs');
    M.h(2, 'Objectif général');
    M.p("Evaluer le niveau de satisfaction des usagers de l'HZ de " + c.nm + ' en ' + Y + '.');
    M.h(2, 'Objectifs spécifiques');
    M.bullets(["Déterminer le niveau de satisfaction des usagers (hospitalisés, ambulatoires et des communautés environnantes) sur les prestations de soins à l'HZ de " + c.nm + ' ;', "Déterminer les domaines d'insatisfaction des usagers à l'HZ de " + c.nm + ' ;', 'Analyser les variations de satisfaction dans le temps ;', 'Identifier quelques facteurs influençant la satisfaction des clients ;', "Proposer des approches d'amélioration de la satisfaction de l'HZ de " + c.nm + '.']);
    /* --- Généralités --- */
    M.h(1, 'Généralités');
    M.h(2, 'Concept de satisfaction'); emit(M, c, T.concept);
    M.h(2, 'Mesure et évaluation de la satisfaction'); emit(M, c, T.mesure);
    /* --- Cadre et méthodes --- */
    M.h(1, "Cadre et méthodes d'étude");
    M.h(2, "Cadre d'étude");
    M.h(3, 'Présentation de la zone sanitaire ' + c.zone);
    if (P.presZone) { String(P.presZone).split(/\n+/).forEach(function (x) { M.p(x); }); if (!/densit/i.test(P.presZone) && B.chiffresCles(P)) M.p('La zone sanitaire ' + c.zone + ' ' + B.chiffresCles(P).replace(/^Elle présente/, 'présente')); }
    else M.p('La zone sanitaire ' + c.zone + (P.departement ? " est l'une des zones sanitaires du département " + (/^[AEIOUYÉ]/i.test(P.departement) ? "de l'" : 'du ') + P.departement : '') + '.' + (P.superficie ? ' Sa superficie est de ' + P.superficie + '.' : '') + (P.population ? ' Sa population est estimée à ' + P.population + ' habitants' + (B.density(P).dens ? ', soit une densité de ' + B.fN(B.density(P).dens) + ' habitants par km²' : '') + '.' : '') + (P.nbFsPub ? ' Elle compte ' + P.nbFsPub + ' formations sanitaires publiques, dont l’hôpital de zone' + (P.nbFsPriv ? ', et ' + P.nbFsPriv + ' formations sanitaires privées' : '') + '.' : ''));
    if (P.carte) { M.p('La figure ci-dessous présente la carte sanitaire de la zone et ses limites.'); M.image('Carte géographique de la Zone Sanitaire ' + c.zone, P.carte); }
    M.h(3, "Présentation de l'hôpital de zone " + c.zone);
    if (P.presHopital) String(P.presHopital).split(/\n+/).forEach(function (x) { M.p(x); });
    else M.p("L'Hôpital de Zone de " + c.zone + ' est un établissement hospitalier de première référence pour la zone sanitaire' + (P.departement ? ', installé dans le département ' + (/^[AEIOUYÉ]/i.test(P.departement) ? "de l'" : 'du ') + P.departement : '') + '.');
    if (P.photo) M.image("Hôpital de zone de " + c.nm, P.photo);
    B.extraFigs(M, P, 'cadre');
    M.h(2, "Méthodes d'étude");
    if (P.populationCible) M.p('La population cible de l’enquête est estimée à ' + b(P.populationCible) + ' personnes.');
    M.h(3, "Type et période d'étude"); M.p(c.sub((T.type || [''])[0]));
    M.h(3, "Population d'étude"); emit(M, c, T.population);
    M.h(3, "Taille de l'échantillon"); emit(M, c, T.taille);
    M.h(3, "Technique d'échantillonnage et de collecte de données"); emit(M, c, T.technique);
    M.h(3, "Variables de l'étude");
    emit(M, c, T.variables);
    M.p("La satisfaction est évaluée sur les domaines suivants :");
    HK.forEach(function (k) { if (ci[k]) { M.p(b(HL[k] + ' (' + nw(ci[k].cats.length) + ' domaines)')); M.bullets(ci[k].cats); } });
    if (T.ponderation) M.p(c.sub(T.ponderation));
    M.h(3, 'Traitement des données'); emit(M, c, T.traitement);
    M.h(3, 'Analyse de données');
    M.h(4, 'Evaluation de la satisfaction par groupes cibles'); emit(M, c, T.analyse1);
    M.h(4, 'Evaluation de la satisfaction de tous les usagers.'); emit(M, c, T.analyse2);
    M.h(3, 'Considérations éthiques et déontologique'); emit(M, c, T.ethique);
    B.extraFigs(M, P, 'methodes');
    /* --- Résultats --- */
    M.h(1, 'Résultats');
    var nPat = c.n.usagers, nH = u && u.hospitalise ? u.hospitalise.n : 0, nA = u && u.ambulatoire ? u.ambulatoire.n : 0;
    M.p("Au terme de notre étude, " + b(c.total) + " usagers ont fait l'objet de l'étude dont " + b(nPat) + ' patients' + (split ? ' (' + nH + ' hospitalisés ou accompagnants et ' + nA + ' ambulatoires)' : '') + ', ' + b(c.n.personnel) + ' prestataires des services de santé et ' + b(c.n.communaute) + ' en communauté.');
    M.h(2, "Description de l'échantillon");
    if (u && u.n) {
      M.h(3, 'Population des patients enquêtés');
      var dA = u.demoAmb || { n: 0 }, dH = u.demoHosp || { n: 0 }, d = u.demo;
      var tS = sumVals(d.sexe), mas = cnt(d.sexe, 'Masculin'), fem = cnt(d.sexe, 'Féminin');
      var ta = d.age || {}, tA = sumVals(ta), a2550 = cnt(ta, 'Entre 25 et 50 ans');
      var tp = top(d.profession), ti = top(d.instruction);
      M.p('Le sexe ' + (mas >= fem ? 'masculin' : 'féminin') + ' domine cette population avec ' + b(pc(B.pctOf(Math.max(mas, fem), tS))) + ' contre ' + pc(B.pctOf(Math.min(mas, fem), tS)) + " chez l'autre sexe. " + (tA ? "L'âge moyen est de " + f1(d.ageMoy) + ' ans' + ' et ' + pc(B.pctOf(a2550, tA)) + ' des enquêtés sont âgés de 25 à 50 ans. ' : '') + (tp ? 'La profession la plus représentée est « ' + tp + ' » (' + pc(B.pctOf(d.profession[tp], sumVals(d.profession))) + ').' : '') + (ti ? ' Le niveau d’instruction le plus fréquent est « ' + ti + ' ».' : ''));
      profilTable(M, c, 'Répartition des patients selon leur profil sociodémographique, ' + HZ + ' en ' + Y, dA, dH, [{ title: 'Sexe', key: 'sexe', cats: ['Féminin', 'Masculin'] }, { title: 'Age', key: 'age', cats: ['Moins de 25 ans', 'Entre 25 et 50 ans', 'Plus de 50 ans'] }, { title: 'Provenance', key: 'provenance' }, { title: 'Profession', key: 'profession' }, { title: "Niveau d'instruction", key: 'instruction' }]);
    }
    HK.slice(1).forEach(function (k) {
      var t = ci[k]; if (!t || !t.n) return;
      M.h(3, k === 'personnel' ? 'Population du personnel enquêté' : 'Population des communautés');
      M.p(k === 'personnel' ? 'Au total, ' + b(t.n) + ' membres du personnel de l’hôpital ont été enquêtés.' : 'Au total, ' + b(t.n) + ' personnes de la communauté environnante ont été enquêtées.');
      var dm = t.demo || {};
      if (sumVals(dm.sexe)) simpleTable(M, 'Répartition ' + (k === 'personnel' ? 'du personnel' : 'des communautés') + ' enquêté(e)s selon le sexe, ' + HZ + ' en ' + Y, 'Sexe', dm.sexe);
      if (sumVals(dm.age)) simpleTable(M, 'Répartition ' + (k === 'personnel' ? 'du personnel' : 'des communautés') + " enquêté(e)s selon la tranche d'âge, " + HZ + ' en ' + Y, "Tranche d'âge", dm.age, ['Moins de 25 ans', 'Entre 25 et 50 ans', 'Plus de 50 ans']);
      if (sumVals(dm.profession)) simpleTable(M, 'Répartition ' + (k === 'personnel' ? 'du personnel' : 'des communautés') + ' enquêté(e)s selon la profession, ' + HZ + ' en ' + Y, 'Profession', dm.profession);
    });
    /* --- Évaluation de la satisfaction --- */
    M.h(2, 'Evaluation de la satisfaction des usagers en ' + Y);
    var sectName = { usagers: 'Mesure de la satisfaction des Patients', personnel: 'Mesure de la satisfaction chez le personnel', communaute: 'Mesure de la satisfaction des communautés' };
    var globName = { usagers: 'Evaluation globale de la satisfaction des patients', personnel: 'Evaluation globale de la satisfaction des personnels', communaute: "Mesure globale de la satisfaction des communautés environnants l'hôpital" };
    var CAPW = { usagers: 'des usagers', personnel: 'des personnels', communaute: 'des communautés' };
    var TGT = { usagers: 'chez les patients', personnel: 'chez le personnel', communaute: 'chez les communautés' };
    HK.forEach(function (k) {
      var t = ci[k]; if (!t || !t.n) return;
      M.h(3, sectName[k]);
      var sp = k === 'usagers' && split;
      t.cats.forEach(function (cn) {
        var items = (t.all.items[cn] || []), ih = sp ? (t.hospitalise.items[cn] || []) : [], ia = sp ? (t.ambulatoire.items[cn] || []) : [], sc = t.all.cats[cn];
        M.h(4, cn);
        M.p('La satisfaction sur « ' + cn + ' » ' + (k === 'usagers' ? 'à l’' + HZ + ' ' : '') + 'était de ' + b(pc(sc)) + (sp ? ' (' + pc(t.hospitalise.cats[cn]) + ' chez les hospitalisés contre ' + pc(t.ambulatoire.cats[cn]) + ' chez les ambulatoires)' : '') + '. Le tableau ci-dessous présente la répartition de cette satisfaction par niveau d’appréciation.');
        var ok = items.filter(function (x) { return isNum(x.s); });
        if (ok.length > 1) { var so = B.sortDesc(ok.map(function (x) { return { name: lab(x.c), score: x.s }; }), 'score'); M.p('Le point le mieux noté est « ' + so[0].name + ' » (' + b(pc(so[0].score)) + ') ; le moins bien noté est « ' + so[so.length - 1].name + ' » (' + b(pc(so[so.length - 1].score)) + ').'); }
        var na = function (v) { return isNum(v) ? f1(v) : 'NA'; };
        var rows = items.map(function (it, j) { var line = [lab(it.c)]; if (sp) line.push(na(ih[j] ? ih[j].s : null), na(ia[j] ? ia[j].s : null)); line.push(f1(it.s)); return line; });
        rows.push([cn].concat(sp ? [na(t.hospitalise.cats[cn]), na(t.ambulatoire.cats[cn])] : [], [f1(sc)]));
        M.table('Répartition ' + CAPW[k] + ' selon leur satisfaction sur « ' + cn + ' », ' + HZ + ' en ' + Y, [k === 'usagers' ? 'Client' : k === 'personnel' ? 'Personnel' : 'Communauté'].concat(sp ? ['Hospitalisés', 'Ambulatoires'] : [], ['Moyenne']), rows, { lastBold: true, note: sp ? 'NA : Données non prises en compte' : '' });
      });
      if (k === 'usagers') {
        var tx = t.texts || {}, nm = Object.keys(tx);
        ['I4', 'I5'].forEach(function (col) { if (tx[col] && tx[col].length) { M.p(b(col === 'I4' ? 'Ce que les patients n’ont pas aimé' : 'Ce que les patients recommandent pour améliorer les prestations')); M.bullets(tx[col].slice(0, 12)); } });
      }
      var sc = t.cats.map(function (n) { return t.all.cats[n]; }), avg = t.all.o, up = [], dn = [];
      t.cats.forEach(function (n, i) { if (!isNum(sc[i])) return; (sc[i] >= avg ? up : dn).push(n + ' (' + pc(sc[i]) + ')'); });
      M.h(3, globName[k]);
      M.p('Au total ' + nw(t.cats.length) + ' domaines de satisfaction ont été évalués ' + TGT[k] + '. A l’issue de cette évaluation, il se révèle un taux global de satisfaction de ' + b(pc(avg)) + ' ' + TGT[k].replace('chez les', 'chez les').replace('chez le ', 'chez le ') + ' enquêté(e)s avec ' + nw(dn.length) + ' domaine(s) non satisfait(s)' + (dn.length ? '. Il s’agit de : ' + dn.join(', ') + '.' : '.'));
      M.chart('Niveau global de satisfaction ' + CAPW[k] + ' par domaine, ' + HZ + ' en ' + Y, { kind: 'bar', labels: t.cats, values: sc, mean: avg, meanLabel: 'moyenne', name: 'Score (%)' });
    });
    /* --- global tous usagers --- */
    var cib = HK.filter(function (k) { return ci[k] && ci[k].n; }).map(function (k) { return { k: k, s: ci[k].all.o }; });
    M.h(3, 'Evaluation globale de la satisfaction des usagers (patients, communautés et personnels) en ' + Y);
    M.p('Le niveau global de satisfaction des usagers de l’' + HZ + ' en ' + Y + ' est de ' + b(pc(data.combined)) + ' (' + b(B.qual(data.combined)) + '), moyenne des taux de satisfaction des cibles enquêtées : ' + B.listFr(cib.map(function (x) { return HDE[x.k] + ' (' + b(pc(x.s)) + ')'; })) + '.');
    M.table('Satisfaction globale des usagers selon la cible, ' + HZ + ' en ' + Y, ['Cible', 'Effectif', 'Moyenne (%)'], cib.map(function (x) { return [HL[x.k], String(c.n[x.k]), f1(x.s)]; }).concat([['Satisfaction globale', String(c.total), f1(data.combined)]]), { lastBold: true });
    M.chart('Satisfaction globale des usagers de l’' + HZ + ' en ' + Y, { kind: 'bar', labels: cib.map(function (x) { return HL[x.k]; }).concat(['Satisfaction globale']), values: cib.map(function (x) { return x.s; }).concat([data.combined]), mean: data.combined, meanLabel: 'satisfaction globale' });
    /* --- évolution --- */
    var cols = [];
    if (cmp) cols.push({ label: cmp.labels.join(' + '), d: cmp });
    bp.forEach(function (p) { cols.push({ label: p.label, d: p }); });
    var lastLab = bp.length ? bp[bp.length - 1].label : Y;
    M.h(3, 'Évolution de la satisfaction des usagers du ' + HZ + (cols.length ? ' : progression entre ' + cols[0].label + ' et ' + lastLab : ''));
    if (!cols.length) M.p('Aucune période de référence n’a été sélectionnée pour ce rapport : sélectionnez une période de comparaison (ex. l’année précédente) ou plusieurs périodes pour obtenir l’évolution de la satisfaction.');
    else {
      var cur = bp.length ? 'Ensemble' : Y, dz = isNum(data.combined) && isNum(cols[0].d.combined) ? data.combined - cols[0].d.combined : null;
      if (isNum(dz)) M.p('Le niveau global de satisfaction passe de ' + b(pc(cols[0].d.combined)) + ' (' + cols[0].label + ') à ' + b(pc(data.combined)) + ' (' + lastLab + '), soit une ' + (dz >= 0 ? 'progression' : 'baisse') + ' de ' + b(f1(Math.abs(dz)) + ' points') + '.');
      var heads = cols.map(function (x) { return x.label + ' (%)'; });
      var rows = HK.filter(function (k) { return ci[k] && ci[k].n; }).map(function (k) { var line = [HL[k]]; cols.forEach(function (x) { line.push(f1(x.d.cibles[k] ? x.d.cibles[k].o : null)); }); line.push(f1(ci[k].all.o)); line.push(delta(ci[k].all.o, cols[0].d.cibles[k] ? cols[0].d.cibles[k].o : null)); return line; });
      rows.push(['Satisfaction globale'].concat(cols.map(function (x) { return f1(x.d.combined); }), [f1(data.combined), delta(data.combined, cols[0].d.combined)]));
      M.table('Evolution de la satisfaction par cible, ' + HZ, ['Cible'].concat(heads, [cur + ' (%)', 'Évolution (points)']), rows, { lastBold: true });
      HK.forEach(function (k) {
        var t = ci[k]; if (!t || !t.n || !cols[0].d.cibles[k] || !cols[0].d.cibles[k].cats) return;
        var r2 = t.cats.map(function (n) { var line = [n]; cols.forEach(function (x) { line.push(f1(x.d.cibles[k] && x.d.cibles[k].cats ? x.d.cibles[k].cats[n] : null)); }); line.push(f1(t.all.cats[n])); line.push(delta(t.all.cats[n], cols[0].d.cibles[k].cats[n])); return line; });
        M.table('Evolution de la satisfaction ' + CAPW[k] + ' par domaine, ' + HZ, ['Domaine'].concat(heads, [cur + ' (%)', 'Évolution (points)']), r2, {});
      });
      M.chart('Evolution du niveau de satisfaction global entre ' + cols[0].label + ' et ' + lastLab + ' à l’hôpital de zone de ' + c.nm + '.', { kind: 'cluster', labels: rows.map(function (r) { return r[0]; }), series: cols.map(function (x, ix) { return { name: x.label, values: rows.map(function (r) { var v = parseFloat(String(r[ix + 1]).replace(',', '.')); return isNum(v) ? v : null; }) }; }).concat([{ name: cur, values: rows.map(function (r) { var v = parseFloat(String(r[cols.length + 1]).replace(',', '.')); return isNum(v) ? v : null; }) }]), unit: ' %', max: 100 });
    }
    /* --- autres paramètres --- */
    M.h(2, 'Autres paramètres');
    if (u && u.n && u.autres) {
      var AA = u.autres, tot = u.n;
      M.h(3, 'Les pratiques parallèles');
      var prow = ['AA1', 'AA2', 'AA3'].map(function (n) { var o = AA[n] || {}, oui = cnt(o, 'Oui'), tt = sumVals(o); return [lab(n), String(oui), f1(B.pctOf(oui, tt)), String(tt)]; });
      var just = AA.AA4 || {}, jn = cnt(just, 'Non'), jt = sumVals(just);
      var any = prow.filter(function (r) { return Number(r[1]) > 0; });
      M.p('Les pratiques parallèles (propositions de vente de médicaments ou de soins contre de l’argent, prestations hors de l’hôpital, sollicitations de cadeaux) ont été signalées par ' + (any.length ? any.map(function (r) { return b(r[2] + ' %') + ' des répondants (« ' + r[0].replace(/^\d+\.\s*/, '').slice(0, 90) + ' »)'; }).join(' ; ') : 'aucun répondant') + '.' + (jt ? ' Un justificatif n’a pas été délivré après paiement pour ' + b(pc(B.pctOf(jn, jt))) + ' des répondants concernés.' : ''));
      M.table('Répartition des usagers victimes des pratiques parallèles, ' + HZ + ' en ' + Y, ['Pratique signalée', 'Oui (n)', '%', 'Répondants'], prow.concat(jt ? [[lab('AA4') + ' (non)', String(jn), f1(B.pctOf(jn, jt)), String(jt)]] : []), {});
      M.h(3, "Adoption de l'hôpital pour les soins");
      var i1 = AA.I1 || {}, i2 = AA.I2 || {}, r1 = cnt(i1, 'Oui'), r2n = cnt(i2, 'Oui');
      M.p('Parmi les usagers enquêtés, ' + b(pc(B.pctOf(r1, sumVals(i1)))) + ' reviendraient dans l’hôpital pour un éventuel problème de santé et ' + b(pc(B.pctOf(r2n, sumVals(i2)))) + ' le recommanderaient à d’autres personnes.');
      M.table("Adoption de l'hôpital par les usagers, " + HZ + ' en ' + Y, ['Question', 'Oui (n)', '%', 'Répondants'], [[lab('I1'), String(r1), f1(B.pctOf(r1, sumVals(i1))), String(sumVals(i1))], [lab('I2'), String(r2n), f1(B.pctOf(r2n, sumVals(i2))), String(sumVals(i2))]], {});
    }
    var RS = (u && u.raisons) || (ci.communaute && ci.communaute.raisons) || null;
    if (RS && sumVals(RS)) {
      M.h(3, "Raisons du choix de l'hôpital");
      var rk = Object.keys(RS).sort(function (x, z) { return RS[z] - RS[x]; }), base = (u && u.n) || 1;
      M.p('Les raisons les plus citées pour le choix de l’hôpital sont « ' + rk[0] + ' » (' + b(pc(B.pctOf(RS[rk[0]], base))) + ' des répondants)' + (rk[1] ? ', « ' + rk[1] + ' » (' + pc(B.pctOf(RS[rk[1]], base)) + ')' : '') + '.');
      M.table("Raisons du choix de l'hôpital, " + HZ + ' en ' + Y, ['Raison', 'Effectif', '% des répondants'], rk.map(function (k) { return [k, String(RS[k]), f1(B.pctOf(RS[k], base))]; }), {});
    }
    /* --- recommandations, conclusion, annexes --- */
    B.extraFigs(M, P, 'resultats');
    var recos = hospRecos(c);
    M.h(1, 'Recommandations'); M.bullets(recos);
    M.h(1, 'Conclusion', { noNum: true });
    if (P.conclusionTexte) String(P.conclusionTexte).split(/\n+/).forEach(function (x) { M.p(x); });
    else M.p("L'évaluation conduite à l'" + HZ + ' (' + c.periode + ') révèle un niveau de satisfaction global ' + b(B.qual(data.combined)) + ' (' + b(pc(data.combined)) + '), avec ' + B.listFr(cib.map(function (x) { return HDE[x.k] + ' (' + b(pc(x.s)) + ')'; })) + (cmp && isNum(cmp.combined) ? '. Par rapport à ' + cmp.labels.join(' + ') + ' (' + pc(cmp.combined) + '), le niveau de satisfaction est ' + (data.combined >= cmp.combined ? 'en progression' : 'en baisse') : '') + '. Les efforts doivent porter en priorité sur les domaines les moins bien notés.');
    M.h(1, 'ANNEXES', { noNum: true });
    B.extraFigs(M, P, 'annexes');
    M.p('Critères de validation : domaines de satisfaction et questions de l’enquête.');
    HK.forEach(function (k) {
      var t = ci[k]; if (!t || !t.cats) return;
      var rows = []; t.cats.forEach(function (cn) { (t.all.items[cn] || []).forEach(function (it, j) { rows.push([j === 0 ? cn : '', lab(it.c)]); }); });
      if (rows.length) M.table('Critères de validation des domaines de satisfaction ' + CAPW[k] + ', ' + HZ + ' en ' + Y, ['Domaine', 'Question'], rows, {});
    });
    var resume = hospResume(c, cib, cmp);
    return { model: M, ctx: c, meta: hospMeta(c, P, data), recos: recos, resume: resume, motsCles: 'Évaluation, satisfaction, hôpital de zone, ' + HZ };
  }
  function lowDomain(ci, k) { var t = ci[k]; if (!t || !t.cats) return null; var a = t.cats.map(function (n) { return { n: n, s: t.all.cats[n] }; }).filter(function (x) { return isNum(x.s); }).sort(function (x, z) { return x.s - z.s; }); return a[0] || null; }
  function hospResume(c, cib, cmp) {
    var data = c.data, ci = c.ci, nm = c.nm, HZ = c.HZ, srt = cib.slice().sort(function (a, z) { return z.s - a.s; }), low = srt[srt.length - 1], ld = low ? lowDomain(ci, low.k) : null;
    return [
      { label: 'Introduction', text: "La satisfaction en milieu de soins est définie comme la réaction de l'individu aux évènements marquants son contact avec les services de soins. Soucieux du niveau de satisfaction de ses usagers afin de l’améliorer, l’Hôpital de Zone de " + nm + " s'est inscrit dans une démarche d'évaluation de la satisfaction des usagers et des prestataires." },
      { label: 'Objectifs', text: "Déterminer le niveau de satisfaction des usagers (hospitalisés, ambulatoires et des communautés environnantes) sur les prestations de soins à l’" + HZ + ", déterminer les domaines d’insatisfaction des usagers, analyser les variations de satisfaction dans le temps et proposer des approches d’amélioration." },
      { label: 'Méthodes', text: "Étude transversale descriptive et comparative menée à l’" + HZ + ' (' + c.periode + ') auprès de ' + b(B.fN(c.total)) + ' personnes : ' + HK.map(function (k) { return B.fN(c.n[k]) + ' ' + ({ usagers: 'usagers', personnel: 'membres du personnel', communaute: 'personnes de la communauté' })[k]; }).join(', ') + '. Les scores sont calculés par domaine de satisfaction à partir des réponses cotées de 0 à 1.' },
      { label: 'Résultats', text: 'Le niveau global de satisfaction de l’' + HZ + ' est de ' + b(pc(data.combined)) + ' (' + b(B.qual(data.combined)) + ')' + (cmp ? ', ' + (data.combined >= cmp.combined ? 'en progression' : 'en baisse') + ' par rapport à ' + cmp.labels.join(' + ') + ' (' + pc(cmp.combined) + ')' : '') + '. ' + srt.map(function (x) { return 'Satisfaction ' + HDE[x.k] + ' : ' + b(pc(x.s)); }).join(' ; ') + '.' + (ld ? ' Le domaine le moins bien noté ' + HDE[low.k] + ' est « ' + ld.n + ' » (' + pc(ld.s) + ').' : '') },
      { label: 'Conclusion', text: 'L’évaluation révèle un niveau de satisfaction ' + b(B.qual(data.combined)) + ' à l’' + HZ + '. ' + (low ? 'La satisfaction ' + HDE[low.k] + ' est la plus faible et appelle des actions d’amélioration prioritaires' + (ld ? ', en particulier sur « ' + ld.n + ' »' : '') + '.' : '') }
    ];
  }
  function hospMeta(c, P, data) {
    var dd = data.generatedAt ? new Date(data.generatedAt) : new Date(), zoneTxt = c.zone === c.nm && !(P.nomZoneSanitaire || P.zone) ? '' : c.zone.toUpperCase();
    var mo = MOIS[dd.getMonth()]; mo = mo.charAt(0).toUpperCase() + mo.slice(1);
    return { kind: 'hopital', title: 'ÉVALUATION DE LA SATISFACTION DES CLIENTS — ' + c.NAME, sub: 'Édition ' + c.edition + ' — ' + c.periode, edition: c.edition, consultant: P.consultant || 'Ets MOUREDE & Associés', commanditaire: P.commanditaire || 'Administration de l’Hôpital de Zone de ' + c.nm, presentateur: P.presentateur || 'Équipe de Consultants', draft: P.statut !== 'final', generatedAt: data.generatedAt, structure: c.name, scope: 'hopital', logo: P.logo || '', fontSz: 26,
      cover: { hopital: true, ligne1: zoneTxt ? 'ZONE SANITAIRE ' + zoneTxt : '', ligne2: 'HÔPITAL DE ZONE DE ' + c.nm.toUpperCase(), titre: 'Enquête de satisfaction des clients de l’hôpital de zone' + (zoneTxt ? ' ' + zoneTxt.charAt(0) + zoneTxt.slice(1).toLowerCase() : ' de ' + c.nm), realisePar: 'Réalisée par l’' + (P.consultant || 'Ets MOUREDE & Associés'), commanditePar: 'Commanditée par : ' + (P.commanditaire || 'Administration de l’Hôpital de Zone de ' + c.nm), date: 'Réalisation : ' + (P.dateRapport || (mo + ' ' + dd.getFullYear())), adresse: ['Adresse 01 BP 882 - COTONOU', 'Tél. : +229 21 33 2178 / 21 33 21 63', 'infos@sante.gouv.bj  —  www.sante.gouv.bj'] } };
  }
  function hospRecos(c) {
    if (c.P.recos && c.P.recos.length) return c.P.recos;
    var out = [], ci = c.ci;
    HK.forEach(function (k) {
      var t = ci[k]; if (!t || !t.n) return;
      t.cats.map(function (n) { return { n: n, s: t.all.cats[n] }; }).filter(function (x) { return isNum(x.s); }).sort(function (a, z) { return a.s - z.s; }).slice(0, 2).forEach(function (x) { out.push('Améliorer le domaine « ' + x.n + ' » (' + pc(x.s) + ', ' + HDE[k] + '), en traitant en priorité les points les moins bien notés.'); });
    });
    out.push('Partager les résultats avec les services de l’hôpital pour l’élaboration d’un plan d’amélioration.');
    out.push('Reconduire l’évaluation régulièrement pour suivre l’évolution de la satisfaction.');
    return out;
  }
  root.RapportHosp = { build: buildHosp, ctx: ctx, recoList: function (data, P) { return hospRecos(ctx(data, P)); }, HK: HK, HL: HL };
})(typeof window !== 'undefined' ? window : globalThis);
/* ===== Graphiques (canvas), aperçu HTML et export Word ===== */
(function (root) {
  'use strict';
  var B = root.RapportBase, f1 = B.f1, isNum = B.isNum;
  var PAL = ['0F4C81', 'F59E0B', '2E9E5B', 'C0392B', '7C3AED', '0891B2'];

  /* numérotation des titres : 1, 1.1, 1.1.1 (les titres « noNum » ne sont pas numérotés) */
  function numberize(blocks) {
    var a = 0, b2 = 0, c2 = 0;
    blocks.forEach(function (x) {
      if (x.t === 'h1') { if (x.noNum) { x.num = ''; return; } a++; b2 = 0; c2 = 0; x.num = a + '.'; }
      else if (x.t === 'h2') { if (x.noNum) { x.num = ''; return; } b2++; c2 = 0; x.num = (a || 1) + '.' + b2 + '.'; }
      else if (x.t === 'h3') { if (x.noNum) { x.num = ''; return; } c2++; x.num = (a || 1) + '.' + (b2 || 1) + '.' + c2 + '.'; }
    });
  }
  function richRuns(text, o) {
    var out = [], parts = String(text).split('**');
    parts.forEach(function (s, i) { if (s) out.push({ text: s, bold: i % 2 === 1 }); });
    return out;
  }
  function wrapText(ctx, text, maxW, maxLines) {
    var words = String(text).split(/\s+/), lines = [], cur = '';
    words.forEach(function (w) { var t = cur ? cur + ' ' + w : w; if (ctx.measureText(t).width > maxW && cur) { lines.push(cur); cur = w; } else cur = t; });
    if (cur) lines.push(cur);
    if (lines.length > maxLines) { lines = lines.slice(0, maxLines); lines[maxLines - 1] = lines[maxLines - 1].replace(/.{0,3}$/, '…'); }
    return lines;
  }
  /** Dessine un graphique : bar (vertical/horizontal, couleur selon la moyenne) ou cluster (barres groupées). */
  function drawChart(spec, W) {
    W = W || 900;
    var n = spec.labels.length, horizontal = spec.kind === 'bar' && (spec.horizontal || n > 7), H;
    if (spec.kind === 'cluster') H = 440; else H = horizontal ? Math.max(260, 90 + n * 34) : 420;
    var cv = document.createElement('canvas'), S = 2; cv.width = W * S; cv.height = H * S;
    var g = cv.getContext('2d'); g.scale(S, S); g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
    g.font = '13px Calibri, Arial, sans-serif'; g.textBaseline = 'middle';
    var pct = spec.unit !== '' && !spec.integer, max = spec.max || (pct ? 100 : Math.max.apply(null, [1].concat(spec.kind === 'cluster' ? [].concat.apply([], spec.series.map(function (s) { return s.values.map(function (v) { return isNum(v) ? v : 0; }); })) : spec.values.map(function (v) { return isNum(v) ? v : 0; }))) * 1.15);
    if (!pct && !spec.max) max = Math.ceil(max);
    var fmt = function (v) { return isNum(v) ? (pct ? f1(v) : String(Math.round(v))) : '—'; };
    function grid(x0, x1, y0, y1, vert) {
      g.strokeStyle = '#e5e7eb'; g.fillStyle = '#6b7280'; g.lineWidth = 1; g.font = '11px Calibri, Arial, sans-serif';
      for (var i = 0; i <= 5; i++) {
        var v = max * i / 5;
        if (vert) { var y = y1 - (y1 - y0) * i / 5; g.beginPath(); g.moveTo(x0, y); g.lineTo(x1, y); g.stroke(); g.textAlign = 'right'; g.fillText(fmt(v), x0 - 6, y); }
        else { var x = x0 + (x1 - x0) * i / 5; g.beginPath(); g.moveTo(x, y0); g.lineTo(x, y1); g.stroke(); g.textAlign = 'center'; g.fillText(fmt(v), x, y1 + 12); }
      }
      g.font = '13px Calibri, Arial, sans-serif';
    }
    if (spec.kind === 'bar' && horizontal) {
      var lw = Math.min(300, W * 0.38), x0 = lw + 10, x1 = W - 60, y0 = 34, y1 = H - 44, bh = (y1 - y0) / n;
      grid(x0, x1, y0, y1, false);
      spec.labels.forEach(function (l, i) {
        var v = spec.values[i], yc = y0 + bh * i + bh / 2, w = isNum(v) ? (x1 - x0) * Math.min(v, max) / max : 0;
        g.fillStyle = '#1f2937'; g.textAlign = 'right'; var lines = wrapText(g, l, lw, 2);
        lines.forEach(function (t, k) { g.fillText(t, lw, yc + (k - (lines.length - 1) / 2) * 14); });
        g.fillStyle = '#' + B.colorFor(v, spec.mean); g.fillRect(x0, yc - bh * 0.32, w, bh * 0.64);
        g.fillStyle = '#1f2937'; g.textAlign = 'left'; g.fillText(isNum(v) ? fmt(v) + (pct ? ' %' : '') : '—', x0 + w + 6, yc);
      });
      if (isNum(spec.mean)) { var mx = x0 + (x1 - x0) * spec.mean / max; g.save(); g.setLineDash([6, 4]); g.strokeStyle = '#0f4c81'; g.lineWidth = 2; g.beginPath(); g.moveTo(mx, y0 - 8); g.lineTo(mx, y1); g.stroke(); g.restore(); g.fillStyle = '#0f4c81'; g.font = 'bold 12px Calibri, Arial, sans-serif'; g.textAlign = 'center'; g.fillText('Moyenne ' + f1(spec.mean) + ' %', Math.min(Math.max(mx, x0 + 50), x1 - 40), 14); }
    } else if (spec.kind === 'bar') {
      var x0 = 50, x1 = W - 20, y0 = 20, y1 = H - 70, bw = (x1 - x0) / n;
      grid(x0, x1, y0, y1, true);
      spec.labels.forEach(function (l, i) {
        var v = spec.values[i], h = isNum(v) ? (y1 - y0) * Math.min(v, max) / max : 0, xc = x0 + bw * i + bw / 2;
        g.fillStyle = '#' + B.colorFor(v, spec.mean); g.fillRect(xc - bw * 0.3, y1 - h, bw * 0.6, h);
        g.fillStyle = '#1f2937'; g.textAlign = 'center'; g.fillText(isNum(v) ? fmt(v) + (pct ? ' %' : '') : '—', xc, y1 - h - 10);
        wrapText(g, l, bw - 6, 3).forEach(function (t, k) { g.fillText(t, xc, y1 + 16 + k * 14); });
      });
      if (isNum(spec.mean)) { var my = y1 - (y1 - y0) * spec.mean / max; g.save(); g.setLineDash([6, 4]); g.strokeStyle = '#0f4c81'; g.lineWidth = 2; g.beginPath(); g.moveTo(x0, my); g.lineTo(x1, my); g.stroke(); g.restore(); g.fillStyle = '#0f4c81'; g.font = 'bold 12px Calibri, Arial, sans-serif'; g.textAlign = 'right'; g.fillText('Moyenne ' + f1(spec.mean) + ' %', x1, my - 9); }
    } else {
      var x0 = 50, x1 = W - 20, y0 = 40, y1 = H - 60, gw = (x1 - x0) / n, ns = spec.series.length, bw2 = gw * 0.8 / ns;
      grid(x0, x1, y0, y1, true);
      spec.labels.forEach(function (l, i) {
        var xs = x0 + gw * i + gw * 0.1;
        spec.series.forEach(function (s, k) {
          var v = s.values[i], h = isNum(v) ? (y1 - y0) * Math.min(v, max) / max : 0;
          g.fillStyle = '#' + PAL[k % PAL.length]; g.fillRect(xs + bw2 * k, y1 - h, bw2 - 2, h);
          if (isNum(v)) { g.fillStyle = '#1f2937'; g.textAlign = 'center'; g.font = '11px Calibri, Arial, sans-serif'; g.fillText(fmt(v), xs + bw2 * k + bw2 / 2 - 1, y1 - h - 8); g.font = '13px Calibri, Arial, sans-serif'; }
        });
        g.fillStyle = '#1f2937'; g.textAlign = 'center'; wrapText(g, l, gw - 6, 2).forEach(function (t, k) { g.fillText(t, x0 + gw * i + gw / 2, y1 + 16 + k * 14); });
      });
      var lx = x0; spec.series.forEach(function (s, k) { g.fillStyle = '#' + PAL[k % PAL.length]; g.fillRect(lx, 12, 14, 14); g.fillStyle = '#1f2937'; g.textAlign = 'left'; g.fillText(s.name, lx + 20, 19); lx += g.measureText(s.name).width + 50; });
    }
    if (spec.kind === 'bar' && isNum(spec.mean)) { g.fillStyle = '#' + B.GOOD; g.fillRect(W - 250, H - 22, 10, 10); g.fillStyle = '#6b7280'; g.font = '11px Calibri, Arial, sans-serif'; g.textAlign = 'left'; g.fillText('≥ moyenne', W - 236, H - 17); g.fillStyle = '#' + B.BAD; g.fillRect(W - 160, H - 22, 10, 10); g.fillStyle = '#6b7280'; g.fillText('< moyenne', W - 146, H - 17); }
    return cv;
  }
  function chartImage(spec, W) {
    var cv = drawChart(spec, W), url = cv.toDataURL('image/png'), bin = atob(url.split(',')[1]), u = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return { data: u, url: url, w: cv.width / 2, h: cv.height / 2 };
  }

  /* ---------- aperçu HTML ---------- */
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function rich(s) { return esc(s).replace(/\*\*(.+?)\*\*/g, '<b>$1</b>'); }
  function toHtml(res, withCharts) {
    var M = res.model, m = res.meta; numberize(M.blocks);
    var h = ['<div class="rp">'];
    h.push('<div class="cover"><div class="ct">' + esc(m.title) + '</div><div class="cs">' + esc(m.sub) + '</div><div class="cm">Réalisée par : ' + esc(m.consultant) + ' — Commanditée par : ' + esc(m.commanditaire) + '</div>' + (m.draft ? '<div class="dr">VERSION PROVISOIRE — rapport en cours de rédaction, données arrêtées au ' + esc(new Date(m.generatedAt || Date.now()).toLocaleString('fr-FR')) + '</div>' : '') + '</div>');
    M.blocks.forEach(function (x) {
      if (x.t === 'h1') h.push('<h2>' + esc((x.num ? x.num + ' ' : '') + x.text) + '</h2>');
      else if (x.t === 'h2') h.push('<h3>' + esc((x.num ? x.num + ' ' : '') + x.text) + '</h3>');
      else if (x.t === 'h3') h.push('<h4>' + esc((x.num ? x.num + ' ' : '') + x.text) + '</h4>');
      else if (x.t === 'h4') h.push('<h5>' + esc(x.text) + '</h5>');
      else if (x.t === 'p') x.text.split('\n').forEach(function (l) { h.push('<p>' + rich(l) + '</p>'); });
      else if (x.t === 'ul') h.push('<ul>' + x.items.map(function (i) { return '<li>' + rich(i) + '</li>'; }).join('') + '</ul>');
      else if (x.t === 'table') {
        h.push('<div class="cap">Tableau ' + B.roman(x.no) + ' : ' + esc(x.caption) + '</div><div class="tw"><table><thead><tr>' + x.head.map(function (c) { return '<th>' + esc(c) + '</th>'; }).join('') + '</tr></thead><tbody>' + x.rows.map(function (r, i) { return '<tr' + ((x.lastBold && i === x.rows.length - 1) || (x.boldRows || []).indexOf(i) >= 0 ? ' class="lb"' : '') + '>' + r.map(function (c, j) { return '<td' + (j ? ' class="n"' : '') + '>' + esc(c) + '</td>'; }).join('') + '</tr>'; }).join('') + '</tbody></table></div>');
      } else if (x.t === 'chart') {
        h.push('<div class="fg">' + (withCharts ? '<img src="' + chartImage(x.spec, 900).url + '" alt="">' : '') + '<div class="cap">Figure ' + x.no + ' : ' + esc(x.caption) + '</div></div>');
      } else if (x.t === 'image') {
        h.push('<div class="fg"><img src="' + esc(x.src) + '" alt=""><div class="cap">Figure ' + x.no + ' : ' + esc(x.caption) + '</div></div>');
      } else if (x.t === 'pb') h.push('<hr>');
    });
    h.push('</div>');
    return h.join('');
  }

  function toDocx(res) { return root.RapportWord.toDocx(res); }
  root.RapportDoc = { drawChart: drawChart, chartImage: chartImage, toHtml: toHtml, toDocx: toDocx, numberize: numberize, esc: esc };
})(typeof window !== 'undefined' ? window : globalThis);
/* ===== Export Word au canevas 2026 : Arial Narrow 12, tableaux bleus, graphiques Word natifs modifiables, couverture à logos ===== */
(function (root) {
  'use strict';
  var B = root.RapportBase, f1 = B.f1, isNum = B.isNum;
  var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
  var BASE = 24;
  var RF = '<w:rFonts w:ascii="Arial Narrow" w:hAnsi="Arial Narrow" w:cs="Arial Narrow"/>';
  function X(s) { return String(s === null || s === undefined ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function rpr(o) {
    o = o || {};
    return '<w:rPr>' + RF + (o.b ? '<w:b/><w:bCs/>' : '') + (o.i ? '<w:i/><w:iCs/>' : '') + '<w:color w:val="' + (o.color || '000000') + '"/>' + (o.sp ? '<w:spacing w:val="' + o.sp + '"/>' : '') + ((o.sz || (BASE !== 24 ? BASE : 0)) ? '<w:sz w:val="' + (o.sz || BASE) + '"/><w:szCs w:val="' + (o.sz || BASE) + '"/>' : '') + (o.hl ? '<w:shd w:val="clear" w:color="auto" w:fill="' + o.hl + '"/>' : '') + '</w:rPr>';
  }
  function run(t, o) { return '<w:r>' + rpr(o) + '<w:t xml:space="preserve">' + X(t) + '</w:t></w:r>'; }
  function runs(text, o) {
    o = o || {}; var out = '';
    String(text).split('**').forEach(function (s, i) { if (s) out += run(s, { b: o.b || i % 2 === 1, i: o.i, color: o.color, sz: o.sz, sp: o.sp, hl: o.hl }); });
    return out;
  }
  function ppr(o) {
    o = o || {};
    var s = '<w:pPr>';
    if (o.style) s += '<w:pStyle w:val="' + o.style + '"/>';
    if (o.keepNext) s += '<w:keepNext/>';
    if (o.keepLines) s += '<w:keepLines/>';
    if (o.pb) s += '<w:pageBreakBefore/>';
    if (o.numId) s += '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="' + o.numId + '"/></w:numPr>';
    if (o.border) s += '<w:pBdr>' + ['top', 'left', 'bottom', 'right'].map(function (k) { return '<w:' + k + ' w:val="single" w:sz="' + (o.border.sz || 12) + '" w:space="' + (o.border.space || 6) + '" w:color="' + o.border.color + '"/>'; }).join('') + '</w:pBdr>';
    if (o.fill) s += '<w:shd w:val="clear" w:color="auto" w:fill="' + o.fill + '"/>';
    s += '<w:spacing' + (o.before !== undefined ? ' w:before="' + o.before + '"' : '') + ' w:after="' + (o.after === undefined ? 0 : o.after) + '" w:line="' + (o.line || 360) + '" w:lineRule="auto"/>';
    if (o.ind) s += '<w:ind' + (o.ind.left !== undefined ? ' w:left="' + o.ind.left + '"' : '') + (o.ind.right !== undefined ? ' w:right="' + o.ind.right + '"' : '') + (o.ind.hanging ? ' w:hanging="' + o.ind.hanging + '"' : '') + '/>';
    s += '<w:jc w:val="' + (o.jc || 'both') + '"/>';
    if (o.outline !== undefined) s += '<w:outlineLvl w:val="' + o.outline + '"/>';
    s += '<w:rPr>' + RF + '<w:color w:val="000000"/></w:rPr>';
    if (o.sect) s += o.sect;
    return s + '</w:pPr>';
  }
  function para(text, o) { o = o || {}; return '<w:p>' + ppr(o) + runs(text, o) + '</w:p>'; }
  function fld(instr, cached, o) {
    return '<w:r>' + rpr(o) + '<w:fldChar w:fldCharType="begin"/></w:r><w:r>' + rpr(o) + '<w:instrText xml:space="preserve"> ' + instr + ' </w:instrText></w:r><w:r>' + rpr(o) + '<w:fldChar w:fldCharType="separate"/></w:r>' + run(cached, o) + '<w:r>' + rpr(o) + '<w:fldChar w:fldCharType="end"/></w:r>';
  }
  function pbreak() { return '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'; }

  /* ---------- tableaux ---------- */
  function tbl(head, rows, o) {
    o = o || {}; var nc = head.length, TW = 8700;
    var first = nc <= 2 ? 5400 : nc === 3 ? 4300 : nc <= 5 ? 2800 : nc <= 7 ? 2300 : 1900, other = Math.floor((TW - first) / Math.max(1, nc - 1)), widths = [first].concat(new Array(nc - 1).fill(other));
    var sz = nc >= 10 ? 18 : nc >= 7 ? 20 : nc >= 5 ? 22 : BASE, long = rows.length > 25;
    function cell(txt, j, fill, bold, keep) {
      return '<w:tc><w:tcPr><w:tcW w:w="' + widths[j] + '" w:type="dxa"/>' + (fill ? '<w:shd w:val="clear" w:color="auto" w:fill="' + fill + '"/>' : '') + '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar><w:vAlign w:val="center"/></w:tcPr>' +
        '<w:p>' + ppr({ keepNext: keep, jc: j === 0 ? 'left' : 'center', line: 276 }) + run(String(txt), { b: bold, sz: sz }) + '</w:p></w:tc>';
    }
    var bd = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (k) { return '<w:' + k + ' w:val="single" w:sz="4" w:space="0" w:color="auto"/>'; }).join('');
    var s = '<w:tbl><w:tblPr><w:tblW w:w="' + TW + '" w:type="dxa"/><w:tblBorders>' + bd + '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>' + widths.map(function (w) { return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
    s += '<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>' + head.map(function (h, j) { return cell(h, j, '2E75B6', true, true); }).join('') + '</w:tr>';
    rows.forEach(function (r, i) {
      var last = o.lastBold && i === rows.length - 1, isLast = i === rows.length - 1, grp = (o.boldRows || []).indexOf(i) >= 0;
      s += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + r.map(function (c, j) { return cell(c, j, last ? 'D9E2F3' : (grp ? 'F2F2F2' : (j === 0 ? 'F2F2F2' : null)), last || grp, !long && !isLast); }).join('') + '</w:tr>';
    });
    return s + '</w:tbl>' + '<w:p><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr></w:p>';
  }

  /* ---------- graphiques Word natifs (données dans le graphique, mise en forme modifiable) ---------- */
  function strLit(a) { return '<c:strLit><c:ptCount val="' + a.length + '"/>' + a.map(function (v, i) { return '<c:pt idx="' + i + '"><c:v>' + X(v) + '</c:v></c:pt>'; }).join('') + '</c:strLit>'; }
  function numLit(a) { return '<c:numLit><c:formatCode>0.0</c:formatCode><c:ptCount val="' + a.length + '"/>' + a.map(function (v, i) { return isNum(v) ? '<c:pt idx="' + i + '"><c:v>' + v + '</c:v></c:pt>' : ''; }).join('') + '</c:numLit>'; }
  function dLbls(int) { return '<c:dLbls><c:numFmt formatCode="' + (int ? '0' : '0.0') + '" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial Narrow"/></a:defRPr></a:pPr><a:endParaRPr lang="fr-FR"/></a:p></c:txPr><c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'; }
  function barSer(i, name, color, cats, vals, int) { return '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/><c:tx><c:v>' + X(name) + '</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>' + dLbls(int) + '<c:cat>' + strLit(cats) + '</c:cat><c:val>' + numLit(vals) + '</c:val></c:ser>'; }
  function chartXml(spec) {
    var cats = spec.labels.map(String), n = cats.length, bars = '', line = '', all = [], pct = !spec.integer;
    if (spec.kind === 'cluster') {
      var pal = ['2E75B6', 'F59E0B', '00B050', 'C00000', '7C3AED', '0891B2'];
      spec.series.forEach(function (s, k) { bars += barSer(k, s.name, pal[k % pal.length], cats, s.values, spec.integer); s.values.forEach(function (v) { if (isNum(v)) all.push(v); }); });
    } else {
      var m = spec.mean, vals = spec.values;
      vals.forEach(function (v) { if (isNum(v)) all.push(v); });
      if (isNum(m)) {
        bars += barSer(0, 'Domaine non satisfait', 'FF0000', cats, vals.map(function (v) { return isNum(v) && v < m ? v : null; }));
        bars += barSer(1, 'Domaine satisfait', '00B050', cats, vals.map(function (v) { return isNum(v) && v >= m ? v : null; }));
        line = '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="2"/><c:order val="2"/><c:tx><c:v>' + X(spec.meanLabel || 'moyenne') + '</c:v></c:tx><c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker><c:cat>' + strLit(cats) + '</c:cat><c:val>' + numLit(cats.map(function () { return Math.round(m * 10) / 10; })) + '</c:val><c:smooth val="0"/></c:ser><c:marker val="1"/><c:axId val="50010"/><c:axId val="50020"/></c:lineChart>';
      } else bars += barSer(0, spec.name || 'Score (%)', '2E75B6', cats, vals);
    }
    var mx = pct ? 100 : null, mn = null;
    if (pct && all.length) { var lo = Math.min.apply(null, all); if (lo >= 60) mn = Math.floor((lo - 12) / 10) * 10; else mn = 0; }
    if (!pct) { mn = 0; }
    var rot = n > 4 ? '-2700000' : '0';
    var cat = '<c:catAx><c:axId val="50010"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="low"/><c:txPr><a:bodyPr rot="' + rot + '" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:latin typeface="Arial Narrow"/></a:defRPr></a:pPr><a:endParaRPr lang="fr-FR"/></a:p></c:txPr><c:crossAx val="50020"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>';
    var val = '<c:valAx><c:axId val="50020"/><c:scaling><c:orientation val="minMax"/>' + (mx !== null ? '<c:max val="' + mx + '"/>' : '') + (mn !== null ? '<c:min val="' + mn + '"/>' : '') + '</c:scaling><c:delete val="1"/><c:axPos val="l"/><c:numFmt formatCode="0" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="fr-FR"/><c:roundedCorners val="0"/><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>' + bars + '<c:gapWidth val="' + (spec.kind === 'cluster' ? 80 : 50) + '"/>' + (spec.kind === 'cluster' ? '' : '<c:overlap val="100"/>') + '<c:axId val="50010"/><c:axId val="50020"/></c:barChart>' + line + cat + val + '</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:latin typeface="Arial Narrow"/></a:defRPr></a:pPr><a:endParaRPr lang="fr-FR"/></a:p></c:txPr></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr><a:latin typeface="Arial Narrow"/></a:defRPr></a:pPr><a:endParaRPr lang="fr-FR"/></a:p></c:txPr></c:chartSpace>';
  }
  function chartPara(rid, id, h) {
    return '<w:p><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="5760000" cy="' + h + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="' + id + '" name="Graphique ' + id + '"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="' + rid + '"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  }

  /* ---------- images ancrées (couverture) ---------- */
  function anchorImg(id, rid, cx, cy, x, y, behind, z) {
    return '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="' + z + '" behindDoc="' + (behind ? 1 : 0) + '" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>' + x + '</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>' + y + '</wp:posOffset></wp:positionV><wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="' + id + '" name="Image ' + id + '"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="' + id + '" name="Image ' + id + '"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';
  }
  function imgSize(u8) {   // dimensions PNG / JPEG
    var d = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (u8[0] === 0x89 && u8[1] === 0x50) return { w: d.getUint32(16), h: d.getUint32(20), ext: 'png' };
    var i = 2; while (i < u8.length) { if (u8[i] !== 0xFF) { i++; continue; } var m = u8[i + 1]; if (m >= 0xC0 && m <= 0xC3) return { h: d.getUint16(i + 5), w: d.getUint16(i + 7), ext: 'jpeg' }; i += 2 + d.getUint16(i + 2); }
    return { w: 1, h: 1, ext: 'png' };
  }
  function fitBox(sz, bw, bh) { var r = Math.min(bw / sz.w, bh / sz.h); return { cx: Math.round(sz.w * r), cy: Math.round(sz.h * r) }; }
  function dataUrlToU8(u) { var m = /^data:image\/(png|jpe?g|gif);base64,(.*)$/i.exec(u || ''); if (!m) return null; var bin = atob(m[2]), a = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i); return a; }

  var ASSETS = null;
  function loadAssets(base) {
    if (ASSETS) return Promise.resolve(ASSETS);
    base = base || 'modele/';
    function get(n) { return fetch(base + n).then(function (r) { if (!r.ok) throw new Error('Ressource introuvable : ' + n); return r.arrayBuffer(); }); }
    return Promise.all([get('modele-2026.docx'), get('ministere.png'), get('bandeau.png'), get('fond.jpg')]).then(function (a) { ASSETS = { modele: a[0], ministere: new Uint8Array(a[1]), bandeau: new Uint8Array(a[2]), fond: new Uint8Array(a[3]) }; return ASSETS; });
  }

  function summaryParas(res) { return (res.resume || []).map(function (s, i) { return para('**' + s.label + ' :** ' + s.text, { sz: 22, line: 300, after: 160, jc: 'both' }); }).join(''); }

  /** Assemble le .docx : parties de style du rapport 2026 + contenu généré. */
  function toDocx(res) {
    var Z = root.JSZip; if (!Z) return Promise.reject(new Error('Bibliothèque JSZip non chargée.'));
    var M = res.model, m = res.meta, blocks = M.blocks; root.RapportDoc.numberize(blocks); BASE = m.fontSz || 24;
    return loadAssets().then(function (A) { return Z.loadAsync(A.modele).then(function (zip) {
      var rels = [], ct = [], nextId = 20, chartN = 0, docId = 100, media = {};
      function addRel(type, target) { var id = 'rId' + (nextId++); rels.push('<Relationship Id="' + id + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + type + '" Target="' + target + '"/>'); return id; }
      function addMedia(name, u8) { zip.file('word/media/' + name, u8); return addRel('image', 'media/' + name); }
      /* ----- couverture ----- */
      var PW = 7560000, PH = 10692000, rFond = addMedia('fond.jpg', A.fond), rMin = addMedia('ministere.png', A.ministere), rBan = addMedia('bandeau.png', A.bandeau), logo = dataUrlToU8(m.logo), rLogo = null, lsz = null;
      var szMin = imgSize(A.ministere), bMin = fitBox(szMin, 2394342, 800000), szBan = imgSize(A.bandeau);
      if (logo) { lsz = imgSize(logo); rLogo = addMedia('logo-centre.' + lsz.ext, logo); }
      var anchors = anchorImg(1, rFond, PW, PH, 0, 0, true, 251650000) + anchorImg(2, rBan, 2615565, Math.round(2615565 * szBan.h / szBan.w), Math.round((PW - 2615565) / 2), 560000, false, 251651000) + anchorImg(3, rMin, bMin.cx, bMin.cy, 545465, 820000, false, 251652000) + anchorImg(4, rBan, 2615565, Math.round(2615565 * szBan.h / szBan.w), Math.round((PW - 2615565) / 2), 10050000, false, 251653000);
      if (rLogo) { var lb = fitBox(lsz, 900000, 900000); anchors += anchorImg(5, rLogo, lb.cx, lb.cy, PW - 1417 * 635 - lb.cx - 200000, 780000, false, 251654000); }
      var cv = m.cover, cover;
      if (cv.hopital) {
        cover = '<w:p>' + ppr({ line: 240 }) + anchors + '</w:p>' +
          para(cv.ligne1, { b: true, sz: 28, sp: 40, jc: 'center', before: 2200, after: 200, line: 276 }) +
          para(cv.ligne2, { b: true, sz: 48, color: '1F4E79', jc: 'center', after: 400, line: 276 }) +
          para(cv.titre, { b: true, sz: 30, jc: 'center', before: 400, after: 0, line: 360, border: { color: 'C00000', sz: 12, space: 10 }, fill: 'D9E2F3', ind: { left: 500, right: 500 } }) +
          para(cv.realisePar, { sz: 26, jc: 'center', before: 1400, after: 120, line: 276 }) +
          para(cv.commanditePar, { sz: 26, jc: 'center', after: 120, line: 276 }) +
          para(cv.date, { b: true, sz: 24, jc: 'right', before: 1200, after: 0 }) +
          cv.adresse.map(function (t, i) { return para(t, { sz: 18, jc: 'center', before: i === 0 ? 900 : 0, after: 0, line: 240, color: '44546A' }); }).join('');
      } else cover = '<w:p>' + ppr({ line: 240 }) + anchors + '</w:p>' +
        para(m.cover.direction, { b: true, sz: 26, sp: 40, jc: 'center', before: 1700, after: 120, line: 276 }) +
        para(m.cover.ligne2, { b: true, sz: 26, sp: 40, jc: 'center', after: 120, line: 276 }) +
        para(m.cover.titre, { b: true, sz: 28, jc: 'center', before: 1500, after: 0, line: 360, border: { color: 'C00000', sz: 12, space: 10 }, fill: 'D9E2F3', ind: { left: 500, right: 500 } }) +
        para(m.cover.date, { b: true, sz: 22, jc: 'right', before: 4200, after: 0 });
      cover = cover;
      var pgb = '<w:pgBorders w:display="firstPage" w:offsetFrom="page">' + ['top', 'left', 'bottom', 'right'].map(function (k) { return '<w:' + k + ' w:val="twistedLines2" w:sz="18" w:space="24" w:color="2F5496"/>'; }).join('') + '</w:pgBorders>';
      var sect1 = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1135" w:right="1841" w:bottom="1135" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>' + pgb + '<w:cols w:space="708"/></w:sectPr>';
      cover += '<w:p>' + ppr({ sect: sect1, line: 240 }) + '</w:p>';
      /* ----- résumé + sommaire + listes ----- */
      var ftRid = addRel('footer', 'footer1.xml');
      var resBox = '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>' + ['top', 'left', 'bottom', 'right'].map(function (k) { return '<w:' + k + ' w:val="single" w:sz="18" w:space="0" w:color="C00000"/>'; }).join('') + '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="200" w:type="dxa"/><w:left w:w="300" w:type="dxa"/><w:bottom w:w="200" w:type="dxa"/><w:right w:w="300" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>' +
        para('RESUME', { b: true, sz: 24, jc: 'center', border: { color: 'C00000', sz: 12, space: 4 }, fill: 'D9E2F3', ind: { left: 3000, right: 3000 }, after: 300, line: 276 }) + summaryParas(res) +
        para('Mots clés : ' + (res.motsCles || ''), { b: true, sz: 22, jc: 'center', hl: 'FAA394', fill: 'FAA394', line: 276, before: 120 }) + '</w:tc></w:tr></w:tbl>';
      function tocEntries(filter, styleOf) {
        var items = blocks.filter(filter), out = '';
        items.forEach(function (x, i) {
          var txt = styleOf.text(x), first = i === 0, last = i === items.length - 1;
          out += '<w:p><w:pPr><w:pStyle w:val="' + styleOf.style(x) + '"/><w:spacing w:after="60"/></w:pPr>' + (first ? '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ' + styleOf.instr + ' </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' : '') + run(txt, { sz: 22 }) + (last ? '<w:r><w:fldChar w:fldCharType="end"/></w:r>' : '') + '</w:p>';
        });
        return out || '<w:p/>';
      }
      var heads = blocks.filter(function (x) { return /^h[123]$/.test(x.t); });
      var toc = para('Sommaire', { b: true, sz: 28, jc: 'center', after: 200, outline: 0, pb: true }) + tocEntries(function (x) { return /^h[12]$/.test(x.t); }, { instr: 'TOC \\o "1-2" \\h \\z \\u', style: function (x) { return 'TM' + x.t.charAt(1); }, text: function (x) { return (x.num ? x.num + ' ' : '') + x.text; } });
      var lot = para('Liste des tableaux', { b: true, sz: 28, jc: 'left', after: 200, outline: 0, pb: true }) + tocEntries(function (x) { return x.t === 'table'; }, { instr: 'TOC \\h \\z \\c "Tableau"', style: function () { return 'Tabledesillustrations'; }, text: function (x) { return 'Tableau ' + B.roman(x.no) + ' : ' + x.caption; } });
      var lof = para('Liste des figures', { b: true, sz: 28, jc: 'left', after: 200, outline: 0, pb: true }) + tocEntries(function (x) { return x.t === 'chart' || x.t === 'image'; }, { instr: 'TOC \\h \\z \\c "Figure"', style: function () { return 'Tabledesillustrations'; }, text: function (x) { return 'Figure ' + x.no + ' : ' + x.caption; } });
      var sect2 = '<w:sectPr><w:footerReference w:type="default" r:id="' + ftRid + '"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1560" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:fmt="lowerRoman" w:start="1"/><w:cols w:space="708"/></w:sectPr>';
      var front = resBox + toc + lot + lof + '<w:p>' + ppr({ sect: sect2, line: 240 }) + '</w:p>';
      /* ----- corps ----- */
      var body = '', FIRSTBODY = true;
      blocks.forEach(function (x, bi) {
        if (x.t === 'h1') {
          var big = !x.noNum || /^(Introduction|Conclusion|Recommandations|Annexes)$/.test(x.text);
          if (x.noNum) body += para(x.text, { b: true, pb: !FIRSTBODY && big, outline: 0, keepNext: true, after: 120 }); else body += '<w:p>' + ppr({ style: 'Titre1', pb: !FIRSTBODY, before: 0, after: 120, outline: undefined }).replace('<w:rPr>' + RF + '<w:color w:val="000000"/></w:rPr>', '<w:rPr>' + RF + '<w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>') + run((x.num ? x.num + ' ' : '') + x.text, { b: true, sz: 24 }) + '</w:p>';
          FIRSTBODY = false;
        } else if (x.t === 'h2') body += '<w:p>' + ppr({ style: 'Titre2', keepNext: true, before: 160, after: 80 }) + run((x.num ? x.num + ' ' : '') + x.text, { b: true, sz: 24 }) + '</w:p>';
        else if (x.t === 'h3') body += '<w:p>' + ppr({ style: 'Titre3', keepNext: true, before: 160, after: 80 }) + run((x.num ? x.num + ' ' : '') + x.text, { b: true, sz: 24 }) + '</w:p>';
        else if (x.t === 'h4') body += para(x.text, { b: true, keepNext: true, before: 80 });
        else if (x.t === 'p') x.text.split('\n').forEach(function (l) { body += para(l); });
        else if (x.t === 'ul') x.items.forEach(function (it) { body += para(it, { numId: 3 }); });
        else if (x.t === 'table') {
          body += '<w:p>' + ppr({ keepNext: true, keepLines: true, before: 120, after: 60 }) + run('Tableau ', { b: true }) + fld('SEQ Tableau \\* ROMAN', B.roman(x.no), { b: true }) + run(' : ' + x.caption, { b: true }) + '</w:p>' + tbl(x.head, x.rows, { lastBold: x.lastBold, boldRows: x.boldRows }) + (x.note ? para(x.note, { sz: 18, i: true, line: 276, after: 120 }) : '');
        } else if (x.t === 'chart') {
          chartN++; var rid = addRel('chart', 'charts/chart' + chartN + '.xml');
          zip.file('word/charts/chart' + chartN + '.xml', chartXml(Object.assign({ meanLabel: x.spec.meanLabel || (/commune/i.test(x.caption) ? 'moyenne communale' : /zone/i.test(x.caption) ? 'moyenne de la zone' : 'moyenne') }, x.spec)));
          ct.push('<Override PartName="/word/charts/chart' + chartN + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>');
          body += chartPara(rid, 200 + chartN, x.spec.labels.length > 8 ? 3700000 : 3200000) + '<w:p>' + ppr({ keepLines: true, before: 60, after: 160 }) + run('Figure ', { b: true }) + fld('SEQ Figure \\* ARABIC', String(x.no), { b: true }) + run(' : ' + x.caption, { b: true }) + '</w:p>';
        } else if (x.t === 'image') {
          var iu = dataUrlToU8(x.src);
          if (iu) {
            var isz = imgSize(iu), irid = addMedia('figure' + x.no + '.' + isz.ext, iu), iw = 5400000, ih = Math.round(iw * isz.h / isz.w); if (ih > 5200000) { ih = 5200000; iw = Math.round(ih * isz.w / isz.h); }
            body += '<w:p><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="0"/><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + iw + '" cy="' + ih + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="' + (300 + x.no) + '" name="Figure ' + x.no + '"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="' + (300 + x.no) + '" name="Figure ' + x.no + '"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="' + irid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + iw + '" cy="' + ih + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' + '<w:p>' + ppr({ keepLines: true, before: 60, after: 160 }) + run('Figure ', { b: true }) + fld('SEQ Figure \\* ARABIC', String(x.no), { b: true }) + run(' : ' + x.caption, { b: true }) + '</w:p>';
          }
        } else if (x.t === 'pb') body += pbreak();
      });
      var sect3 = '<w:sectPr><w:footerReference w:type="default" r:id="' + ftRid + '"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/><w:pgNumType w:start="1"/><w:cols w:space="708"/></w:sectPr>';
      var doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + NS + '><w:body>' + cover + front + body + sect3 + '</w:body></w:document>';
      zip.file('word/document.xml', doc);
      /* ----- relations, types de contenu, réglages ----- */
      var fixed = [['styles', 'styles.xml', 'rId1'], ['numbering', 'numbering.xml', 'rId2'], ['settings', 'settings.xml', 'rId3'], ['fontTable', 'fontTable.xml', 'rId4'], ['webSettings', 'webSettings.xml', 'rId5'], ['footnotes', 'footnotes.xml', 'rId6'], ['endnotes', 'endnotes.xml', 'rId7'], ['theme', 'theme/theme1.xml', 'rId8']];
      var fx = fixed.map(function (f) { return '<Relationship Id="' + f[2] + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + f[0] + '" Target="' + f[1] + '"/>'; }).join('');
      zip.file('word/_rels/document.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + fx + rels.join('') + '</Relationships>');
      var W = 'application/vnd.openxmlformats-officedocument.wordprocessingml.';
      zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Override PartName="/word/document.xml" ContentType="' + W + 'document.main+xml"/><Override PartName="/word/styles.xml" ContentType="' + W + 'styles+xml"/><Override PartName="/word/numbering.xml" ContentType="' + W + 'numbering+xml"/><Override PartName="/word/settings.xml" ContentType="' + W + 'settings+xml"/><Override PartName="/word/fontTable.xml" ContentType="' + W + 'fontTable+xml"/><Override PartName="/word/webSettings.xml" ContentType="' + W + 'webSettings+xml"/><Override PartName="/word/footnotes.xml" ContentType="' + W + 'footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="' + W + 'endnotes+xml"/><Override PartName="/word/footer1.xml" ContentType="' + W + 'footer+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + ct.join('') + '</Types>');
      zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>');
      zip.file('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + X(m.title) + '</dc:title><dc:creator>' + X(m.consultant) + '</dc:creator></cp:coreProperties>');
      zip.remove('README.txt');
      return zip.file('word/settings.xml').async('string').then(function (s) {
        if (!/updateFields/.test(s)) s = s.replace(/(<w:settings[^>]*>)/, '$1<w:updateFields w:val="true"/>');
        zip.file('word/settings.xml', s);
        return zip.generateAsync({ type: 'blob', mimeType: W + 'document', compression: 'DEFLATE' });
      });
    }); });
  }
  root.RapportWord = { toDocx: toDocx, loadAssets: loadAssets, chartXml: chartXml };
})(typeof window !== 'undefined' ? window : globalThis);
/* ===== Export PowerPoint au style de la présentation ZoBoZa 2026 (pptxgenjs) : graphiques natifs modifiables ===== */
(function (root) {
  'use strict';
  var B = root.RapportBase, f1 = B.f1, isNum = B.isNum;
  var AR = 'Arial', IMP = 'Impact', RED = 'C00000', ORG = 'ED7D31', NAVY = '002060', GRN = '1E8E3E', BAD = 'C0392B';
  function plain(s) { return String(s).replace(/\*\*/g, ''); }
  function shortT(s, n) { s = plain(s); return s.length > n ? s.slice(0, n - 1) + '…' : s; }
  function noAcc(s) { return String(s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase(); }
  function u8(url) { var m = /^data:image\/(png|jpe?g|gif);base64,(.*)$/i.exec(url || ''); return m ? url : null; }
  var CAB = null;
  function loadCabinet(base) {
    if (CAB !== null) return Promise.resolve(CAB);
    return fetch((base || 'modele/') + 'cabinet.png').then(function (r) { return r.ok ? r.arrayBuffer() : null; }).then(function (b) {
      if (!b) return (CAB = '');
      var u = new Uint8Array(b), s = ''; for (var i = 0; i < u.length; i += 8192) s += String.fromCharCode.apply(null, u.subarray(i, i + 8192));
      return (CAB = 'data:image/png;base64,' + btoa(s));
    }).catch(function () { return (CAB = ''); });
  }
  function fsTitle(cap) {
    var mm = /Score de satisfaction (.+?) pour le domaine « (.+?) », par centre de santé, (commune .+?|hôpital .+?)(?:, (?:20\d\d))?$/.exec(cap);
    if (!mm) return cap.replace(/^Score de satisfaction /, '').replace(/, (20\d\d)$/, '');
    var c = /usagers/.test(mm[1]) ? 'Usagers' : /personnel/.test(mm[1]) ? 'Personnel soignant' : /COGECS/.test(mm[1]) ? 'Membres COGECS' : mm[1];
    return c + ' — Domaine « ' + mm[2] + ' » — ' + mm[3].charAt(0).toUpperCase() + mm[3].slice(1);
  }
  function toPptx(res) {
    return loadCabinet().then(function (cab) { return build(res, cab); });
  }
  function build(res, cab) {
    var P = new root.PptxGenJS(), m = res.meta, M = res.model, W = 13.33;
    P.layout = 'LAYOUT_WIDE'; P.title = m.title; P.author = m.consultant; P.company = String(m.consultant).replace(/&/g, '&amp;');
    root.RapportDoc.numberize(M.blocks);
    var section = '';
    function frame(titleBig, boxText, boxW) {
      var s = P.addSlide(); s.background = { color: 'FFFFFF' };
      s.addShape(P.ShapeType.line, { x: 0, y: 0.1, w: W, h: 0, line: { color: 'F8CBAD', width: 1.5 } });
      s.addShape(P.ShapeType.line, { x: 0, y: 7.38, w: W, h: 0, line: { color: 'F8CBAD', width: 1.5 } });
      if (titleBig) s.addText(noAcc(titleBig), { x: 0.5, y: 0.2, w: W - 1, h: 0.75, fontFace: IMP, fontSize: 36, color: '0D1B2A', align: 'center', valign: 'middle', fit: 'shrink' });
      if (boxText) s.addText(plain(boxText), { x: (W - (boxW || 10.5)) / 2, y: 1.05, w: boxW || 10.5, h: 0.62, fontFace: AR, fontSize: boxText.length > 95 ? 13 : 16, bold: true, color: '000000', align: 'center', valign: 'middle', line: { color: m.draft ? RED : ORG, width: 1.5 }, fit: 'shrink' });
      s.addShape(P.ShapeType.rect, { x: W - 0.9, y: 6.95, w: 0.5, h: 0.35, fill: { color: 'FFFFFF' }, line: { color: ORG, width: 1 } });
      s.slideNumber = { x: W - 0.9, y: 6.95, w: 0.5, h: 0.35, fontFace: AR, fontSize: 11, color: '404040', align: 'center' };
      if (m.draft) s.addText('PROVISOIRE', { x: 0.35, y: 6.95, w: 1.6, h: 0.32, fontFace: AR, fontSize: 10, bold: true, color: RED, align: 'center', valign: 'middle', line: { color: RED, width: 1 } });
      return s;
    }
    function divider(text) { var s = frame(null, null); s.addText(plain(text), { x: 2.2, y: 2.9, w: W - 4.4, h: 1.2, fontFace: AR, fontSize: 28, bold: true, color: '000000', align: 'center', valign: 'middle', line: { color: ORG, width: 2 }, fit: 'shrink' }); }
    function bullets(title, items, sz, box) {
      for (var i = 0; i < items.length; i += 6) {
        var s = frame(title, box || null, 9);
        s.addText(items.slice(i, i + 6).map(function (t) { return { text: shortT(t, 380), options: { bullet: true, breakLine: true } }; }), { x: 0.9, y: box ? 1.9 : 1.3, w: W - 1.8, h: box ? 4.9 : 5.5, fontFace: AR, fontSize: sz || 17, color: '1F2937', valign: 'top', paraSpaceAfter: 10, fit: 'shrink' });
      }
    }
    function textSlide(title, items) {
      var s = frame(title, null);
      s.addText(items.map(function (t) { return { text: shortT(t, 700), options: { breakLine: true } }; }), { x: 0.9, y: 1.3, w: W - 1.8, h: 5.5, fontFace: AR, fontSize: 16, color: '1F2937', valign: 'top', paraSpaceAfter: 12, lineSpacingMultiple: 1.3, align: 'justify', fit: 'shrink' });
    }
    function chartSlide(title, spec, box) {
      var s = frame(title, box || title);
      var labels = spec.labels.slice(), values = spec.values ? spec.values.slice() : null;
      if (spec.kind === 'bar') {
        var idx = labels.map(function (l, i) { return i; }); if (!spec.keepOrder) idx.sort(function (a, z) { return (isNum(values[z]) ? values[z] : -1) - (isNum(values[a]) ? values[a] : -1); });
        labels = idx.map(function (i) { return spec.labels[i]; }); values = idx.map(function (i) { return isNum(spec.values[i]) ? spec.values[i] : 0; });
        var mean = spec.mean, n = labels.length, lo = Math.min.apply(null, values.filter(function (v) { return v > 0; }).concat([100])), min = (!spec.integer && lo >= 60) ? Math.floor((lo - 12) / 10) * 10 : 0;
        var common = { x: 0.6, y: 1.8, w: W - 1.2, h: 4.95, valAxisHidden: true, valGridLine: { style: 'none' }, catAxisLabelFontSize: n > 8 ? 10 : 12, catAxisLabelFontFace: AR, catAxisLabelRotate: n > 5 ? 315 : 0, showLegend: false, valAxisMinVal: min, valAxisMaxVal: spec.integer ? undefined : 100, dataLabelFontSize: 11, dataLabelFormatCode: spec.integer ? '0' : '0.0', showValue: true, barGapWidthPct: 35 };
        var barData = [{ name: 'Score (%)', labels: labels, values: values }], colors = values.map(function (v) { return B.colorFor(v, mean) === B.GOOD ? GRN : isNum(mean) ? BAD : '2E75B6'; });
        if (isNum(mean)) {
          try {
            s.addChart([{ type: P.charts.BAR, data: barData, options: { barDir: 'col', chartColors: colors, barGapWidthPct: 35 } }, { type: P.charts.LINE, data: [{ name: 'moyenne', labels: labels, values: labels.map(function () { return Math.round(mean * 10) / 10; }) }], options: { chartColors: ['000000'], lineSize: 2, lineDataSymbol: 'none', lineDash: 'dash', showValue: false } }], common);
          } catch (e) { s.addChart(P.charts.BAR, barData, Object.assign({ barDir: 'col', chartColors: colors }, common)); }
          s.addText('Moyenne : ' + f1(mean) + ' %', { x: 4.2, y: 6.78, w: 5, h: 0.4, fontFace: AR, fontSize: 14, bold: true, color: '000000', align: 'center' });
        } else s.addChart(P.charts.BAR, barData, Object.assign({ barDir: 'col', chartColors: colors }, common));
      } else {
        var cols = ['2E75B6', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
        s.addChart(P.charts.BAR, spec.series.map(function (x) { return { name: x.name, labels: spec.labels, values: x.values.map(function (v) { return isNum(v) ? v : 0; }) }; }), { x: 0.6, y: 1.8, w: W - 1.2, h: 5.0, barDir: 'col', chartColors: cols.slice(0, spec.series.length), showLegend: true, legendPos: 'b', legendFontSize: 12, valAxisHidden: true, valGridLine: { style: 'none' }, catAxisLabelFontSize: 12, showValue: true, dataLabelFontSize: 11, dataLabelFormatCode: spec.integer ? '0' : '0.0', valAxisMinVal: 0, valAxisMaxVal: spec.integer ? undefined : 100, barGapWidthPct: 60 });
      }
    }
    function tableSlide(title, head, rows, lastBold) {
      var per = 10, nc = head.length, chunks = Math.max(1, Math.ceil(rows.length / per));
      for (var k = 0; k < chunks; k++) {
        var part = rows.slice(k * per, (k + 1) * per), s = frame(section ? section : null, title + (chunks > 1 ? ' (' + (k + 1) + '/' + chunks + ')' : ''), 11.5);
        var fW = nc <= 3 ? 6 : Math.min(3.6, (W - 1.6) * 0.28), colW = [fW].concat(new Array(nc - 1).fill((W - 1.6 - fW) / Math.max(1, nc - 1)));
        var data = [head.map(function (c) { return { text: String(c), options: { bold: true, color: '000000', fill: { color: '2E75B6' }, align: 'center', valign: 'middle' } }; })];
        part.forEach(function (r, i) { var last = lastBold && k === chunks - 1 && i === part.length - 1; data.push(r.map(function (c, j) { return { text: String(c), options: { bold: last, align: j ? 'center' : 'left', valign: 'middle', fill: { color: last ? 'D9E2F3' : 'FFFFFF' } } }; })); });
        s.addTable(data, { x: 0.8, y: 1.95, w: W - 1.6, colW: colW, fontFace: AR, fontSize: nc >= 7 ? 11 : 13, border: { type: 'solid', pt: 0.75, color: '000000' }, rowH: 0.4, autoPage: false });
      }
    }
    /* ----- présentation hôpital : plan des présentations HZ 2026 (Kandi, Malanville, Pobè, Bassila) ----- */
    function hospDeck() {
      var c = res.ctx, ci = c.ci, data = c.data, P = c.P, T = root.HZT || {}, HZ = c.HZ, Y = c.edition, HKs = ['usagers', 'personnel', 'communaute'];
      var LAB = root.HLABELS || {}, labH = function (k) { return String(LAB[k] || k).replace(/\s+/g, ' ').replace(/\?$/, ' ?'); };
      var split = ci.usagers && ci.usagers.hospitalise && ci.usagers.hospitalise.n && ci.usagers.ambulatoire && ci.usagers.ambulatoire.n;
      var PLN = ['INTRODUCTION', 'CADRE CONCEPTUEL', 'OBJECTIFS', 'METHODES', 'RESULTATS', 'SUGGESTION/RECOMMANDATIONS', 'CONCLUSION'], pl = frame('Plan', null), c6 = ['ED7D31', '70AD47', 'FFC000', '7030A0', '2E75B6', 'C00000', '00B0F0'];
      PLN.forEach(function (t, i) { pl.addText(t, { x: 3.7, y: 1.2 + i * 0.78, w: 5.9, h: 0.55, fontFace: AR, fontSize: 16, bold: true, align: 'center', valign: 'middle', color: '000000', fill: { color: 'FFFFFF' }, line: { color: c6[i], width: 2 }, rectRadius: 0.2 }); });
      var sub = function (t) { return c.sub(t); };
      divider('Introduction & Cadre conceptuel');
      textSlide('Introduction', [sub(T.intro1 || '').split('. ').slice(0, 4).join('. ') + '.', "L'Hôpital de zone de " + c.nm + ' a initié, pour le compte de l’année ' + Y + ', une enquête de satisfaction des clients afin de prouver scientifiquement les faiblesses et les domaines de soins à améliorer pour le bien-être des populations.']);
      textSlide('Cadre conceptuel', (T.cadre || []).filter(function (x) { return x.t === 'p'; }).slice(0, 4).map(function (x) { return sub(x.text); }));
      divider('Objectifs');
      bullets('Objectifs', ['**Objectif général :** Evaluer la satisfaction des usagers à l’' + HZ + ' en ' + Y + '.', 'Déterminer le niveau de satisfaction des usagers (hospitalisés, ambulatoires et des communautés environnantes) sur les prestations de soins à l’' + HZ + ' ;', 'Déterminer les domaines d’insatisfaction des usagers ;', 'Analyser les variations de satisfaction dans le temps ;', 'Proposer des approches d’amélioration de la satisfaction.'], 16);
      divider('Méthodes');
      var sz = frame('Méthodes', 'Présentation de l’' + HZ, 8);
      sz.addText([P.superficie ? 'Superficie : ' + P.superficie : '', P.population ? 'Population estimée : ' + P.population + ' habitants' + (B.density(P).dens ? ' (densité : ' + B.fN(B.density(P).dens) + ' hab./km²)' : '') : '', P.populationCible ? 'Population cible : ' + P.populationCible : '', P.nbFsPub ? P.nbFsPub + ' formations sanitaires publiques' + (P.nbFsPriv ? ' et ' + P.nbFsPriv + ' privées' : '') : '', 'Hôpital de première référence pour la zone sanitaire ' + c.zone].filter(Boolean).map(function (t) { return { text: t, options: { bullet: true, breakLine: true } }; }), { x: 0.8, y: 1.9, w: 5.8, h: 4.6, fontFace: AR, fontSize: 16, color: '1F2937', valign: 'top', paraSpaceAfter: 10 });
      var cm = u8(P.carte); if (cm) { try { sz.addImage({ data: cm, x: 7.0, y: 1.9, w: 5.6, h: 4.7, sizing: { type: 'contain', w: 5.6, h: 4.7 } }); } catch (er) {} }
      HKs.forEach(function (k) { if (ci[k]) bullets('Critères d’appréciation (' + ({ usagers: 'Usagers', personnel: 'Personnel', communaute: 'Communautés' })[k] + ')', ci[k].cats.map(function (n, i) { return (i + 1) + '. ' + n; }), 16); });
      bullets('Population source de l’enquête', ((T.population || []).filter(function (x) { return x.t === 'ul'; })[0] || { items: [] }).items.map(sub), 16);
      bullets('Echantillon et techniques', [c.total + ' enquêtés : ' + ci.usagers.n + ' patients' + (split ? ' (' + ci.usagers.hospitalise.n + ' hospitalisés contre ' + ci.usagers.ambulatoire.n + ' ambulatoires)' : '') + ', ' + c.n.communaute + ' communautés et ' + c.n.personnel + ' personnels.', 'Echantillonnage par commodité jusqu’à saturation des données.', 'Collecte par questionnaires numériques (3 questionnaires : patients, communautés, personnel).'], 16);
      bullets('Considérations éthiques', ['Consentement éclairé verbal des enquêtés.', 'Assurance de l’adhésion volontaire, de l’anonymat et de la confidentialité.', 'Autorisation de l’autorité administrative obtenue.'], 18);
      var rs = frame('Résultats', null); ['1. Description de la population', '2. Satisfaction des Patients', '3. Satisfaction des personnels', '4. Satisfaction des communautés', '5. Satisfaction globale et évolution'].forEach(function (t, i) { rs.addText(t, { x: 3.2, y: 1.5 + i * 0.9, w: 6.9, h: 0.6, fontFace: AR, fontSize: 18, bold: true, align: 'center', valign: 'middle', line: { color: c6[i], width: 2 }, rectRadius: 0.2 }); });
      var cnt = { 'Ambulatoire': split ? ci.usagers.ambulatoire.n : 0, 'Hospitalisés': split ? ci.usagers.hospitalise.n : ci.usagers.n, 'Communautés': c.n.communaute, 'Personnels': c.n.personnel };
      chartSlide('Résultats', { kind: 'bar', labels: Object.keys(cnt), values: Object.keys(cnt).map(function (k) { return cnt[k]; }), integer: true, keepOrder: true }, 'Description des enquêtés');
      var dm = ci.usagers.demo || {}, sx = dm.sexe || {}, tsx = Object.keys(sx).reduce(function (a, k) { return a + sx[k]; }, 0), ag = dm.age || {}, tag = Object.keys(ag).reduce(function (a, k) { return a + ag[k]; }, 0);
      var rowsP = Object.keys(sx).map(function (k) { return ['Sexe : ' + k, String(sx[k]), f1(B.pctOf(sx[k], tsx))]; }).concat(Object.keys(ag).map(function (k) { return ['Âge : ' + k, String(ag[k]), f1(B.pctOf(ag[k], tag))]; }));
      if (rowsP.length) tableSlide('Description des Patients enquêtés', ['Caractéristique', 'Effectif', '%'], rowsP, false);
      var CIBN = { usagers: 'Patients', personnel: 'Personnels', communaute: 'Communautés' };
      function domSlides(k) {
        var t = ci[k]; if (!t || !t.n) return; var sp = k === 'usagers' && split;
        divider(CIBN[k]);
        t.cats.forEach(function (cn, di) {
          var items = t.all.items[cn] || [], ih = sp ? t.hospitalise.items[cn] || [] : [], ia = sp ? t.ambulatoire.items[cn] || [] : [], na = function (v) { return isNum(v) ? f1(v) : 'NA'; };
          var rows = items.map(function (it, j) { var l = [shortT(labH(it.c), 95)]; if (sp) l.push(na(ih[j] ? ih[j].s : null), na(ia[j] ? ia[j].s : null)); l.push(f1(it.s)); return l; });
          rows.push([cn].concat(sp ? [na(t.hospitalise.cats[cn]), na(t.ambulatoire.cats[cn])] : [], [f1(t.all.cats[cn])]));
          var head = [k === 'usagers' ? 'Client' : CIBN[k]].concat(sp ? ['Hospitalisés', 'Ambulatoires'] : [], ['Moyenne']), per = 9, ch = Math.ceil(rows.length / per);
          for (var q = 0; q < ch; q++) {
            var part = rows.slice(q * per, (q + 1) * per), s = frame('Résultats', (di + 1) + '. ' + cn + (ch > 1 ? ' (' + (q + 1) + '/' + ch + ')' : ''), 9);
            var colW = sp ? [7.3, 1.7, 1.7, 1.4] : [10.6, 1.6], data2 = [head.map(function (h) { return { text: h, options: { bold: true, color: '000000', fill: { color: '2E75B6' }, align: 'center', valign: 'middle' } }; })];
            part.forEach(function (r, i) { var last = q === ch - 1 && i === part.length - 1; data2.push(r.map(function (x, j) { return { text: String(x), options: { bold: last, align: j ? 'center' : 'left', valign: 'middle', fill: { color: last ? 'D9E2F3' : 'FFFFFF' } } }; })); });
            s.addTable(data2, { x: 0.6, y: 1.95, w: colW.reduce(function (a, z) { return a + z; }, 0), colW: colW, fontFace: AR, fontSize: 11, border: { type: 'solid', pt: 0.75, color: '000000' }, rowH: 0.42, autoPage: false });
          }
        });
        divider('Evaluation de la satisfaction globale ' + ({ usagers: 'des patients', personnel: 'des personnels', communaute: 'des communautés' })[k]);
        var sc = t.cats.map(function (n) { return t.all.cats[n]; });
        chartSlide('Résultats', { kind: 'bar', labels: t.cats, values: sc, mean: t.all.o, keepOrder: true }, 'Evaluation globale de la satisfaction chez ' + ({ usagers: 'les patients', personnel: 'les personnels', communaute: 'les communautés' })[k] + ' : ' + f1(t.all.o) + ' %');
        if (k === 'usagers' && ci.usagers.autres) {
          var AA = ci.usagers.autres, pr = ['AA1', 'AA2', 'AA3'].map(function (n) { var o = AA[n] || {}, oui = o.Oui || 0, tt = Object.keys(o).reduce(function (a, z) { return a + o[z]; }, 0); return [shortT(labH(n).replace(/^\d+\.\s*/, ''), 110), String(oui), f1(B.pctOf(oui, tt))]; });
          tableSlide('Pratiques parallèles', ['Pratique signalée', 'Oui (n)', '%'], pr, false);
        }
      }
      domSlides('usagers'); domSlides('communaute'); domSlides('personnel');
      var cib = HKs.filter(function (k) { return ci[k] && ci[k].n; });
      divider('Evaluation de la satisfaction globale des usagers');
      chartSlide('Résultats', { kind: 'bar', labels: cib.map(function (k) { return ({ usagers: 'Patients', personnel: 'Personnels', communaute: 'Communautés' })[k]; }).concat(['Satisfaction globale']), values: cib.map(function (k) { return ci[k].all.o; }).concat([data.combined]), mean: data.combined, keepOrder: true }, 'Evaluation globale de la satisfaction chez les usagers : ' + f1(data.combined) + ' %');
      var cols = []; if (data.compare && isNum(data.compare.combined)) cols.push({ label: data.compare.labels.join(' + '), d: data.compare }); (data.byPeriod || []).forEach(function (p) { cols.push({ label: p.label, d: p }); });
      if (cols.length) {
        var last = cols[cols.length - 1], cur = (data.byPeriod || []).length ? 'Ensemble' : Y;
        tableSlide('Comparaison ' + cols[0].label + ' Vs ' + ((data.byPeriod || []).length ? c.periode : Y), ['Période', 'Niveau de satisfaction (%)', 'Évolution (points)'], cols.map(function (x) { return [x.label, f1(x.d.combined), '']; }).concat([[cur, f1(data.combined), B.delta(data.combined, cols[0].d.combined)]]), true);
      }
      if (ci.usagers.autres) { var A2 = ci.usagers.autres, i1 = A2.I1 || {}, i2 = A2.I2 || {}, t1 = Object.keys(i1).reduce(function (a, z) { return a + i1[z]; }, 0), t2 = Object.keys(i2).reduce(function (a, z) { return a + i2[z]; }, 0); bullets('Adoption / Fidélisation', ['Parmi les usagers enquêtés, ' + f1(B.pctOf(i1.Oui || 0, t1)) + ' % reviendraient dans l’hôpital pour un éventuel problème de santé.', f1(B.pctOf(i2.Oui || 0, t2)) + ' % recommanderaient l’hôpital à d’autres personnes.'].concat(ci.usagers.raisons ? ['Raisons du choix les plus citées : ' + Object.keys(ci.usagers.raisons).sort(function (x, z) { return ci.usagers.raisons[z] - ci.usagers.raisons[x]; }).slice(0, 3).join(', ') + '.'] : []), 18); }
      var tx = ci.usagers.texts || {};
      if (tx.I4 && tx.I4.length) bullets('Ce qui a marqué plus négativement les Patients', tx.I4.slice(0, 6), 16);
      if (tx.I5 && tx.I5.length) bullets('Ce que suggèrent les Patients', tx.I5.slice(0, 6), 16);
      var rec = frame('Recommandations', 'Administration de l’hôpital de zone de ' + c.nm, 9);
      rec.addText([{ text: res.recos.length + ' axes de recommandations pour améliorer la satisfaction des usagers', options: { bold: true, align: 'center', breakLine: true, fontSize: 18 } }].concat(res.recos.slice(0, 5).map(function (t) { return { text: shortT(t, 200), options: { bullet: true, breakLine: true, fontSize: 16 } }; })), { x: 0.9, y: 1.9, w: W - 1.8, h: 4.8, fontFace: AR, color: '1F2937', valign: 'top', paraSpaceAfter: 10, fill: { color: 'F2F2F2' }, line: { color: ORG, width: 1 }, fit: 'shrink' });
      var low = (function () { var best = null; cib.forEach(function (k) { ci[k].cats.forEach(function (n) { var s = ci[k].all.cats[n]; if (isNum(s) && (!best || s < best.s)) best = { n: n, s: s }; }); }); return best; })();
      var cc = frame('Conclusion', null); cc.addText([{ text: 'Niveau global de satisfaction des usagers en ' + Y + ' : ' + f1(data.combined) + ' % (' + B.qual(data.combined) + ')', options: { bold: true, breakLine: true, fontSize: 22 } }].concat(cib.map(function (k) { return { text: ({ usagers: 'Patients', personnel: 'Personnels', communaute: 'Communautés' })[k] + ' : ' + f1(ci[k].all.o) + ' %', options: { bullet: true, breakLine: true, fontSize: 18 } }; })).concat(low ? [{ text: 'Domaine le moins bien noté : ' + low.n + ' (' + f1(low.s) + ' %)', options: { bullet: true, breakLine: true, fontSize: 18 } }] : []), { x: 1.2, y: 1.5, w: W - 2.4, h: 5, fontFace: AR, color: '1F2937', valign: 'top', paraSpaceAfter: 12, line: { color: ORG, width: 1.5 }, fit: 'shrink' });
      divider('Merci pour votre attention');
    }
    /* ----- couverture ----- */
    var cv = P.addSlide(); cv.background = { color: 'FFFFFF' };
    if (cab) cv.addImage({ data: cab, x: 0.3, y: 0.25, w: 4.4, h: 0.59 });
    var lg = u8(m.logo); if (lg) cv.addImage({ data: lg, x: W - 1.7, y: 0.2, w: 1.2, h: 1.2, sizing: { type: 'contain', w: 1.2, h: 1.2 } });
    cv.addText(noAcc(m.scope === 'hopital' ? 'Evaluation de la satisfaction des usagers de l’hôpital de zone de ' + (res.ctx && res.ctx.nm ? res.ctx.nm : m.structure) : 'Evaluation de la satisfaction des usagers de la zone sanitaire ' + m.structure), { x: 1.4, y: 1.6, w: W - 2.8, h: 1.8, fontFace: AR, fontSize: 30, bold: true, color: NAVY, align: 'center', valign: 'middle', fill: { color: 'DEEAF6' }, line: { color: NAVY, width: 2 }, fit: 'shrink' });
    cv.addText('Edition ' + m.edition + (m.scope === 'hopital' ? '' : (m.sub && /—/.test(m.sub) ? '  —  ' + m.sub.split('—')[1].trim() : '')), { x: 2, y: 3.75, w: W - 4, h: 0.5, fontFace: AR, fontSize: 20, color: '1F2937', align: 'center' });
    cv.addText([{ text: 'Réalisée par : ', options: { bold: true, underline: true } }, { text: m.consultant }], { x: 2, y: 4.35, w: W - 4, h: 0.45, fontFace: AR, fontSize: 16, align: 'center' });
    cv.addText([{ text: 'Présentateur :', options: { bold: true, underline: true, breakLine: true } }, { text: m.presentateur }], { x: 1.6, y: 5.0, w: 4.5, h: 0.9, fontFace: AR, fontSize: 15, align: 'center' });
    cv.addText([{ text: 'Commanditée par :', options: { bold: true, underline: true, breakLine: true } }, { text: m.commanditaire }], { x: 7.2, y: 5.0, w: 4.8, h: 0.9, fontFace: AR, fontSize: 15, align: 'center' });
    cv.addText(((m.cover && m.cover.date) || '').replace(/^Réalisation : /, ''), { x: W - 3.2, y: 6.9, w: 2.8, h: 0.35, fontFace: AR, fontSize: 12, align: 'right', color: '404040' });
    if (m.draft) cv.addText('VERSION PROVISOIRE — données arrêtées au ' + new Date(m.generatedAt || Date.now()).toLocaleDateString('fr-FR'), { x: 3.4, y: 6.2, w: 6.5, h: 0.45, fontFace: AR, fontSize: 14, bold: true, color: RED, align: 'center', valign: 'middle', line: { color: RED, width: 1.5 } });
    if (m.scope === 'hopital') { hospDeck(); return P.write({ outputType: 'blob' }); }
    /* ----- plan ----- */
    var plan = M.blocks.filter(function (x) { return x.t === 'h1' && !/^(Annexes|Justification|Cadre conceptuel|Objectifs)$/.test(x.text); }).map(function (x) { return noAcc(x.text.replace(/^Résultats.*$/, 'Résultats').replace(/^Méthodologie$/, 'Méthodes')); });
    var pl = frame('Plan', null), cols6 = ['ED7D31', '70AD47', 'FFC000', '7030A0', '2E75B6', 'C00000', '00B0F0'];
    plan.forEach(function (t, i) { pl.addText(t, { x: 3.7, y: 1.2 + i * 0.78, w: 5.9, h: 0.55, fontFace: AR, fontSize: 16, bold: true, align: 'center', valign: 'middle', color: '000000', fill: { color: 'FFFFFF' }, line: { color: cols6[i % cols6.length], width: 2 }, rectRadius: 0.2 }); });
    /* ----- corps ----- */
    var blocks = M.blocks, i, TXT = { 'Introduction': 1, 'Justification': 1, 'Cadre conceptuel': 1, 'Objectifs': 1, 'Conclusion': 1, 'Recommandations': 1, "Cadre d'étude": 1 };
    for (i = 0; i < blocks.length; i++) {
      var x = blocks[i];
      if (x.t === 'h1') {
        section = noAcc(x.text.replace(/^Résultats.*$/, 'Résultats').replace(/^Méthodologie$/, 'Méthodes'));
        if (x.text === 'Recommandations') { bullets('Suggestions / recommandations', res.recos, 16); continue; }
        if (x.text === 'Annexes') break;
        if (TXT[x.text]) {
          var items = [], j = i + 1;
          while (j < blocks.length && blocks[j].t !== 'h1') { var y = blocks[j]; if (y.t === 'ul') items = items.concat(y.items); else if (y.t === 'p' && items.length < 10) items.push(y.text.split('\n')[0]); j++; }
          if (/^(Introduction|Justification|Cadre conceptuel|Conclusion)$/.test(x.text)) textSlide(x.text, items.slice(0, 4)); else bullets(x.text, items.slice(0, 12), 16);
          i = j - 1; continue;
        }
        if (x.text === 'Méthodologie') {
          divider('Méthodes');
          var k = i + 1, its = [];
          while (k < blocks.length && blocks[k].t !== 'h1') { var yy = blocks[k]; if (yy.t === 'h2' && !/^(Variables|Calendrier|Considérations)/.test(yy.text)) { its.push({ h: yy.text, t: [] }); } else if (its.length && (yy.t === 'p' || yy.t === 'ul')) { (yy.t === 'ul' ? yy.items : [yy.text.split('\n')[0]]).forEach(function (z) { if (its[its.length - 1].t.length < 3) its[its.length - 1].t.push(z); }); } else if (yy.t === 'h2' && /^Variables/.test(yy.text)) { its.push({ h: 'Variables de l’évaluation', t: [] }); } k++; }
          for (var q = 0; q < its.length; q += 2) { var s2 = frame('Méthodes', null); var chunk = its.slice(q, q + 2); s2.addText(chunk.map(function (c) { return [{ text: c.h, options: { bold: true, color: RED, breakLine: true, fontSize: 17 } }].concat(c.t.map(function (t) { return { text: shortT(t, 330), options: { bullet: true, breakLine: true, fontSize: 14 } }; })); }).reduce(function (a, z) { return a.concat(z); }, []), { x: 0.9, y: 1.3, w: W - 1.8, h: 5.5, fontFace: AR, color: '1F2937', valign: 'top', paraSpaceAfter: 6, fit: 'shrink' }); }
          i = k - 1; continue;
        }
        divider((x.num ? x.num + ' ' : '') + x.text);
      } else if (x.t === 'h2' && /^(Présentation|Synthèse|Comparaison|Évolution|Description|Satisfaction|Niveau)/.test(x.text)) divider(x.text);
      else if (x.t === 'image') { var si = frame(section || 'Présentation', x.caption, 10.5); try { si.addImage({ data: x.src, x: 2.2, y: 1.85, w: W - 4.4, h: 4.9, sizing: { type: 'contain', w: W - 4.4, h: 4.9 } }); } catch (er) {} }
      else if (x.t === 'chart') chartSlide(section || 'Résultats', x.spec, x.caption.replace(/, (20\d\d)$/, ''));
      else if (x.t === 'table') {
        if (x.head[0] === 'Formation sanitaire' && x.head.length === 3) {
          var rows = x.rows.slice(0, -1), avg = parseFloat(String(x.rows[x.rows.length - 1][2]).replace(',', '.'));
          chartSlide(section || 'Résultats', { kind: 'bar', labels: rows.map(function (r) { return r[0]; }), values: rows.map(function (r) { var v = parseFloat(String(r[2]).replace(',', '.')); return isNum(v) ? v : null; }), mean: avg }, fsTitle(x.caption));
        } else if (x.head.length <= 8 && x.rows.length <= 40 && !/profession|commune de résidence/.test(x.caption) && x.head[0] !== 'Question') tableSlide(x.caption, x.head, x.rows, x.lastBold);
      }
    }
    divider('Merci pour votre attention');
    return P.write({ outputType: 'blob' });
  }
  root.RapportPpt = { toPptx: toPptx, loadCabinet: loadCabinet };
})(typeof window !== 'undefined' ? window : globalThis);
