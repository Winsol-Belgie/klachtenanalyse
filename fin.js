/* Klachtenanalyse – rapport "Financieel": klachten en NTR-kost tegenover verkoop en omzet.
 *
 * Teller  = klachten in de selectie (alle filters gelden). Drie manieren van tellen:
 *           orders  → verschillende orders met minstens één klacht (standaard)
 *           dos     → dossiers (unieke problemen)
 *           all     → alle klachten
 * Noemer  = verkoop in dezelfde periode (BI.FACTSALES op de AS400, via query 6 en 7):
 *           - klant gekozen            → omzet en orders van die klant(en), per volledig jaar
 *           - verkoper / directeur     → omzet en orders van de eigen klanten met klachten, per volledig jaar
 *           - product(en) gekozen      → verkoop van die producten, per maand
 *           - productgroep(en) gekozen → verkoop van die productgroepen, per maand
 *           - anders                   → alle verkoop, per maand
 * Filters die niet in de verkoopdata bestaan (TREE, NTR, foutcode, status, VT, regio, actie, vervolg,
 * open, ter info, zoekterm) tellen enkel in de teller. Dat staat in de toelichting boven het rapport.
 */
'use strict';

let FIN = null;       // laatste berekening (ook voor het PDF-rapport)
let finChart = null;
const FIN_MIN_KLANT = 10000;   // minimumomzet om in de top-lijst klanten te komen
const FIN_MIN_PROD = 25000;    // idem voor producten

const fnf1 = () => new Intl.NumberFormat(currentUiLang() + '-BE', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const fpct = (v) => (v == null || !Number.isFinite(v) ? '–' : fnf1().format(v) + '%');
const monthIdx = (dn) => { const d = new Date(dn * 864e5); return d.getUTCFullYear() * 12 + d.getUTCMonth(); };
const midxLabel = (mi) => t('months').split(',')[mi % 12] + ' ' + String(Math.floor(mi / 12)).slice(2);
const isScoped = () => ME && ['sales', 'dir'].includes(ME.role);

function finRange() {
  let a = F.from, b = F.to;
  if (a == null || b == null) {
    let mn = Infinity, mx = -Infinity; for (const r of DATA.rows) { if (r.d < mn) mn = r.d; if (r.d > mx) mx = r.d; }
    if (a == null) a = mn; if (b == null) b = mx;
  }
  const m0 = monthIdx(a), m1 = monthIdx(b);
  return { m0, m1, y0: Math.floor(m0 / 12), y1: Math.floor(m1 / 12) };
}

// Teller voor een groep klachten
function tellerOf(rows) {
  if (UI.finTeller === 'all') return rows.length;
  const s = new Set();
  if (UI.finTeller === 'dos') { for (const r of rows) s.add(r.dos); return s.size; }
  for (const r of rows) if (isOrder(r.order)) s.add(r.order);
  return s.size;
}
const ratioOf = (teller, orders) => (orders > 0 ? (teller / orders) * 100 : null);

// Welke noemer geldt voor de huidige filters?
function finDenomMode() {
  const S = DATA.sales || {};
  if (F.klant.size) return S.klant ? 'klant' : null;
  if (isScoped()) return S.klant ? 'klantAll' : null;
  return S.prod ? 'prod' : null;
}
// In dit rapport volgt de productgroep de VERKOOPINDELING van het product van de klacht.
// De productgroep op de klacht zelf (KPROD) wijkt daar soms van af (bv. Pergola Z!P wordt verkocht
// onder Outdoor living); door de verkoopindeling te gebruiken passen teller en noemer bij elkaar.
let SPG = null, SPG_FOR = null;
function salesPg(prod) {
  if (SPG_FOR !== DATA) {
    const c = new Map();
    if (DATA.sales && DATA.sales.prod) for (const x of DATA.sales.prod) { let m = c.get(x[2]); if (!m) c.set(x[2], (m = new Map())); m.set(x[3], (m.get(x[3]) || 0) + Math.abs(x[4]) + 1); }
    SPG = new Map([...c].map(([p, m]) => [p, [...m].sort((a, b) => b[1] - a[1])[0][0]]));
    SPG_FOR = DATA;
  }
  return SPG.get(prod) || null;
}
function prodMatch(x) { // x = [y, m, prod, pg, omzet, aantal, orders]
  if (F.prod.size) return F.prod.has(String(x[2]));
  if (F.pg.size) return F.pg.has(x[3]);
  return true;
}
// De klachten voor dit rapport: alle filters, maar de productgroep volgens de verkoopindeling.
function finRows(mode) {
  if (mode !== 'prod') {
    // omzet per klant bestaat enkel per jaar: dan ook de klachten van die volledige jaren nemen
    const R = finRange(); const a = Date.UTC(R.y0, 0, 1) / 864e5, b = Date.UTC(R.y1, 11, 31) / 864e5;
    return DATA.rows.filter((r) => r.d >= a && r.d <= b && matches(r, 'period'));
  }
  if (!F.pg.size) return filtered;
  return DATA.rows.filter((r) => matches(r, 'pg') && F.pg.has(salesPg(r.prod)));
}
function klantKeys(mode) {
  if (mode === 'klant') return [...F.klant];
  return [...new Set(DATA.rows.map((r) => r.klant))];
}

// Noemer: totaal + per periode (sleutel = functie van maandindex of jaar)
function finDenom(mode, R, keyOfMonth) {
  const S = DATA.sales;
  const per = new Map(); let omzet = 0, orders = 0;
  const add = (k, o, n) => { omzet += o; orders += n; if (k == null) return; const e = per.get(k) || [0, 0]; e[0] += o; e[1] += n; per.set(k, e); };
  if (mode === 'prod') {
    for (const x of S.prod) {
      const mi = x[0] * 12 + (x[1] - 1); if (mi < R.m0 || mi > R.m1 || !prodMatch(x)) continue;
      add(keyOfMonth ? keyOfMonth(mi) : null, x[4], x[6]);
    }
  } else {
    for (const k of klantKeys(mode)) {
      const ys = S.klant[k]; if (!ys) continue;
      for (let y = R.y0; y <= R.y1; y++) { const v = ys[y]; if (v) add(keyOfMonth ? keyOfMonth(y * 12) : null, v[0], v[1]); }
    }
  }
  return { omzet, orders, per };
}

function finDenomLabel(mode, R) {
  const per = mode === 'prod' ? `${midxLabel(R.m0)} – ${midxLabel(R.m1)}` : (R.y0 === R.y1 ? String(R.y0) : `${R.y0}–${R.y1}`);
  if (mode === 'klant') return t('finDenKlant', { n: F.klant.size, per });
  if (mode === 'klantAll') return t('finDenKlantAll', { per });
  if (F.prod.size) return t('finDenProd', { n: F.prod.size, per });
  if (F.pg.size) return t('finDenPg', { list: [...F.pg].map((c) => DATA.sales.pgn[c] || c).join(', '), per });
  return t('finDenAll', { per });
}

function tellerLabel() { return t({ orders: 'finKOrders', dos: 'finKDos', all: 'finKAll' }[UI.finTeller]); }
function ratioLabel() { return t(UI.finTeller === 'orders' ? 'finKRatioOrders' : UI.finTeller === 'dos' ? 'finKRatioDos' : 'finKRatioAll'); }

function renderFin() {
  const mode = finDenomMode();
  const R = finRange();
  // toelichting: welke filters tellen enkel voor de klachten?
  const only = [];
  const nm = { tree: 'TREE', ntr: 'NTR', fout: t('lbl_fout'), status: t('lbl_status'), vert: t('lbl_vert'), regio: t('lbl_regio'), actie: t('lbl_actie') };
  for (const k of Object.keys(nm)) if (F[k].size) only.push(nm[k]);
  if (F.open) only.push(t('lbl_open')); if (F.vervolg) only.push(t('lbl_vervolg')); if (F.info !== 'excl') only.push(t('lbl_terinfo')); if (F.text) only.push(t('chipSearch', { q: F.text }));
  if (F.klant.size && (F.pg.size || F.prod.size)) only.push(F.prod.size ? t('lbl_prod') : t('lbl_pg'));
  $('finNote').innerHTML = esc(t('finNoteBase') + ' ' + t('finNotePg')) + (only.length ? ' <b>' + esc(t('finNoteOnly', { list: only.join(', ') })) + '</b>' : '');

  if (!mode) {
    $('finBody').hidden = true; $('finEmpty').hidden = false;
    $('finEmpty').textContent = t(isScoped() || F.klant.size ? 'finNoKlantData' : 'finNoData');
    FIN = null; return;
  }
  $('finBody').hidden = false; $('finEmpty').hidden = true;

  // ── Kerncijfers ──
  const rows = finRows(mode);
  const teller = tellerOf(rows);
  const cost = rows.reduce((s, r) => s + r.cost, 0);
  const den = finDenom(mode, R, null);
  const ratio = ratioOf(teller, den.orders);
  const kp = den.omzet > 0 ? (cost / den.omzet) * 100 : null;
  const dl = finDenomLabel(mode, R);
  const noOrder = UI.finTeller === 'orders' ? rows.filter((r) => !isOrder(r.order)).length : 0;
  const kpis = [
    { l: tellerLabel(), v: nf.format(teller), c: t('finCalcTeller', { n: nf.format(rows.length) }) + (noOrder ? ' ' + t('finCalcNoOrder', { n: nf.format(noOrder) }) : '') },
    { l: t('finKSold'), v: nf.format(den.orders), c: dl },
    { l: ratioLabel(), v: fpct(ratio).replace('%', UI.finTeller === 'orders' ? '%' : ''), c: `${nf.format(teller)} / ${nf.format(den.orders)}` },
    { l: t('finKCost'), v: eur.format(cost), c: t('finCalcCost', { n: nf.format(rows.length) }) },
    { l: t('finKOmzet'), v: eur.format(den.omzet), c: dl },
    { l: t('finKCostPct'), v: fpct(kp), c: `${eur.format(cost)} / ${eur.format(den.omzet)}` },
  ];
  $('finKpi').innerHTML = kpis.map((k) => `<div class="stat"><b>${esc(k.v)}</b><span>${esc(k.l)}</span><span class="calc">${esc(k.c)}</span></div>`).join('');

  // ── Evolutie ──
  const gran = mode === 'prod' ? UI.finGran : 'Y';
  document.querySelectorAll('#segFinGran button').forEach((b) => { b.classList.toggle('on', b.dataset.v === gran); b.disabled = mode !== 'prod' && b.dataset.v !== 'Y'; });
  const keyOf = (mi) => { const y = Math.floor(mi / 12), m = mi % 12; return gran === 'Y' ? String(y) : gran === 'Q' ? `${y}-${t('qShort')}${Math.floor(m / 3) + 1}` : `${y}-${String(m + 1).padStart(2, '0')}`; };
  const keys = []; for (let mi = R.m0; mi <= R.m1; mi++) { const k = keyOf(mi); if (keys[keys.length - 1] !== k) keys.push(k); }
  const byKey = new Map(keys.map((k) => [k, []]));
  for (const r of rows) { const l = byKey.get(keyOf(monthIdx(r.d))); if (l) l.push(r); }
  const denP = finDenom(mode, R, keyOf).per;
  const prettyK = (k) => (gran === 'M' ? midxLabel(Number(k.slice(0, 4)) * 12 + Number(k.slice(5)) - 1) : k);
  const serRatio = [], serCost = [], cnt = [];
  for (const k of keys) {
    const l = byKey.get(k), d = denP.get(k) || [0, 0];
    const te = tellerOf(l), c = l.reduce((s, r) => s + r.cost, 0);
    serRatio.push(d[1] ? Math.round(ratioOf(te, d[1]) * 10) / 10 : null);
    serCost.push(d[0] ? Math.round((c / d[0]) * 1000) / 10 : null);
    cnt.push({ te, o: d[1], c, om: d[0] });
  }
  if (finChart) finChart.destroy();
  const rl = ratioLabel(), cl = t('finKCostPct');
  finChart = new Chart($('finChart'), {
    type: 'line',
    data: { labels: keys.map(prettyK), datasets: [
      { label: rl, data: serRatio, yAxisID: 'y', borderColor: COLORS[0], backgroundColor: COLORS[0], borderWidth: 2.5, pointRadius: keys.length > 40 ? 0 : 3, tension: 0.25 },
      { label: cl, data: serCost, yAxisID: 'y1', borderColor: COLORS[1], backgroundColor: COLORS[1], borderWidth: 2, borderDash: [6, 4], pointRadius: keys.length > 40 ? 0 : 3, tension: 0.25 },
    ] },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false, interaction: { mode: 'index', intersect: false },
      plugins: { legend: { display: false }, tooltip: { backgroundColor: '#111', padding: 10, boxPadding: 4, callbacks: {
        label: (c) => { const x = cnt[c.dataIndex]; return c.datasetIndex === 0 ? ` ${rl}: ${fpct(c.raw)} (${nf.format(x.te)} / ${nf.format(x.o)})` : ` ${cl}: ${fpct(c.raw)} (${eur.format(x.c)} / ${eur.format(x.om)})`; },
      } } },
      scales: {
        x: { grid: { display: false }, ticks: { color: '#6b7280', maxRotation: 0, autoSkip: true, maxTicksLimit: 14 } },
        y: { beginAtZero: true, position: 'left', grid: { color: '#eef0f2' }, border: { display: false }, title: { display: true, text: rl, color: COLORS[0] }, ticks: { color: '#6b7280', callback: (v) => v + '%' } },
        y1: { beginAtZero: true, position: 'right', grid: { display: false }, border: { display: false }, title: { display: true, text: cl, color: COLORS[1] }, ticks: { color: '#6b7280', callback: (v) => v + '%' } },
      },
    },
  });
  const maxD = DATA.rows.reduce((m, r) => Math.max(m, r.d), 0);
  const lastPartial = keys.length && keyOf(monthIdx(maxD + 1)) === keys[keys.length - 1];
  const evolNote = [t(mode === 'prod' ? 'finEvolNoteProd' : 'finEvolNoteKlant')]; if (lastPartial) evolNote.push(t('finEvolPartial'));
  $('finLegend').innerHTML = `<span><i style="background:${COLORS[0]}"></i>${esc(rl)}</span><span><i style="background:${COLORS[1]}"></i>${esc(cl)}</span><div style="flex-basis:100%;color:#6b7280">${esc(evolNote.join(' '))}</div>`;

  // ── Tabel per productgroep (of per jaar bij een klant) ──
  const S = DATA.sales;
  const head = [t('finColOmzet'), t('finKSold'), tellerLabel(), ratioLabel(), t('finKCost'), t('finKCostPct')];
  let pgRows = [], pgTitle;
  if (mode === 'prod') {
    pgTitle = t('finPgTitle');
    const sal = new Map();
    for (const x of S.prod) { const mi = x[0] * 12 + (x[1] - 1); if (mi < R.m0 || mi > R.m1 || !prodMatch(x)) continue; const e = sal.get(x[3]) || [0, 0]; e[0] += x[4]; e[1] += x[6]; sal.set(x[3], e); }
    const kl = new Map(); for (const r of rows) { const pg = salesPg(r.prod) || '–'; let l = kl.get(pg); if (!l) kl.set(pg, (l = [])); l.push(r); }
    for (const pg of new Set([...sal.keys(), ...kl.keys()])) {
      const e = sal.get(pg) || [0, 0], l = kl.get(pg) || []; const c = l.reduce((s, r) => s + r.cost, 0), te = tellerOf(l);
      if (!e[0] && !l.length) continue;
      pgRows.push({ key: pg === '–' ? null : pg, name: pg === '–' ? t('finNoSalesPg') : (S.pgn[pg] || pg), om: e[0], o: e[1], te, ra: ratioOf(te, e[1]), c, kp: e[0] > 0 ? (c / e[0]) * 100 : null });
    }
    pgRows.sort((a, b) => b.om - a.om);
  } else {
    pgTitle = t('finYearTitle');
    for (let y = R.y0; y <= R.y1; y++) {
      const d = finDenom(mode, { m0: y * 12, m1: y * 12 + 11, y0: y, y1: y }, null);
      const l = rows.filter((r) => new Date(r.d * 864e5).getUTCFullYear() === y); const c = l.reduce((s, r) => s + r.cost, 0), te = tellerOf(l);
      pgRows.push({ key: null, name: String(y), om: d.omzet, o: d.orders, te, ra: ratioOf(te, d.orders), c, kp: d.omzet > 0 ? (c / d.omzet) * 100 : null });
    }
  }
  $('finPgCard').querySelector('[data-i18n="finPgTitle"],[data-i18n="finYearTitle"]').textContent = pgTitle;
  const rowHtml = (x, extra) => `<tr${x.key != null ? ` data-k="${esc(x.key)}"` : ''}><td>${esc(x.name)}</td><td class="num">${eur.format(x.om)}</td><td class="num">${nf.format(x.o)}</td><td class="num">${nf.format(x.te)}</td><td class="num">${fpct(x.ra)}</td><td class="num">${eur.format(x.c)}</td><td class="num"><b>${fpct(x.kp)}</b></td>${extra || ''}</tr>`;
  const firstCol = mode === 'prod' ? t('lbl_pg') : t('year');
  $('finPg').innerHTML = `<thead><tr><th>${esc(firstCol)}</th>${head.map((h) => `<th class="num">${esc(h)}</th>`).join('')}</tr></thead><tbody>${pgRows.map((x) => rowHtml(x)).join('') || `<tr><td colspan="7" class="empty">${esc(t('noData'))}</td></tr>`}</tbody>`;
  $('finPg').querySelectorAll('tr[data-k]').forEach((tr) => tr.addEventListener('click', () => addFilter('pg', tr.dataset.k)));

  // ── Top klanten (kost t.o.v. omzet) ──
  const klRows = [];
  if (S.klant) {
    const kl = new Map(); for (const r of rows) { let l = kl.get(r.klant); if (!l) kl.set(r.klant, (l = [])); l.push(r); }
    for (const [k, l] of kl) {
      const ys = S.klant[k]; if (!ys) continue; let om = 0, o = 0;
      for (let y = R.y0; y <= R.y1; y++) if (ys[y]) { om += ys[y][0]; o += ys[y][1]; }
      if (om < FIN_MIN_KLANT) continue;
      const c = l.reduce((s, r) => s + r.cost, 0), te = tellerOf(l);
      klRows.push({ key: k, name: `${l[0].klantNaam || t('nameUnknown')} (${l[0].klnr})`, om, o, te, ra: ratioOf(te, o), c, kp: (c / om) * 100 });
    }
    klRows.sort((a, b) => b.kp - a.kp);
  }
  const top = (list) => list.slice(0, 15);
  const tblHead = (first) => `<thead><tr><th>${esc(first)}</th>${head.map((h) => `<th class="num">${esc(h)}</th>`).join('')}</tr></thead>`;
  $('finKlant').innerHTML = tblHead(t('lbl_klant')) + `<tbody>${top(klRows).map((x) => rowHtml(x)).join('') || `<tr><td colspan="7" class="empty">${esc(t('noData'))}</td></tr>`}</tbody>`;
  $('finKlant').querySelectorAll('tr[data-k]').forEach((tr) => tr.addEventListener('click', () => addFilter('klant', tr.dataset.k)));
  const yrs = R.y0 === R.y1 ? String(R.y0) : `${R.y0}–${R.y1}`;
  $('finKlantSub').textContent = t('finKlantSub', { min: eur.format(FIN_MIN_KLANT), yrs }) + (F.pg.size || F.prod.size ? ' ' + t('finKlantSubAll') : '');

  // ── Top producten (enkel met verkoop per product) ──
  const prRows = [];
  $('finProdCard').hidden = mode !== 'prod';
  if (mode === 'prod') {
    const sal = new Map();
    for (const x of S.prod) { const mi = x[0] * 12 + (x[1] - 1); if (mi < R.m0 || mi > R.m1 || !prodMatch(x)) continue; const e = sal.get(x[2]) || [0, 0]; e[0] += x[4]; e[1] += x[6]; sal.set(x[2], e); }
    const kl = new Map(); for (const r of rows) { let l = kl.get(r.prod); if (!l) kl.set(r.prod, (l = [])); l.push(r); }
    for (const [p, e] of sal) {
      if (e[0] < FIN_MIN_PROD) continue;
      const l = kl.get(p) || []; const c = l.reduce((s, r) => s + r.cost, 0), te = tellerOf(l);
      prRows.push({ key: String(p), name: `${S.pn[p] || (l[0] && l[0].prodNaam) || 'Product ' + p} (${p})`, om: e[0], o: e[1], te, ra: ratioOf(te, e[1]), c, kp: (c / e[0]) * 100 });
    }
    prRows.sort((a, b) => b.kp - a.kp);
    $('finProd').innerHTML = tblHead(t('lbl_prod')) + `<tbody>${top(prRows).map((x) => rowHtml(x)).join('') || `<tr><td colspan="7" class="empty">${esc(t('noData'))}</td></tr>`}</tbody>`;
    $('finProd').querySelectorAll('tr[data-k]').forEach((tr) => tr.addEventListener('click', () => addFilter('prod', tr.dataset.k)));
    $('finProdSub').textContent = t('finProdSub', { min: eur.format(FIN_MIN_PROD) });
  }

  FIN = { kpis, note: $('finNote').textContent, head, firstCol, pgTitle, pgRows, klRows: top(klRows), prRows: top(prRows), klSub: $('finKlantSub').textContent, prSub: mode === 'prod' ? $('finProdSub').textContent : '', evolNote: evolNote.join(' '), mode };
}

/* ── bediening ── */
$('segFinTeller').addEventListener('click', (e) => {
  const v = e.target.dataset.v; if (!v) return; UI.finTeller = v;
  $('segFinTeller').querySelectorAll('button').forEach((b) => b.classList.toggle('on', b.dataset.v === v));
  if (DATA) renderFin();
});
$('segFinGran').addEventListener('click', (e) => {
  const v = e.target.dataset.v; if (!v || e.target.disabled) return; UI.finGran = v;
  if (DATA) renderFin();
});
