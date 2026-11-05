/* Winsol Klachtenanalyse – online versie
 * Login via de Cloudflare Worker (worker.js). De data komt van de Worker, per
 * vertegenwoordiger opgesplitst in gzip-blokken; ongewijzigde blokken komen uit
 * de browseropslag (IndexedDB). Een admin laadt KLACHTEN_DATA.xlsx op (4 AS400-
 * queries: KLACHTEN, TREE, OORZAAK, LOOKUP): de browser rekent alles uit op de
 * volledige data en stuurt enkel de gewijzigde blokken.
 */
'use strict';

/* ───────────── Hulpfuncties ───────────── */
const $ = (id) => document.getElementById(id);
let nf = new Intl.NumberFormat('nl-BE');
let eur = new Intl.NumberFormat('nl-BE', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
function setNumberLocale() { const l = (typeof currentUiLang === 'function' ? currentUiLang() : 'nl') + '-BE'; nf = new Intl.NumberFormat(l); eur = new Intl.NumberFormat(l, { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 }); }
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const txt = (s) => String(s ?? '').replace(/_x000D_/g, '').replace(/\r/g, '').trim(); // Excel-restjes van regeleindes
const clean = (s) => txt(s).replace(/\|/g, '!'); // SO| -> SO!, Z|P -> Z!P
const num = (v) => { const n = Number(v); return Number.isFinite(n) ? n : 0; };
// Bedragen boven 100.000 € zijn invoerfouten (bv. een ordernummer in het bedragveld) en tellen niet mee.
const amt = (v) => { const n = num(v); return n > 100000 || n < 0 ? 0 : n; };
const FIRMA = { W: 'Winsol', B: 'Winbalu', H: 'Helios' };
const COLORS = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6'];
const OTHER_COLOR = '#9ca3af';
// Statussen die als afgesloten tellen (A afgesloten, W geweigerd, I/O ter info, Y interventie geweigerd)
const CLOSED_STATUS = new Set(['A', 'W', 'I', 'O', 'Y']);
const INFO_STATUS = new Set(['I', 'O']);
const INTERV_STATUS = new Set(['N', 'V', 'Y']); // Interventies (herstellingsflow), klacht met interventie, interventie geweigerd
const ACTIE = { 0: 'actie0', 1: 'actie1', 2: 'actie2', 3: 'actie3', 4: 'actie4', 5: 'actie5' };
const actieNaam = (a) => (a < 0 ? t('actieUnknown') : ACTIE[a] ? t(ACTIE[a]) : t('actieN', { n: a }));
const ntrLabel = (r) => (r.ntr ? `${r.ntr} – ${r.ntrNaam}` : t('ntrMissing'));
const vlinkLabel = (v) => (v === 'herstelorder' ? t('vlinkRepair') : v ? t('vlinkTree') : '');
const yesNo = (b) => (b ? t('yes') : t('no'));
const VERVOLG_DAGEN = 365; // een vervolgklacht valt binnen 1 jaar na de vorige klacht in hetzelfde dossier
const isOrder = (o) => o >= 1000000; // korte nummers (1, 12, 43…) zijn regionummers, geen order

/* Dossiers = problemen. Een klacht is een VERVOLGKLACHT (zelfde dossier) als ze binnen 1 jaar volgt op een eerdere klacht en
 *  1. op het herstelorder (NTR-order, KNORDER) van die eerdere klacht staat, of
 *  2. op hetzelfde order (KORNR) staat EN hetzelfde onderwerp heeft (zelfde TREE niveau 3, bij beide ingevuld).
 * Staat ze op hetzelfde order maar is het onderwerp anders of onbekend, dan is het een ANDERE KLACHT OP HETZELFDE ORDER:
 * een nieuw probleem, een eigen dossier. */
const fmtOrder = (o) => { const s = String(o || ''); return /^\d{9}$/.test(s) ? `${s.slice(0, 2)}-${s.slice(2, 4)}-${s.slice(4, 8)}-${s.slice(8)}` : s; };
const treeL3 = (r) => new Set(r.tree.map((p) => { const n = treeNodes(p); return n[Math.min(2, n.length - 1)]; }));
function linkDossiers(rows) {
  const sorted = rows.slice().sort((a, b) => a.d - b.d || a.id - b.id);
  const parent = new Map();
  const find = (x) => { while (parent.get(x) !== x) { parent.set(x, parent.get(parent.get(x))); x = parent.get(x); } return x; };
  const byOrder = new Map(), ntrOwner = new Map();
  for (const r of sorted) {
    parent.set(r.id, r.id); r.vervolg = false; r.vgap = null; r.vlink = ''; r.andereOpOrder = false;
    let pred = null;
    if (isOrder(r.order)) {
      // regel 1: klacht op het herstelorder van een eerdere klacht
      const p1 = ntrOwner.get(r.order);
      if (p1 && p1 !== r && r.d - p1.d <= VERVOLG_DAGEN) { pred = p1; r.vlink = 'herstelorder'; }
      // regel 2: zelfde order + zelfde TREE niveau 3 (meest recente eerdere klacht die past)
      const earlier = byOrder.get(r.order) || [];
      if (!pred && r.tree.length) {
        const mine = treeL3(r);
        for (let i = earlier.length - 1; i >= 0; i--) {
          const p = earlier.at(i);
          if (r.d - p.d > VERVOLG_DAGEN) break;
          if (p.tree.length && [...treeL3(p)].some((x) => mine.has(x))) { pred = p; r.vlink = 'zelfde order + TREE'; break; }
        }
      }
      if (!pred && earlier.length) r.andereOpOrder = true;
    }
    if (pred) {
      r.vervolg = true; r.vgap = r.d - pred.d;
      const a = find(pred.id), b = find(r.id); if (a !== b) parent.set(b, a);
    }
    if (isOrder(r.order)) { let l = byOrder.get(r.order); if (!l) byOrder.set(r.order, (l = [])); l.push(r); }
    if (isOrder(r.ntrOrder)) ntrOwner.set(r.ntrOrder, r);
  }
  const members = new Map();
  for (const r of rows) { r.dos = find(r.id); let m = members.get(r.dos); if (!m) members.set(r.dos, (m = [])); m.push(r.id); }
  for (const r of rows) r.dosSize = members.get(r.dos).length;
  DOSSIERS = members;
  // order van het dossier = het order van de eerste klacht; zo komen ook klachten op herstelorders bij hun oorspronkelijk order
  const byId = new Map(rows.map((r) => [r.id, r]));
  const ORD = new Map();
  for (const r of rows) {
    const root = byId.get(r.dos); r.rootOrder = isOrder(root.order) ? root.order : 0;
    if (r.rootOrder) { let l = ORD.get(r.rootOrder); if (!l) ORD.set(r.rootOrder, (l = [])); l.push(r.id); }
  }
  ORDERS = ORD;
}
/* Toegewezen product en productgroep (bij het inlezen, met de verkoopdata).
 * De NTR-kost moet in de productgroep van het product waar de klacht over gaat, zodat kost en omzet bij elkaar passen.
 * Productgroep = verkoopindeling (de groep waaronder het product verkocht wordt).
 *  1. product    : het artikel van de klacht is een echt product → de verkoopgroep van dat product (bv. Pergola Z!P → Outdoor living)
 *  2. order      : Diversen-/Webshop-artikel, onbekend artikel of artikel verkocht onder Diversen → het artikel van het order van de klacht
 *  3. klacht-pg  : geen bruikbaar order → de productgroep op de klacht, omgezet naar de verkoopindeling
 *  4. diversen   : ook dat lukt niet → blijft zoals het is (bv. Webshop Marketing: folders, brochures)
 * Wat op de klacht staat, blijft bewaard in pgK/pgKNaam/prodK/prodKNaam. */
const DIVERSEN_PG = '9';
function assignProduct(rows, sales, lk) {
  const w = new Map();
  for (const x of sales.prod) { let m = w.get(x[2]); if (!m) w.set(x[2], (m = new Map())); m.set(x[3], (m.get(x[3]) || 0) + Math.abs(x[4]) + 1); }
  const spg = new Map([...w].map(([p, m]) => [p, [...m].sort((a, b) => b[1] - a[1])[0][0]]));
  const pName = (p) => sales.pn[p] || lk('PRODUCT', p) || '';
  const generic = (p) => { if (!p || !spg.has(p) || spg.get(p) === DIVERSEN_PG) return true; const n = pName(p).toUpperCase(); return n.includes('DIVERS') || n.includes('WEBSHOP'); };
  // productgroep op de klacht → verkoopgroep: de verkoopgroep die het vaakst voorkomt bij echte producten met die klacht-productgroep
  const cnt = new Map();
  for (const r of rows) { if (generic(r.prod) || !r.pg || r.pg === DIVERSEN_PG) continue; let m = cnt.get(r.pg); if (!m) cnt.set(r.pg, (m = new Map())); const g = spg.get(r.prod); m.set(g, (m.get(g) || 0) + 1); }
  const pgMap = new Map([...cnt].map(([pg, m]) => [pg, [...m].sort((a, b) => b[1] - a[1])[0][0]]));
  for (const r of rows) {
    Object.assign(r, { pgK: r.pg, pgKNaam: r.pgNaam, prodK: r.prod, prodKNaam: r.prodNaam });
    let pg, prod = r.prod, via;
    if (!generic(r.prod)) { pg = spg.get(r.prod); via = 'product'; }
    else if (r.oArt && !generic(r.oArt)) { prod = r.oArt; pg = spg.get(prod); via = 'order'; }
    else if (pgMap.get(r.pg) && pgMap.get(r.pg) !== DIVERSEN_PG) { pg = pgMap.get(r.pg); via = 'klacht'; }
    else { pg = spg.get(r.prod) || (pgMap.get(r.pg) || r.pg); via = 'diversen'; }
    r.aVia = via;
    if (pg !== r.pg) { r.pg = pg; r.pgNaam = sales.pgn[pg] || lk('PG', pg) || pg; }
    if (prod !== r.prod) { r.prod = prod; r.prodNaam = pName(prod) || 'Product ' + prod; }
  }
}
let ORDERS = new Map();
let DOSSIERS = new Map();
const isInterv = (v) => { const s = String(v ?? '').trim().toUpperCase(); return s !== '' && !['N', '0', 'NEE', 'NO'].includes(s); };
const WINPRO_KLACHT = 'https://winpro.winsol.eu:10089/winpro/public/klacht/detail/id/';

function ymdToDate(v) {
  // AS400-datums komen als getal JJJJMMDD; Excel kan ook een echte datum geven.
  if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const n = Math.round(num(v));
  if (n < 19000101) return null;
  const y = Math.floor(n / 10000), m = Math.floor(n / 100) % 100, d = n % 100;
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  return new Date(y, m - 1, d);
}
const dayNum = (d) => (d ? Math.floor(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 864e5) : null);
const fmtDate = (dn) => { if (dn == null) return ''; const d = new Date(dn * 864e5); return `${String(d.getUTCDate()).padStart(2, '0')}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${d.getUTCFullYear()}`; };
const isoFromDay = (dn) => new Date(dn * 864e5).toISOString().slice(0, 10);
const dayFromIso = (s) => (s ? Math.floor(Date.parse(s + 'T00:00:00Z') / 864e5) : null);

/* ───────────── Opslag in de browser (IndexedDB) ───────────── */
const DB = {
  open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('klachtenanalyse', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  },
  async get(k) {
    try { const db = await this.open(); return await new Promise((res) => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => res(undefined); }); }
    catch { return undefined; }
  },
  async clear() {
    try { const db = await this.open(); await new Promise((res) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').clear(); t.oncomplete = res; t.onerror = res; }); }
    catch { /* */ }
  },
  async set(k, v) {
    try { const db = await this.open(); await new Promise((res) => { const t = db.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = res; t.onerror = res; }); }
    catch { /* opslag niet beschikbaar: geen probleem, enkel trager bij volgende start */ }
  },
};

/* ───────────── Excel inlezen ───────────── */
function sheetRows(ws) {
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
}
function findSheets(wb) {
  const found = {};
  for (const name of wb.SheetNames) {
    const rows = sheetRows(wb.Sheets[name]);
    if (!rows.length) continue;
    // De kop kan op rij 1 staan (Power Query); zoek de eerste rij met gekende kolommen.
    for (let i = 0; i < Math.min(rows.length, 5); i++) {
      const hdr = (rows[i] || []).map((h) => String(h ?? '').trim().toUpperCase());
      const kind = hdr.includes('KVOLGN') && hdr.includes('KDATE') ? 'klachten'
        : hdr.includes('TVOLGN') && hdr.includes('TTREEOM') ? 'tree'
        : hdr.includes('BVOLGN') ? 'oorzaak'
        : hdr.includes('LTYPE') && hdr.includes('LCODE') ? 'lookup'
        : hdr.includes('OORNR') && hdr.includes('OBRON') ? 'orders'
        : hdr.includes('VPROD') && hdr.includes('PMAAND') ? 'vprodP'
        : hdr.includes('VPROD') && hdr.includes('VMAAND') ? 'vprod'
        : hdr.includes('VKLANT') && hdr.includes('VOMZET') ? 'vklant'
        : hdr.includes('POORNR') && hdr.includes('POPLAA') ? 'plaatsers' : null;
      if (kind) {
        const idx = {}; hdr.forEach((h, j) => { idx[h] = j; });
        found[kind] = { idx, rows: rows.slice(i + 1) };
        break;
      }
    }
  }
  return found;
}

function buildData(sheets, fileName, fileDate) {
  const miss = ['klachten', 'tree', 'lookup'].filter((k) => !sheets[k]);
  if (miss.length) throw new Error('Tabblad niet gevonden voor: ' + miss.join(', ') + '. Controleer of de queries uit DATA_OPHALEN.md geladen zijn.');

  // Opzoektabellen
  const L = {};
  { const { idx, rows } = sheets.lookup;
    for (const r of rows) {
      const t = String(r[idx.LTYPE] ?? '').trim(); if (!t) continue;
      const code = String(r[idx.LCODE] ?? '').trim().replace(/\.0*$/, '').replace(/^0+(?=\d)/, '');
      (L[t] ||= {});
      if (!(code in L[t])) L[t][code] = { naam: clean(r[idx.LNAAM]), extra: String(r[idx.LEXTRA] ?? '').trim() };
    }
  }
  const lk = (t, c) => { const k = String(c).trim().replace(/^0+(?=\d)/, ''); return (L[t] && L[t][k]) ? L[t][k].naam : ''; };

  // Tree per klacht (dubbels eruit)
  const trees = new Map();
  { const { idx, rows } = sheets.tree;
    for (const r of rows) {
      const k = num(r[idx.TVOLGN]); const p = clean(r[idx.TTREEOM]); if (!k || !p) continue;
      let a = trees.get(k); if (!a) trees.set(k, (a = []));
      if (!a.includes(p)) a.push(p);
    }
  }
  // Oorzaak per klacht
  const oorz = new Map();
  if (sheets.oorzaak) {
    const { idx, rows } = sheets.oorzaak;
    for (const r of rows) {
      const k = num(r[idx.BVOLGN]); if (!k) continue;
      const o = { oorzaak: txt(r[idx.OORZAAK]), corr: txt(r[idx.CORRIGEREND]), corrA: txt(r[idx.CORRECTIEF]) };
      if (!o.oorzaak && !o.corr && !o.corrA) continue;
      let a = oorz.get(k); if (!a) oorz.set(k, (a = [])); a.push(o);
    }
  }

  // Orderkop (query 5): per order de eerste bron in de volgorde Winsol, Winbalu, Helios, Industries.
  // (Een order van Winsol staat soms ook bij Helios als interne bestelling; die versie negeren we.)
  const ORD = new Map();
  if (sheets.orders) {
    const { idx: oi, rows: orows } = sheets.orders;
    const prio = { W: 1, B: 2, H: 3, I: 4 };
    const ddmmyyyy = (v) => { const n = Math.round(num(v)); if (n < 1010000) return null; const y = n % 10000, m = Math.floor(n / 10000) % 100, d = Math.floor(n / 1000000); return m >= 1 && m <= 12 && d >= 1 && d <= 31 ? dayNum(new Date(y, m - 1, d)) : null; };
    for (const r of orows) {
      const o = num(r[oi.OORNR]); if (!o) continue;
      const pr = prio[String(r[oi.OBRON] ?? '').trim()] || 9;
      const cur = ORD.get(o); if (cur && cur.pr <= pr) continue;
      const jj = num(r[oi.ODATJJ]), mm = num(r[oi.ODATMM]), dd = num(r[oi.ODATDD]);
      ORD.set(o, {
        pr, fakt: num(r[oi.OFAKTW]), plk: num(r[oi.OVOOPL]), art: oi.OARTNR != null ? num(r[oi.OARTNR]) : 0, hoof: num(r[oi.OHOOFO]),
        d: jj > 1900 && mm >= 1 && mm <= 12 && dd >= 1 ? dayNum(new Date(jj, mm - 1, dd)) : null,
        dP: ddmmyyyy(r[oi.ODATPR]), dL: ddmmyyyy(r[oi.ODATLE]), dU: ddmmyyyy(r[oi.ODATUI]), lp: num(r[oi.OLEVPL]),
      });
    }
  }
  // Verkoop (query 6 en 7): compacte tabellen voor het rapport "Financieel".
  let sales = null;
  if (sheets.vprod) {
    const { idx: vi, rows: vrows } = sheets.vprod;
    const prod = [];
    for (const r of vrows) {
      const y = num(r[vi.VJAAR]), m = num(r[vi.VMAAND]); if (!y || !m) continue;
      prod.push([y, m, num(r[vi.VPROD]), String(r[vi.VPG] ?? '').trim(), Math.round(num(r[vi.VOMZET]) * 100) / 100, num(r[vi.VAANTAL]), num(r[vi.VORDERS])]);
    }
    // namen van producten en productgroepen (ook voor producten zonder klacht)
    const pn = {}, pgn = {};
    for (const x of prod) { if (!(x[2] in pn)) pn[x[2]] = lk('PRODUCT', x[2]) || ''; if (!(x[3] in pgn)) pgn[x[3]] = lk('PG', x[3]) || ''; }
    // Query 9: dezelfde verkoop per PRODUCTIEMAAND van het order (voor de cohort-grafiek)
    let prodP = null;
    if (sheets.vprodP) {
      const { idx: pi, rows: prows } = sheets.vprodP; prodP = [];
      for (const r of prows) {
        const y = num(r[pi.PJAAR]), m = num(r[pi.PMAAND]); if (y < 2000 || m < 1 || m > 12) continue;
        const x = [y, m, num(r[pi.VPROD]), String(r[pi.VPG] ?? '').trim(), Math.round(num(r[pi.VOMZET]) * 100) / 100, num(r[pi.VAANTAL]), num(r[pi.VORDERS])];
        prodP.push(x); if (!(x[2] in pn)) pn[x[2]] = lk('PRODUCT', x[2]) || ''; if (!(x[3] in pgn)) pgn[x[3]] = lk('PG', x[3]) || '';
      }
    }
    sales = { prod, prodP, pn, pgn };
  }
  let klantSales = null;
  if (sheets.vklant) {
    const { idx: ki, rows: krows } = sheets.vklant;
    const PFX = { WV: 'W', BA: 'B', HE: 'H' };
    klantSales = {};
    for (const r of krows) {
      const id = String(r[ki.VKLANT] ?? '').trim(); const f = PFX[id.slice(0, 2)]; if (!f) continue;
      const key = `${f}-${num(id.slice(2))}`, y = num(r[ki.VJAAR]);
      (klantSales[key] ||= {})[y] = [Math.round(num(r[ki.VOMZET]) * 100) / 100, num(r[ki.VORDERS])];
    }
  }

  const { idx, rows } = sheets.klachten;
  const g = (r, c) => (idx[c] != null ? r[idx[c]] : null);
  const out = [];
  for (const r of rows) {
    const id = num(g(r, 'KVOLGN')); if (!id) continue;
    const created = dayNum(ymdToDate(g(r, 'KDATE')));
    if (created == null) continue;
    const firma = String(g(r, 'KFIRMA') ?? '').trim();
    const klnr = num(g(r, 'KKLNR'));
    const closedDay = dayNum(ymdToDate(g(r, 'KDATA')));
    const ntrB = num(g(r, 'KNNTR')), ntrV = num(g(r, 'KNTR'));
    const ntr = ntrB || ntrV; // bevestigd, anders vermoedelijk
    const fout = num(g(r, 'KNFOUT')) || num(g(r, 'KFOUT'));
    const pg = String(g(r, 'KPROD') ?? '').trim();
    const prod = num(g(r, 'KARTNR'));
    const fase = String(g(r, 'KFASE') ?? '').trim();
    const t = trees.get(id) || [];
    out.push({
      id, firma, klnr,
      klant: `${firma}-${klnr}`,
      klantNaam: lk('KLANT_' + firma, klnr) || '',
      land: ((L['KLANT_' + firma] || {})[String(klnr)] || {}).extra || '',
      d: created,
      dB: dayNum(ymdToDate(g(r, 'KDATB'))), dM: dayNum(ymdToDate(g(r, 'KDATM'))), dD: dayNum(ymdToDate(g(r, 'KDATD'))),
      dK: dayNum(ymdToDate(g(r, 'KDATK'))), dA: closedDay,
      open: !CLOSED_STATUS.has(fase),
      info: INFO_STATUS.has(fase),
      // Interventie = vlag IINTVEN op de klacht, of status N/V/Y, of een TREE-pad onder INTERVENTIE
      actie: idx.KLEVPL != null ? num(g(r, 'KLEVPL')) : -1,
      // Interventie = uit te voeren actie "Interventie" (KLEVPL 4). Zonder die kolom: oude benadering (vlag, status, TREE).
      interv: idx.KLEVPL != null ? num(g(r, 'KLEVPL')) === 4 : (isInterv(g(r, 'IINTVEN')) || INTERV_STATUS.has(fase) || t.some((p) => /\/INTERVENTIE(\/|$)/i.test(p))),
      intervFlag: String(g(r, 'IINTVEN') ?? '').trim(),
      fase, faseNaam: lk('STATUS', fase) || fase,
      pg, pgNaam: lk('PG', pg) || pg || '(onbekend)',
      prod, prodNaam: lk('PRODUCT', prod) || (prod ? 'Product ' + prod : '(onbekend)'),
      ntr, ntrBevestigd: !!ntrB, ntrNaam: ntr ? (lk('NTR', ntr) || 'NTR ' + ntr) : 'NTR ontbreekt',
      fout, foutNaam: fout ? (lk('FOUT', fout) || 'Fout ' + fout) : '',
      cost: amt(g(r, 'KNBDRB')) + amt(g(r, 'KNWPL')), costMat: amt(g(r, 'KNBDRB')), costPl: amt(g(r, 'KNWPL')),
      order: num(g(r, 'KORNR')), ntrOrder: num(g(r, 'KNORDER')),
      vert: num(g(r, 'KVTNR')), vertNaam: lk('VERT', num(g(r, 'KVTNR'))) || '',
      regio: num(g(r, 'KREGIO')), regioNaam: lk('REGIO', num(g(r, 'KREGIO'))) || '',
      beh: String(g(r, 'KBEHA') ?? '').trim(),
      type: String(g(r, 'KTYPE') ?? '').trim(),
      web: num(g(r, 'KWEB')),
      oms: txt(g(r, 'OMSCHRIJVING')),
      opl: txt(g(r, 'OPLOSSING')),
      tree: t,
      oorz: oorz.get(id) || [],
    });
    const od = ORD.get(out[out.length - 1].order);
    if (od) {
      // productiedatum: van het order zelf, anders van het hoofdorder
      const hd = od.hoof && od.hoof !== out[out.length - 1].order ? ORD.get(od.hoof) : null;
      Object.assign(out[out.length - 1], { oFakt: od.fakt, oPlk: od.plk, oD: od.d, oDP: od.dP, oDL: od.dL, oDU: od.dU, oLp: od.lp, oArt: od.art, dPr: od.dP || (hd && hd.dP) || null });
    }
  }
  if (!out.length) throw new Error('Geen klachten gevonden in het bestand.');
  if (sales) assignProduct(out, sales, lk);
  linkDossiers(out);
  return { rows: out, sales, klantSales, fileName, fileDate, loadedAt: Date.now() };
}

/* ───────────── Toestand ───────────── */
let DATA = null;          // { rows, fileName, fileDate }
const F = { from: null, to: null, pg: new Set(), prod: new Set(), tree: new Set(), ntr: new Set(), fout: new Set(), klant: new Set(), status: new Set(), vert: new Set(), regio: new Set(), open: '', info: 'excl', actie: new Set(), vervolg: '', text: '' };
let HIDDEN_YEARS = new Set(); try { HIDDEN_YEARS = new Set(JSON.parse(localStorage.getItem('kl_hiddenYears') || '[]')); } catch { /* geen opslag */ }
const saveHidden = () => { try { localStorage.setItem('kl_hiddenYears', JSON.stringify([...HIDDEN_YEARS])); } catch { /* geen opslag */ } };
const UI = { report: 'klachten', finTeller: 'orders', finGran: 'Q', finAxis: 'klacht', finWin12: true, finCurve: 'ratio', mode: 'time', cumGran: 'W', target: '', gran: 'M', measure: 'n', split: '', page: 0, sort: { key: 'd', dir: -1 }, treeLevel: 3 };
let filtered = [];
let CHIPS = []; // labels van de actieve filters (ook voor het PDF-rapport)
let chart = null;

/* TREE: niveaus 1-3 betrouwbaar, de rest = detailtekst (kan zelf '/' bevatten) */
function treeNodes(path) {
  const s = path.split('/').map((x) => x.trim());
  const n = [s[0]];
  if (s.length > 1) n.push(s.slice(0, 2).join(' › '));
  if (s.length > 2) n.push(s.slice(0, 3).join(' › '));
  if (s.length > 3) n.push(s.slice(0, 3).join(' › ') + ' › ' + s.slice(3).join('/'));
  return n;
}

/* ───────────── Server (Cloudflare Worker) ───────────── */
// Adres van de Worker. Lokaal testen (npx wrangler dev): pagina openen via http://localhost:… dan gebruikt de app localhost:8787.
const API_URL = /^(localhost|127\.0\.0\.1)$/.test(location.hostname)
  ? 'http://localhost:8787/'
  : 'https://klachtenanalyse.gwenn-vanthournout.workers.dev/';
const TOKEN_KEY = 'klachtenanalyse_token';
const SALES_VT = '999999'; // gereserveerd bloknummer voor de verkoop per product (geen echte VT)
const PARALLEL = 6;
let authToken = '';
let ME = null; // { username, role, vts, regio }
let META = null; // data:meta zoals de Worker die voor deze gebruiker teruggeeft

const authHeaders = () => (authToken ? { 'X-Auth-Token': authToken } : {});
async function apiRequest(bodyObj) {
  const res = await fetch(API_URL, { method: 'POST', headers: { 'Content-Type': 'application/json', ...authHeaders() }, body: JSON.stringify(bodyObj) });
  let data = {}; try { data = await res.json(); } catch { /* geen JSON */ }
  return { ok: res.ok, status: res.status, data };
}
async function apiBinary(mode, params, body) {
  const q = new URLSearchParams({ mode, ...params });
  const res = await fetch(API_URL + '?' + q, { method: 'POST', headers: { ...authHeaders(), 'Content-Type': 'application/octet-stream' }, body: body || '' });
  if (!res.ok) { let msg = 'HTTP ' + res.status; try { msg = (await res.json()).error || msg; } catch { /* */ } const e = new Error(msg); e.status = res.status; throw e; }
  return res;
}
function handleAuthFailure(status) {
  if (status === 401) { showLogin(t('sessionExpired')); return true; }
  return false;
}

/* gzip in de browser (CompressionStream) */
async function gzip(str) { return new Uint8Array(await new Response(new Blob([str]).stream().pipeThrough(new CompressionStream('gzip'))).arrayBuffer()); }
async function gunzip(buf) { return new Response(new Blob([buf]).stream().pipeThrough(new DecompressionStream('gzip'))).text(); }
async function sha(str) { const b = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str)); return [...new Uint8Array(b)].slice(0, 16).map((x) => x.toString(16).padStart(2, '0')).join(''); }
// Vingerafdruk per klacht: enkel de velden uit de AS400 (niet de berekende dossier-velden).
const DERIVED = new Set(['vervolg', 'vgap', 'vlink', 'andereOpOrder', 'dos', 'dosSize', 'rootOrder', 'fp']);
function rowFp(r) {
  const s = JSON.stringify(r, (k, v) => (DERIVED.has(k) ? undefined : v));
  let h = 0x811c9dc5; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193); }
  return (h >>> 0).toString(16).padStart(8, '0');
}
async function pool(items, n, fn) {
  let i = 0; const run = async () => { while (i < items.length) { const k = i++; await fn(items[k], k); } };
  await Promise.all(Array.from({ length: Math.min(n, items.length) }, run));
}

/* ───────────── Laden ───────────── */
function setLoading(on, txt) { $('loading').classList.toggle('on', on); if (txt) $('loadingTxt').textContent = txt; }
function setStatus(msg, err) { const s = $('dataStatus'); s.textContent = msg || ''; s.classList.toggle('err', !!err); }

// Haalt de data op die deze gebruiker mag zien. Blokken die niet veranderd zijn, komen uit de browseropslag.
async function loadFromServer(retry = true) {
  setLoading(true, t('dataLoading'));
  try {
    const { ok, status, data } = await apiRequest({ mode: 'data_meta' });
    if (handleAuthFailure(status)) return;
    if (!ok) throw new Error(data.error || t('unknownError'));
    META = data;
    if (ME.role === 'admin') showLastUpload();
    if (!data.available) {
      DATA = null; $('main').hidden = true; $('pdfBtn').disabled = true;
      $('dataInfo').textContent = ME.role === 'admin' ? t('dataNoneAdmin') : t('dataNone');
      return;
    }
    const vts = Object.keys(data.chunks);
    const parts = new Array(vts.length);
    let done = 0;
    await pool(vts, PARALLEL, async (vt, k) => {
      const c = data.chunks[vt];
      const cached = await DB.get('c:' + vt);
      if (cached && cached.h === c.h) { parts[k] = cached.rows; }
      else {
        const res = await apiBinary('chunk_get', { vt });
        const payload = JSON.parse(await gunzip(await res.arrayBuffer()));
        DB.set('c:' + vt, { h: c.h, rows: payload });
        parts[k] = payload;
      }
      done++; setLoading(true, t('dataLoadingN', { done, total: vts.length }));
    });
    // Een blok is een lijst klachten (oud formaat) of { rows, ks } / { sp } (sinds de verkoopdata).
    const rows = [], klantSales = {}; let sp = null;
    for (const p of parts) {
      if (Array.isArray(p)) { rows.push(...p); continue; }
      if (p.rows) rows.push(...p.rows);
      if (p.ks) Object.assign(klantSales, p.ks);
      if (p.sp) sp = p;
    }
    applyData({ rows, sales: { prod: sp ? sp.sp : null, prodP: sp ? sp.spp || null : null, pn: sp ? sp.pn || {} : {}, pgn: sp ? sp.pgn || {} : {}, klant: Object.keys(klantSales).length ? klantSales : null }, version: data.version, uploadedAt: data.uploadedAt, fileName: data.fileName });
    setStatus('');
  } catch (e) {
    console.error(e);
    if (e.status === 401) { showLogin(t('sessionExpired')); return; }
    if (retry && e.status === 404) { setLoading(false); return loadFromServer(false); } // data werd net vernieuwd
    setStatus(t('dataError', { msg: e.message }), true);
  } finally {
    setLoading(false);
  }
}

// Dossier-velden zijn bij het opladen berekend op de volledige data; hier enkel de indexen opbouwen.
function indexDossiers(rows) {
  const dos = new Map(), ord = new Map();
  for (const r of rows) {
    let m = dos.get(r.dos); if (!m) dos.set(r.dos, (m = [])); m.push(r.id);
    if (r.rootOrder) { let l = ord.get(r.rootOrder); if (!l) ord.set(r.rootOrder, (l = [])); l.push(r.id); }
  }
  DOSSIERS = dos; ORDERS = ord;
}

function renderDataInfo() {
  if (!DATA) return;
  const rows = DATA.rows;
  const maxD = rows.reduce((m, r) => Math.max(m, r.d), 0);
  const up = DATA.uploadedAt ? new Date(DATA.uploadedAt) : null;
  $('dataInfo').innerHTML = t('dataInfo', { n: nf.format(rows.length), date: up ? up.toLocaleDateString(currentUiLang() + '-BE') : '–', last: fmtDate(maxD) }) +
    (ME && ['sales', 'dir'].includes(ME.role) ? '<br>' + esc(t('dataScope', { vts: ME.vts.join(', ') })) : '');
}

function applyData(data) {
  DATA = data;
  indexDossiers(data.rows);
  renderDataInfo();
  $('main').hidden = false;
  $('pdfBtn').disabled = false;
  if (F.from == null) quickPeriod('24');
  buildFilters();
  renderTabs();
  update();
}

/* ───────────── Login ───────────── */
function roleLabel(r) { return t('role_' + r); }
function showLogin(message) {
  authToken = ''; ME = null; DATA = null; META = null;
  try { localStorage.removeItem(TOKEN_KEY); sessionStorage.removeItem(TOKEN_KEY); } catch { /* */ }
  $('appRoot').style.display = 'none';
  $('loginScreen').style.display = 'flex';
  $('loginPassword').value = '';
  $('loginError').textContent = message || '';
  setLoading(false);
}
function showApp(user) {
  ME = user;
  $('loginScreen').style.display = 'none';
  $('appRoot').style.display = 'block';
  $('topbarUser').textContent = user.username;
  $('topbarRole').textContent = roleLabel(user.role);
  const isAdmin = user.role === 'admin';
  $('adminDataCard').hidden = !isAdmin;
  $('usersCard').hidden = !isAdmin;
  loadFromServer();
}
async function doLogin() {
  const username = $('loginUsername').value.trim(), password = $('loginPassword').value, remember = $('loginRemember').checked;
  $('loginError').textContent = '';
  if (!username || !password) { $('loginError').textContent = t('loginMissing'); return; }
  const btn = $('loginSubmitBtn'); btn.disabled = true; btn.textContent = t('loggingIn');
  try {
    const { ok, status, data } = await apiRequest({ mode: 'login', username, password, remember });
    if (!ok) { $('loginError').textContent = status === 401 ? t('loginFailed') : status === 429 ? t('tooManyAttempts') : (data.error || t('loginGenericFail')); return; }
    authToken = data.token;
    try { (remember ? localStorage : sessionStorage).setItem(TOKEN_KEY, authToken); } catch { /* enkel in geheugen */ }
    // Andere gebruiker dan vorige keer op deze browser? Dan de opgeslagen data wissen.
    if ((await DB.get('owner')) !== data.username) { await DB.clear(); await DB.set('owner', data.username); }
    showApp(data);
  } catch { $('loginError').textContent = t('noConnection'); }
  finally { btn.disabled = false; btn.textContent = t('loginBtn'); }
}
async function tryRestoreSession() {
  let token = '';
  try { token = localStorage.getItem(TOKEN_KEY) || sessionStorage.getItem(TOKEN_KEY) || ''; } catch { token = ''; }
  if (!token) { showLogin(); return; }
  authToken = token;
  try {
    const { ok, data } = await apiRequest({ mode: 'auth_me' });
    if (ok) showApp(data); else showLogin();
  } catch { showLogin(t('noConnectionRelogin')); }
}

/* ───────────── Admin: klachtendata opladen ───────────── */
let PENDING = null;
function wireDrop(zoneId, inputId, onFile) {
  const zone = $(zoneId), input = $(inputId);
  zone.addEventListener('click', () => input.click());
  zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', (e) => { e.preventDefault(); zone.classList.remove('drag-over'); if (e.dataTransfer.files.length) onFile(e.dataTransfer.files[0]); });
  input.addEventListener('change', (e) => { if (e.target.files.length) onFile(e.target.files[0]); input.value = ''; });
}
function setMsg(id, html, kind) { const el = $(id); el.innerHTML = html || ''; el.className = 'status' + (kind ? ' ' + kind : ''); }
function showLastUpload() {
  const m = META;
  if (!m || !m.available) { setMsg('lastUpload', ''); return; }
  const d = m.diff || {};
  setMsg('lastUpload', esc(t('lastUpload', { date: new Date(m.uploadedAt).toLocaleString(currentUiLang() + '-BE'), by: m.uploadedBy || '?', file: m.fileName || '?', new: d.new ?? 0, changed: d.changed ?? 0, removed: d.removed ?? 0 })));
}

async function prepareUpload(file) {
  PENDING = null; $('uploadBtn').disabled = true; setMsg('diffStatus', '');
  $('dataDrop').classList.remove('has-file');
  setLoading(true, t('reading'));
  await new Promise((r) => setTimeout(r, 30));
  try {
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array', dense: true, cellDates: true });
    const built = buildData(findSheets(wb), file.name, file.lastModified); // berekent dossiers op de volledige data
    if (DATA) indexDossiers(DATA.rows); // buildData overschreef de indexen van de huidige weergave
    const rows = built.rows.sort((a, b) => a.id - b.id);
    for (const r of rows) r.fp = rowFp(r);
    // Vergelijken met wat online staat (de admin ziet altijd alle data).
    const old = new Map((DATA ? DATA.rows : []).map((r) => [r.id, r.fp]));
    const diff = { new: [], changed: [], removed: [] };
    const seen = new Set();
    for (const r of rows) { seen.add(r.id); const f = old.get(r.id); if (f == null) diff.new.push(r.id); else if (f !== r.fp) diff.changed.push(r.id); }
    for (const id of old.keys()) if (!seen.has(id)) diff.removed.push(id);
    // Blokken per vertegenwoordiger
    const byVt = new Map();
    for (const r of rows) { const vt = String(r.vert || 0); let l = byVt.get(vt); if (!l) byVt.set(vt, (l = [])); l.push(r); }
    const oldChunks = (META && META.available && META.chunks) || {};
    const chunks = {}, changed = [];
    // Elk blok bevat de klachten van één VT + de omzet van de klanten in die klachten.
    // De verkoop per product zit in een apart blok (SALES_VT) dat enkel admin en user mogen lezen.
    const payloads = new Map();
    for (const [vt, list] of byVt) {
      const ks = {};
      if (built.klantSales) for (const r of list) if (built.klantSales[r.klant] && !ks[r.klant]) ks[r.klant] = built.klantSales[r.klant];
      payloads.set(vt, { rows: list, ks });
    }
    if (built.sales) payloads.set(SALES_VT, { sp: built.sales.prod, spp: built.sales.prodP, pn: built.sales.pn, pgn: built.sales.pgn });
    for (const [vt, payload] of payloads) {
      const list = payload.rows || [];
      const json = JSON.stringify(payload);
      const h = await sha(json);
      chunks[vt] = { n: list.length, h, bytes: (oldChunks[vt] && oldChunks[vt].h === h) ? oldChunks[vt].bytes : 0 };
      if (!oldChunks[vt] || oldChunks[vt].h !== h) changed.push({ vt, json, payload });
    }
    const removedVts = Object.keys(oldChunks).filter((vt) => !chunks[vt]);
    PENDING = { chunks, changed, diff, fileName: file.name, fileDate: file.lastModified, nothing: !changed.length && !removedVts.length };
    $('dataDrop').classList.add('has-file');
    const same = rows.length - diff.new.length - diff.changed.length;
    setMsg('diffStatus', PENDING.nothing ? esc(t('diffNothing')) : t('diffSummary', { n: nf.format(rows.length), new: nf.format(diff.new.length), changed: nf.format(diff.changed.length), removed: nf.format(diff.removed.length), same: nf.format(same), chunks: changed.length, total: payloads.size }), PENDING.nothing ? 'ok' : '');
    $('uploadBtn').disabled = PENDING.nothing;
  } catch (e) {
    console.error(e);
    setMsg('diffStatus', esc(t('readError', { msg: e.message })), 'err');
  } finally { setLoading(false); }
}

async function doUpload() {
  if (!PENDING) return;
  const P = PENDING; const btn = $('uploadBtn'); btn.disabled = true;
  try {
    let done = 0; const total = P.changed.length;
    setMsg('diffStatus', esc(t('uploading', { done, total })));
    await pool(P.changed, 4, async (c) => {
      const gz = await gzip(c.json);
      P.chunks[c.vt].bytes = gz.byteLength;
      for (let attempt = 1; ; attempt++) {
        try { await apiBinary('chunk_put', { vt: c.vt, h: P.chunks[c.vt].h }, gz); break; }
        catch (e) { if (attempt >= 3 || e.status === 401 || e.status === 403) throw e; await new Promise((r) => setTimeout(r, 1000 * attempt)); }
      }
      done++; setMsg('diffStatus', esc(t('uploading', { done, total })));
    });
    await apiBinary('diff_put', {}, await gzip(JSON.stringify({ at: new Date().toISOString(), ...P.diff })));
    const { ok, status, data } = await apiRequest({ mode: 'data_commit', chunks: P.chunks, fileName: P.fileName, fileDate: P.fileDate, diff: { new: P.diff.new.length, changed: P.diff.changed.length, removed: P.diff.removed.length } });
    if (handleAuthFailure(status)) return;
    if (!ok) throw new Error(data.error || t('unknownError'));
    // De opgeladen blokken meteen in de browseropslag zetten: niet opnieuw downloaden.
    for (const c of P.changed) await DB.set('c:' + c.vt, { h: P.chunks[c.vt].h, rows: c.payload });
    PENDING = null; $('dataDrop').classList.remove('has-file');
    setMsg('diffStatus', esc(t('uploadDone', { n: nf.format(data.rows), new: nf.format(data.diff.new), changed: nf.format(data.diff.changed), removed: nf.format(data.diff.removed) })), 'ok');
    await loadFromServer();
  } catch (e) {
    console.error(e);
    if (e.status === 401) { showLogin(t('sessionExpired')); return; }
    setMsg('diffStatus', esc(t('uploadError', { msg: e.message })), 'err');
    btn.disabled = false;
  }
}

/* ───────────── Admin: gebruikers ───────────── */
let PENDING_USERS = null;
async function prepareUsers(file) {
  PENDING_USERS = null; $('usersUpdateBtn').disabled = true; setMsg('usersStatus', '');
  try {
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const users = [];
    for (const name of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '' });
      const hi = rows.findIndex((r) => r.some((c) => String(c).trim().toUpperCase() === 'USER'));
      if (hi < 0) continue;
      const hdr = rows[hi].map((c) => String(c).trim().toUpperCase());
      const col = (n) => hdr.indexOf(n);
      const [cU, cP, cR, cV, cG] = ['USER', 'PW', 'ROLE', 'VTNR', 'REGIO'].map(col);
      for (const r of rows.slice(hi + 1)) {
        const username = String(r[cU] ?? '').trim(); if (!username) continue;
        users.push({ username, password: cP >= 0 ? String(r[cP] ?? '') : '', role: cR >= 0 ? String(r[cR] ?? '') : '', vtnr: cV >= 0 ? String(r[cV] ?? '') : '', regio: cG >= 0 ? String(r[cG] ?? '') : '' });
      }
      break;
    }
    if (!users.length) { setMsg('usersStatus', esc(t('usersNoRows')), 'err'); return; }
    PENDING_USERS = users;
    $('usersDrop').classList.add('has-file');
    setMsg('usersStatus', esc(`${file.name}: ${users.length} gebruikers`));
    $('usersUpdateBtn').disabled = false;
  } catch (e) { setMsg('usersStatus', esc(t('readError', { msg: e.message })), 'err'); }
}
async function doUsers() {
  if (!PENDING_USERS) return;
  $('usersUpdateBtn').disabled = true;
  try {
    const { ok, status, data } = await apiRequest({ mode: 'admin_upsert_users', users: PENDING_USERS });
    if (handleAuthFailure(status)) return;
    if (!ok) throw new Error(data.error || t('unknownError'));
    PENDING_USERS = null; $('usersDrop').classList.remove('has-file');
    setMsg('usersStatus', esc(t('usersDone', data)) + (data.errors && data.errors.length ? '<br>' + data.errors.map(esc).join('<br>') : ''), data.errors && data.errors.length ? 'err' : 'ok');
    listUsers();
  } catch (e) { setMsg('usersStatus', esc(e.message), 'err'); $('usersUpdateBtn').disabled = false; }
}
async function listUsers() {
  const { ok, status, data } = await apiRequest({ mode: 'admin_list_users' });
  if (handleAuthFailure(status) || !ok) return;
  $('usersList').innerHTML = `<table class="user-tbl"><tr><th>USER</th><th>ROLE</th><th>VTNR</th><th>Regio</th></tr>${data.users.map((u) => `<tr><td>${esc(u.username)}</td><td>${esc(u.role)}</td><td>${esc((u.ownVts || u.vts || []).join('; '))}${u.role === 'dir' ? ` <span style="color:var(--muted)">(ziet: ${esc((u.vts || []).join('; '))})</span>` : ''}</td><td>${esc(u.regio || '')}</td></tr>`).join('')}</table>`;
}

/* ───────────── Productgroep-sets (codes uit PFPRODGRO) ─────────────
   Vul per knop de productgroep-codes in, bv. Winsol: ['1A','1B','2A'] */
const PG_PRESETS = {
  Winsol: ['4-', '4S', '4A', '2A', '4B', '4C', '4Z', '4', '2C', '5', '4E', '4D'],
  WinVista: ['1A', '1B', '3A', '1D', '1E', '3B', '8', '3D'],
  WinBalu: ['6'],
};

/* ───────────── Multi-select filters ───────────── */
const MS = {};
function makeMS(key, opts) {
  // opts: { placeholder, items: () => [{v, label, lvl, cnt}] , limit }
  const host = $('f_' + key);
  host.innerHTML = `<button type="button" class="ms-btn"></button>
    <div class="ms-panel"><input type="search" placeholder="${esc(t('msSearch'))}"><div class="ms-list"></div>
    ${opts.presets ? `<div class="ms-presets">${Object.keys(opts.presets).map((n) => `<button type="button" class="btn-outline ms-preset" data-p="${esc(n)}">${esc(n)}</button>`).join('')}</div>` : ''}
    <div class="ms-foot"><span class="info"></span><span class="ms-acts"><button type="button" class="btn-outline ms-all" title="${esc(t('msAllTip'))}">${esc(t('msAll'))}</button><button type="button" class="btn-outline ms-clear">${esc(t('msClear'))}</button></span></div></div>`;
  const btn = host.querySelector('.ms-btn'), list = host.querySelector('.ms-list'), search = host.querySelector('input[type=search]'), info = host.querySelector('.info');
  const st = { key, opts, host, items: opts.items() };
  const labelOf = new Map(st.items.map((i) => [i.v, i.label]));
  st.labelOf = (v) => labelOf.get(v) || v;
  st.total = st.items.length;
  const render = () => {
    const q = search.value.trim().toLowerCase();
    const limit = opts.limit || 400;
    let shown = 0, html = '';
    const sel = F[key];
    // geselecteerde eerst tonen
    const items = q ? st.items.filter((i) => i.search.includes(q)) : st.items;
    const ordered = sel.size ? [...items.filter((i) => sel.has(i.v)), ...items.filter((i) => !sel.has(i.v))] : items;
    for (const i of ordered) {
      if (shown >= limit) break;
      html += `<label class="ms-item lvl${i.lvl || 1}"><input type="checkbox" value="${esc(i.v)}"${sel.has(i.v) ? ' checked' : ''}><span class="lbl">${esc(i.label)}</span>${i.cnt != null ? `<span class="cnt">${nf.format(i.cnt)}</span>` : ''}</label>`;
      shown++;
    }
    list.innerHTML = html || `<div class="empty">${esc(t('msNone'))}</div>`;
    info.textContent = items.length > shown ? t('msShown', { n: shown, total: nf.format(items.length) }) : t('msOptions', { n: nf.format(items.length) });
  };
  st.render = render;
  st.syncBtn = () => {
    const sel = F[key];
    btn.classList.toggle('has', sel.size > 0);
    btn.textContent = sel.size === 0 ? t('optAll') : sel.size === 1 ? st.labelOf([...sel][0]) : t('msChosen', { n: sel.size });
    btn.title = [...sel].map(st.labelOf).join('\n');
  };
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    const open = host.classList.contains('open');
    document.querySelectorAll('.ms.open').forEach((m) => m.classList.remove('open'));
    if (!open) { host.classList.add('open'); render(); search.focus(); }
  });
  host.querySelector('.ms-panel').addEventListener('click', (e) => e.stopPropagation());
  search.addEventListener('input', render);
  list.addEventListener('change', (e) => {
    const v = e.target.value; if (e.target.checked) F[key].add(v); else F[key].delete(v);
    st.syncBtn(); UI.page = 0; update();
  });
  host.querySelector('.ms-clear').addEventListener('click', () => { F[key].clear(); st.syncBtn(); render(); UI.page = 0; update(); });
  host.querySelector('.ms-all').addEventListener('click', () => {
    const q = search.value.trim().toLowerCase();
    for (const i of st.items) if (!q || i.search.includes(q)) F[key].add(i.v);
    st.syncBtn(); render(); UI.page = 0; update();
  });
  host.querySelectorAll('.ms-preset').forEach((b) => b.addEventListener('click', () => {
    const codes = opts.presets[b.dataset.p] || [];
    if (!codes.length) { info.textContent = t('presetEmpty', { p: b.dataset.p }); return; }
    const have = new Set(st.items.map((i) => i.v));
    F[key].clear(); for (const c of codes) if (have.has(c)) F[key].add(c);
    st.syncBtn(); render(); UI.page = 0; update();
  }));
  st.syncBtn();
  MS[key] = st;
}
document.addEventListener('click', () => document.querySelectorAll('.ms.open').forEach((m) => m.classList.remove('open')));

function countBy(fn) {
  const m = new Map();
  for (const r of DATA.rows) { const ks = fn(r); for (const k of (Array.isArray(ks) ? ks : [ks])) { if (k == null || k === '') continue; const e = m.get(k[0]); if (e) e.cnt++; else m.set(k[0], { v: k[0], label: k[1], cnt: 1, lvl: k[2] }); } }
  return [...m.values()].map((i) => ({ ...i, search: (i.v + ' ' + i.label).toLowerCase() }));
}
const byCnt = (a, b) => b.cnt - a.cnt;
const byLabel = (a, b) => a.label.localeCompare(b.label, 'nl');

function buildFilters() {
  makeMS('pg', { presets: PG_PRESETS, items: () => countBy((r) => [[r.pg, `${r.pgNaam}${r.pg ? ' (' + r.pg + ')' : ''}`]]).sort(byCnt) });
  makeMS('prod', { items: () => countBy((r) => [[String(r.prod), `${r.prodNaam} (${r.prod})`]]).sort(byCnt) });
  makeMS('ntr', { items: () => countBy((r) => [[String(r.ntr), ntrLabel(r)]]).sort((a, b) => num(a.v) - num(b.v)) });
  makeMS('fout', { items: () => countBy((r) => (r.fout ? [[String(r.fout), `${r.fout} – ${r.foutNaam}`]] : [])).sort((a, b) => num(a.v) - num(b.v)) });
  makeMS('klant', { limit: 200, items: () => countBy((r) => [[r.klant, `${r.klantNaam || t('nameUnknown')} – ${FIRMA[r.firma] ? r.firma + ' ' : ''}${r.klnr}`]]).sort(byCnt) });
  makeMS('actie', { items: () => countBy((r) => [[String(r.actie), actieNaam(r.actie)]]).sort((a, b) => num(a.v) - num(b.v)) });
  makeMS('status', { items: () => countBy((r) => [[r.fase, `${r.faseNaam} (${r.fase})`]]).sort(byCnt) });
  makeMS('vert', { items: () => countBy((r) => [[String(r.vert), `${r.vertNaam || 'VT'} (${r.vert})`]]).sort(byLabel) });
  makeMS('regio', { items: () => countBy((r) => [[String(r.regio), `${r.regio} ${r.regioNaam}`.trim()]]).sort((a, b) => num(a.v) - num(b.v)) });
  makeMS('tree', {
    limit: 600,
    items: () => {
      // hiërarchisch: niveau 1 > 2 > 3 > detail, gesorteerd als een boom
      const m = new Map();
      for (const r of DATA.rows) {
        const seen = new Set();
        for (const p of r.tree) for (const [i, n] of treeNodes(p).entries()) {
          if (seen.has(n)) continue; seen.add(n);
          const e = m.get(n); if (e) e.cnt++; else m.set(n, { v: n, label: n.split(' › ').pop(), lvl: i + 1, cnt: 1 });
        }
      }
      const items = [...m.values()].sort((a, b) => a.v.localeCompare(b.v, 'nl'));
      return items.map((i) => ({ ...i, search: i.v.toLowerCase() }));
    },
  });
  // labels van de tree-opties: volledige pad bij zoeken is duidelijker
  MS.tree.labelOf = (v) => v;
}

/* ───────────── Filteren ───────────── */
function matches(r, skip) {
  if (skip !== 'period' && F.from != null && r.d < F.from) return false;
  if (skip !== 'period' && F.to != null && r.d > F.to) return false;
  if (skip !== 'pg' && F.pg.size && !F.pg.has(r.pg)) return false;
  if (skip !== 'prod' && F.prod.size && !F.prod.has(String(r.prod))) return false;
  if (skip !== 'ntr' && F.ntr.size && !F.ntr.has(String(r.ntr))) return false;
  if (skip !== 'fout' && F.fout.size && !F.fout.has(String(r.fout))) return false;
  if (skip !== 'klant' && F.klant.size && !F.klant.has(r.klant)) return false;
  if (skip !== 'status' && F.status.size && !F.status.has(r.fase)) return false;
  if (skip !== 'vert' && F.vert.size && !F.vert.has(String(r.vert))) return false;
  if (skip !== 'regio' && F.regio.size && !F.regio.has(String(r.regio))) return false;
  if (F.info === 'excl' && r.info) return false;
  if (skip !== 'actie' && F.actie.size && !F.actie.has(String(r.actie))) return false;
  if (F.vervolg === 'first' && (r.vervolg || r.andereOpOrder)) return false;
  if (F.vervolg === 'vervolg' && !r.vervolg) return false;
  if (F.vervolg === 'ander' && !r.andereOpOrder) return false;
  if (F.info === 'only' && !r.info) return false;
  if (F.open === 'open' && !r.open) return false;
  if (F.open === 'closed' && r.open) return false;
  if (skip !== 'tree' && F.tree.size) {
    let ok = false;
    outer: for (const p of r.tree) for (const n of treeNodes(p)) if (F.tree.has(n)) { ok = true; break outer; }
    if (!ok) return false;
  }
  if (F.text) {
    const t = F.textLc;
    // Ordernummer in eender welke notatie (67-02-0201-6, 67.02.0201.6, 670202016, ook een deel): zoekt in order en NTR-order.
    const od = F.textOrd;
    if (od && (String(r.order || '').includes(od) || String(r.ntrOrder || '').includes(od))) return true;
    if (!(r.oms.toLowerCase().includes(t) || r.opl.toLowerCase().includes(t) || r.oorz.some((o) => (o.oorzaak + ' ' + o.corr + ' ' + o.corrA).toLowerCase().includes(t)))) return false;
  }
  return true;
}

function update() {
  if (!DATA) return;
  filtered = DATA.rows.filter((r) => matches(r));
  renderChips();
  activeReport().render();
}

/* ───────────── Rapporten ─────────────
 * Elk rapport = één tabblad onder de filters. De filters gelden voor alle rapporten.
 * Een nieuw rapport toevoegen: een <div class="rep" id="rep_<id>"> in index.html,
 * een regel hieronder (id, titel-sleutel in i18n.js, render-functie, optioneel pdf-functie in report.js). */
const REPORTS = [
  { id: 'klachten', label: 'repKlachten', render: () => { renderKpi(); renderTrend(); renderTops(); renderTable(); } },
  { id: 'fin', label: 'repFin', render: () => renderFin() },
];
try { const r = localStorage.getItem('kl_report'); if (REPORTS.some((x) => x.id === r)) UI.report = r; } catch { /* geen opslag */ }
function activeReport() { return REPORTS.find((r) => r.id === UI.report) || REPORTS[0]; }
function renderTabs() {
  const cur = activeReport().id;
  $('repTabs').innerHTML = REPORTS.map((r) => `<button type="button" role="tab" data-rep="${r.id}" class="${r.id === cur ? 'on' : ''}" aria-selected="${r.id === cur}">${esc(t(r.label))}</button>`).join('');
  for (const r of REPORTS) $('rep_' + r.id).hidden = r.id !== cur;
}
$('repTabs').addEventListener('click', (e) => {
  const id = e.target.dataset && e.target.dataset.rep; if (!id || id === UI.report) return;
  UI.report = id; try { localStorage.setItem('kl_report', id); } catch { /* */ }
  renderTabs(); update();
});

function renderChips() {
  const c = []; const add = (label, clear, tip, key) => c.push({ label, clear, tip, key });
  if (F.from != null || F.to != null) add(`${t('chipPeriod')} ${F.from != null ? fmtDate(F.from) : '…'} – ${F.to != null ? fmtDate(F.to) : '…'}`, null);
  const names = { pg: t('lbl_pg'), prod: t('lbl_prod'), tree: 'TREE', ntr: 'NTR', fout: t('chipFout'), klant: t('lbl_klant'), status: t('lbl_status'), vert: t('chipVt'), regio: t('lbl_regio'), actie: t('chipActie') };
  for (const k of Object.keys(names)) {
    if (!F[k].size || !MS[k]) continue;
    const st = MS[k], sel = [...F[k]], lab = (v) => st.labelOf(v);
    const missing = st.items.filter((i) => !F[k].has(i.v)).map((i) => i.v);
    let txt;
    const preset = st.opts.presets && Object.entries(st.opts.presets).find(([, c]) => c.length && c.filter((x) => st.items.some((i) => i.v === x)).length === sel.length && c.every((x) => F[k].has(x) || !st.items.some((i) => i.v === x)));
    if (preset) txt = preset[0];
    else if (!missing.length) txt = t('chipAllOf');
    else if (sel.length <= 2) txt = sel.map(lab).join(', ');
    else if (missing.length <= 2) txt = t('chipExcept', { list: missing.map(lab).join(', ') });
    else txt = t('chipNofM', { n: sel.length, m: st.total });
    const tip = missing.length && missing.length < sel.length ? t('chipExceptTip') + '\n' + missing.map(lab).join('\n') : sel.map(lab).join('\n');
    add(`${names[k]}: ${txt}`, () => { F[k].clear(); MS[k].syncBtn(); }, tip, k);
  }
  if (F.open) add(F.open === 'open' ? t('onlyOpen') : t('onlyClosed'), () => { F.open = ''; $('fOpen').value = ''; });
  if (F.vervolg) add(t({ first: 'vvFirst', vervolg: 'vvOnly', ander: 'vvAnder' }[F.vervolg]), () => { F.vervolg = ''; $('fVervolg').value = ''; });
  if (F.info !== 'excl') add(F.info === 'only' ? t('infoOnly') : t('chipInfoIncl'), () => { F.info = 'excl'; $('fInfo').value = 'excl'; });
  if (F.text) add(t('chipSearch', { q: F.text }), () => { F.text = ''; $('fText').value = ''; });
  CHIPS = c.map((x) => ({ label: x.label, tip: x.tip }));
  const el = $('chips');
  el.innerHTML = c.map((x, i) => `<span class="chip${x.key ? ' link' : ''}"${x.tip ? ` title="${esc(x.tip)}"` : ''}${x.key ? ` data-k="${x.key}"` : ''}>${esc(x.label)}${x.clear ? `<button type="button" data-i="${i}" title="${esc(t('remove'))}">×</button>` : ''}</span>`).join('') +
    (c.length > 1 ? `<button type="button" class="btn-outline clear" id="clearAll">${esc(t('clearAll'))}</button>` : '');
  el.querySelectorAll('button[data-i]').forEach((b) => b.addEventListener('click', (e) => { e.stopPropagation(); c[+b.dataset.i].clear(); UI.page = 0; update(); }));
  el.querySelectorAll('.chip[data-k]').forEach((ch) => ch.addEventListener('click', (e) => {
    e.stopPropagation(); const h = MS[ch.dataset.k].host;
    h.scrollIntoView({ block: 'center', behavior: 'smooth' });
    if (!h.classList.contains('open')) h.querySelector('.ms-btn').click();
  }));
  const ca = $('clearAll');
  if (ca) ca.addEventListener('click', () => {
    for (const k of Object.keys(names)) { F[k].clear(); MS[k].syncBtn(); }
    F.open = ''; F.info = 'excl'; $('fInfo').value = 'excl'; F.vervolg = ''; $('fVervolg').value = ''; F.text = ''; $('fOpen').value = ''; $('fText').value = ''; UI.page = 0; update();
  });
}

function renderKpi() {
  const n = filtered.length;
  let open = 0, cost = 0, lead = 0, leadN = 0, withTree = 0;
  for (const r of filtered) {
    if (r.open) open++;
    cost += r.cost;
    if (!r.open && r.dA != null && r.dA >= r.d) { lead += r.dA - r.d; leadN++; }
    if (r.tree.length) withTree++;
  }
  $('kTotal').textContent = nf.format(n);
  $('kOpen').textContent = nf.format(open);
  $('kCost').textContent = eur.format(cost);
  $('kLead').textContent = leadN ? nf.format(Math.round(lead / leadN)) : '–';
  $('kTree').textContent = n ? Math.round((withTree / n) * 100) + '%' : '–';
  const dos = new Set(); let iv = 0, vv = 0, ao = 0;
  for (const r of filtered) { dos.add(r.dos); if (r.interv) iv++; if (r.vervolg) vv++; if (r.andereOpOrder) ao++; }
  const pc = (x) => (n ? ` <small>${Math.round((x / n) * 100)}%</small>` : '');
  $('kDos').textContent = nf.format(dos.size);
  $('kInterv').innerHTML = nf.format(iv) + pc(iv);
  $('kVervolg').innerHTML = nf.format(vv) + pc(vv);
  $('kAnder').innerHTML = nf.format(ao) + pc(ao);
}

/* ───────────── Evolutie-grafiek ───────────── */
function periodKey(dn) {
  const d = new Date(dn * 864e5); const y = d.getUTCFullYear(), m = d.getUTCMonth() + 1;
  if (UI.gran === 'Y') return String(y);
  if (UI.gran === 'Q') return `${y}-K${Math.ceil(m / 3)}`;
  return `${y}-${String(m).padStart(2, '0')}`;
}
function periodRange() {
  // alle periodes tussen begin en einde, ook lege
  let a = F.from, b = F.to;
  if (a == null || b == null) {
    let mn = Infinity, mx = -Infinity; for (const r of filtered) { if (r.d < mn) mn = r.d; if (r.d > mx) mx = r.d; }
    if (a == null) a = mn; if (b == null) b = mx;
  }
  if (!Number.isFinite(a) || !Number.isFinite(b)) return [];
  const keys = []; const d = new Date(a * 864e5); d.setUTCDate(1);
  const end = new Date(b * 864e5);
  while (d <= end) { const k = periodKey(dayNum(new Date(d.getUTCFullYear(), d.getUTCMonth(), 1))); if (keys[keys.length - 1] !== k) keys.push(k); d.setUTCMonth(d.getUTCMonth() + 1); }
  return keys;
}
function splitKeys(r) {
  switch (UI.split) {
    case 'pg': return [r.pgNaam];
    case 'prod': return [r.prodNaam];
    case 'ntr': return [ntrLabel(r)];
    case 'klant': return [r.klantNaam || r.klant];
    case 'status': return [r.open ? t('open') : t('closed')];
    case 'actie': return [actieNaam(r.actie)];
    case 'vervolg': return [r.vervolg ? t('sVervolg') : r.andereOpOrder ? t('sAnder') : t('sFirst')];
    case 't1': case 't2': case 't3': {
      const lv = +UI.split[1];
      const s = new Set(r.tree.map((p) => treeNodes(p)[lv - 1]).filter(Boolean));
      return [...s]; // klachten zonder TREE tellen niet mee in deze opsplitsing
    }
    default: return [t('total')];
  }
}
function renderTrend() {
  if (UI.mode === 'cumul' || UI.mode === 'abs') return renderCumul(UI.mode === 'abs');
  if (UI.mode === 'qual') return renderQual();
  const keys = periodRange();
  const pos = new Map(keys.map((k, i) => [k, i]));
  const val = (r) => (UI.measure === 'cost' ? r.cost : 1);
  const series = new Map();
  for (const r of filtered) {
    const i = pos.get(periodKey(r.d)); if (i == null) continue;
    for (const s of splitKeys(r)) { let a = series.get(s); if (!a) series.set(s, (a = new Array(keys.length).fill(0))); a[i] += val(r); }
  }
  // max 7 reeksen + "Overig"
  let list = [...series.entries()].map(([name, data]) => ({ name, data, tot: data.reduce((x, y) => x + y, 0) })).sort((a, b) => b.tot - a.tot);
  if (list.length > 8) {
    const rest = list.slice(7); const other = new Array(keys.length).fill(0);
    for (const s of rest) s.data.forEach((v, i) => { other[i] += v; });
    list = [...list.slice(0, 7), { name: t('other', { n: rest.length }), data: other, other: true }];
  }
  const ds = list.map((s, i) => ({
    label: s.name, data: s.data,
    borderColor: s.other ? OTHER_COLOR : COLORS[i % COLORS.length],
    backgroundColor: s.other ? OTHER_COLOR : COLORS[i % COLORS.length],
    borderWidth: 2, pointRadius: keys.length > 40 ? 0 : 3, pointHoverRadius: 5, tension: 0.25,
    borderDash: s.other ? [4, 3] : undefined,
  }));
  const isCost = UI.measure === 'cost';
  const cfg = {
    type: 'line',
    data: { labels: keys.map(prettyPeriod), datasets: ds },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#111', padding: 10, boxPadding: 4, itemSort: (a, b) => b.raw - a.raw,
          callbacks: { label: (c) => ` ${c.dataset.label}: ${isCost ? eur.format(c.raw) : nf.format(c.raw)}` },
        },
      },
      scales: {
        x: { grid: { display: false }, ticks: { color: '#6b7280', maxRotation: 0, autoSkip: true, maxTicksLimit: 14 } },
        y: { beginAtZero: true, grid: { color: '#eef0f2' }, border: { display: false }, ticks: { color: '#6b7280', precision: 0, callback: (v) => (isCost ? eur.format(v) : nf.format(v)) } },
      },
    },
  };
  if (chart) chart.destroy();
  chart = new Chart($('trend'), cfg);
  const maxD = DATA.rows.reduce((m, r) => Math.max(m, r.d), 0);
  const lastKey = keys[keys.length - 1];
  const partial = lastKey && periodKey(maxD + 1) === lastKey;
  const notes = [];
  if (UI.split.startsWith('t')) notes.push(t('noteTreeSplit'));
  if (partial) notes.push(t('notePartial', { p: prettyPeriod(lastKey) }));
  $('legend').innerHTML = (ds.length > 1 ? ds.map((d) => `<span><i style="background:${d.borderColor}"></i>${esc(d.label)}</span>`).join('') : '') +
    (notes.length ? `<div style="flex-basis:100%;color:#6b7280">${esc(notes.join(' '))}</div>` : '');
}
/* Jaarvergelijking: per jaar cumulatief van januari tot december.
 * Alle filters gelden, behalve de periode: elk jaar in de data krijgt een lijn. */
const MONTHS = () => t('months').split(',');
function yearSlot(dn) {
  const d = new Date(dn * 864e5); const y = d.getUTCFullYear();
  if (UI.cumGran === 'M') return [y, d.getUTCMonth()];
  const doy = Math.floor((Date.UTC(y, d.getUTCMonth(), d.getUTCDate()) - Date.UTC(y, 0, 1)) / 864e5);
  return [y, Math.min(52, Math.floor(doy / 7))]; // week 1..53 (dag 1-7 = week 1)
}
/* Interventie & vervolg: per periode het aandeel klachten met interventie en het aandeel vervolgklachten. */
function renderQual() {
  const keys = periodRange();
  const pos = new Map(keys.map((k, i) => [k, i]));
  const n = new Array(keys.length).fill(0), iv = new Array(keys.length).fill(0), vv = new Array(keys.length).fill(0), ao = new Array(keys.length).fill(0);
  for (const r of filtered) { const i = pos.get(periodKey(r.d)); if (i == null) continue; n[i]++; if (r.interv) iv[i]++; if (r.vervolg) vv[i]++; if (r.andereOpOrder) ao[i]++; }
  const pct = (a) => a.map((v, i) => (n[i] ? Math.round((v / n[i]) * 1000) / 10 : null));
  const ds = [
    { label: t('qVervolg'), data: pct(vv), cnt: vv, borderColor: COLORS[0], backgroundColor: COLORS[0], borderWidth: 2.5, pointRadius: keys.length > 40 ? 0 : 3, tension: 0.25 },
    { label: t('qInterv'), data: pct(iv), cnt: iv, borderColor: COLORS[1], backgroundColor: COLORS[1], borderWidth: 2, pointRadius: keys.length > 40 ? 0 : 3, tension: 0.25 },
    { label: t('qAnder'), data: pct(ao), cnt: ao, borderColor: OTHER_COLOR, backgroundColor: OTHER_COLOR, borderWidth: 2, borderDash: [5, 4], pointRadius: keys.length > 40 ? 0 : 3, tension: 0.25 },
  ];
  if (chart) chart.destroy();
  chart = new Chart($('trend'), {
    type: 'line',
    data: { labels: keys.map(prettyPeriod), datasets: ds },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: { backgroundColor: '#111', padding: 10, boxPadding: 4, callbacks: {
          label: (c) => ` ${c.dataset.label}: ${c.raw == null ? '–' : nf.format(c.raw) + '%'} (${t('xOfY', { x: nf.format(c.dataset.cnt[c.dataIndex]), y: nf.format(n[c.dataIndex]) })})`,
        } },
      },
      scales: {
        x: { grid: { display: false }, ticks: { color: '#6b7280', maxRotation: 0, autoSkip: true, maxTicksLimit: 14 } },
        y: { beginAtZero: true, grid: { color: '#eef0f2' }, border: { display: false }, title: { display: true, text: t('qAxis'), color: '#6b7280' }, ticks: { color: '#6b7280', callback: (v) => v + '%' } },
      },
    },
  });
  const maxD = DATA.rows.reduce((m, r) => Math.max(m, r.d), 0);
  const notes = [t('noteQual')];
  if (keys.length && periodKey(maxD + 1) === keys[keys.length - 1]) notes.push(t('notePartialQual', { p: prettyPeriod(keys[keys.length - 1]) }));
  $('legend').innerHTML = ds.map((d) => `<span><i style="background:${d.borderColor}"></i>${esc(d.label)}</span>`).join('') + `<div style="flex-basis:100%;color:#6b7280">${esc(notes.join(' '))}</div>`;
}

// Jaren cumulatief (opgeteld) of Jaren absoluut (abs = per week/maand, niet opgeteld): één lijn per jaar, jan–dec.
function renderCumul(abs) {
  const rows = DATA.rows.filter((r) => matches(r, 'period'));
  const nSlots = UI.cumGran === 'M' ? 12 : 53;
  const labels = UI.cumGran === 'M' ? MONTHS() : Array.from({ length: 53 }, (_, i) => String(i + 1));
  const val = (r) => (UI.measure === 'cost' ? r.cost : 1);
  const per = new Map();
  for (const r of rows) {
    const [y, i] = yearSlot(r.d);
    let a = per.get(y); if (!a) per.set(y, (a = new Array(nSlots).fill(0)));
    a[i] += val(r);
  }
  const maxD = DATA.rows.reduce((m, r) => Math.max(m, r.d), 0);
  const [curY, curI] = yearSlot(maxD);
  const years = [...per.keys()].sort((a, b) => a - b);
  const cum = (a, upto) => { let t = 0; return a.map((v, i) => (i > upto ? null : (t += v))); };
  const plain = (a, upto) => a.map((v, i) => (i > upto ? null : v));
  const series = years.map((y) => ({ y, data: (abs ? plain : cum)(per.get(y), y === curY ? curI : nSlots - 1) }));
  // kleuren volgen het jaar (niet de volgorde in de selectie), zodat een jaar altijd dezelfde kleur houdt
  const colorOf = (y) => (y === curY ? '#111111' : COLORS[((y - 2020) % COLORS.length + COLORS.length) % COLORS.length]);
  const ds = series.map((s) => ({
    key: String(s.y), label: String(s.y), data: s.data, borderColor: colorOf(s.y), backgroundColor: colorOf(s.y),
    borderWidth: s.y === curY ? 3 : 2, pointRadius: 0, pointHoverRadius: 4, tension: 0.2, spanGaps: false,
  }));
  const notes = [t('noteCumul')]; if (abs) notes.push(t('noteAbs'));
  if (!abs && UI.target && per.has(curY - 1)) {
    const f = 1 + Number(UI.target) / 100;
    const prev = cum(per.get(curY - 1), nSlots - 1).map((v) => (v == null ? null : Math.round(v * f * 100) / 100));
    ds.push({ key: 'doel', label: t('targetLbl', { y: curY, p: UI.target, py: curY - 1 }), data: prev, borderColor: '#e34948', backgroundColor: '#e34948', borderWidth: 2, borderDash: [6, 4], pointRadius: 0, pointHoverRadius: 4, tension: 0.2 });
  } else if (!abs && UI.target) notes.push(t('noTargetData', { y: curY - 1 }));
  const isCost = UI.measure === 'cost';
  if (chart) chart.destroy();
  chart = new Chart($('trend'), {
    type: 'line',
    data: { labels, datasets: ds.filter((d) => !HIDDEN_YEARS.has(d.key)) },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#111', padding: 10, boxPadding: 4, filter: (c) => c.raw != null,
          callbacks: {
            title: (c) => (abs ? (UI.cumGran === 'M' ? c[0].label : t('week') + ' ' + c[0].label) : UI.cumGran === 'M' ? t('uptoMonth', { m: c[0].label }) : t('uptoWeek', { w: c[0].label })),
            label: (c) => ` ${c.dataset.label}: ${isCost ? eur.format(c.raw) : nf.format(c.raw)}`,
          },
        },
      },
      scales: {
        x: { grid: { display: false }, title: { display: true, text: UI.cumGran === 'M' ? t('month') : t('week'), color: '#6b7280' }, ticks: { color: '#6b7280', maxRotation: 0, autoSkip: true, maxTicksLimit: 14 } },
        y: { beginAtZero: true, grid: { color: '#eef0f2' }, border: { display: false }, title: { display: true, text: abs ? (isCost ? t('ntrCost') : t('absCount')) : isCost ? t('cumCost') : t('cumCount'), color: '#6b7280' }, ticks: { color: '#6b7280', precision: 0, callback: (v) => (isCost ? eur.format(v) : nf.format(v)) } },
      },
    },
  });
  // Vinkjes per jaar: kies welke jaren getoond worden (onthouden in deze browser)
  $('legend').innerHTML = `<span class="yr-lbl">${esc(t('show'))}</span>` + ds.map((d) => `<label class="yr${HIDDEN_YEARS.has(d.key) ? ' off' : ''}"><input type="checkbox" data-yr="${esc(d.key)}"${HIDDEN_YEARS.has(d.key) ? '' : ' checked'}><i style="background:${d.borderColor}${d.borderDash ? ';background:repeating-linear-gradient(90deg,#e34948 0 5px,transparent 5px 8px)' : ''}"></i>${esc(d.label)}</label>`).join('') +
    `<button type="button" class="yr-q" data-q="all">${esc(t('optAll'))}</button><button type="button" class="yr-q" data-q="3">${esc(t('last3'))}</button>` +
    `<div style="flex-basis:100%;color:#6b7280">${esc(notes.join(' '))}</div>`;
  $('legend').querySelectorAll('input[data-yr]').forEach((cb) => cb.addEventListener('change', () => {
    if (cb.checked) HIDDEN_YEARS.delete(cb.dataset.yr); else HIDDEN_YEARS.add(cb.dataset.yr);
    saveHidden(); renderTrend();
  }));
  $('legend').querySelectorAll('.yr-q').forEach((b) => b.addEventListener('click', () => {
    HIDDEN_YEARS = new Set();
    if (b.dataset.q === '3') years.slice(0, -3).forEach((y) => HIDDEN_YEARS.add(String(y)));
    saveHidden(); renderTrend();
  }));
}

function prettyPeriod(k) {
  if (UI.gran === 'M') { const [y, m] = k.split('-'); return MONTHS()[+m - 1] + ' ' + y.slice(2); }
  if (UI.gran === 'Q') return k.replace('-K', '-' + t('qShort'));
  return k;
}

/* ───────────── Top-lijsten ───────────── */
function renderBars(el, entries, onClick, total) {
  if (!entries.length) { el.innerHTML = `<div class="empty">${esc(t('noData'))}</div>`; return; }
  const max = entries[0].n;
  el.innerHTML = entries.map((e, i) => `<div class="bar-row" data-i="${i}" title="${esc(e.label)}">
      <div class="bar-track"><div class="bar-fill" style="width:${Math.max(2, (e.n / max) * 100)}%"></div><div class="bar-lbl">${esc(e.label)}</div></div>
      <div class="bar-val">${nf.format(e.n)}${total ? ` <span style="color:#9ca3af">${Math.round((e.n / total) * 100)}%</span>` : ''}</div></div>`).join('');
  el.querySelectorAll('.bar-row').forEach((row) => row.addEventListener('click', () => onClick(entries[+row.dataset.i])));
}
function topBy(fn, n = 10) {
  const m = new Map();
  for (const r of filtered) for (const k of fn(r)) { if (!k) continue; const e = m.get(k[0]); if (e) e.n++; else m.set(k[0], { key: k[0], label: k[1], n: 1 }); }
  return [...m.values()].sort((a, b) => b.n - a.n).slice(0, n);
}
// Klik op een rij/balk = inzoomen: de filter wordt vervangen door enkel die waarde
// (toevoegen deed niets als de waarde al in een gekozen groep zat, bv. preset Winsol).
function addFilter(key, v) { F[key].clear(); F[key].add(String(v)); MS[key].syncBtn(); UI.page = 0; update(); }
function renderTops() {
  const total = filtered.length;
  const lv = UI.treeLevel;
  const treeRows = topBy((r) => { const s = new Set(r.tree.map((p) => treeNodes(p)[Math.min(lv, treeNodes(p).length) - 1])); return [...s].map((n) => [n, n]); }, 12);
  renderBars($('topTree'), treeRows, (e) => addFilter('tree', e.key), total);
  renderBars($('topProd'), topBy((r) => [[String(r.prod), r.prodNaam]]), (e) => addFilter('prod', e.key), total);
  renderBars($('topNtr'), topBy((r) => [[String(r.ntr), ntrLabel(r)]]), (e) => addFilter('ntr', e.key), total);
  renderBars($('topKlant'), topBy((r) => [[r.klant, `${r.klantNaam || t('nameUnknown')} (${r.klnr})`]]), (e) => addFilter('klant', e.key), total);
}

/* ───────────── Tabel ───────────── */
const COLS = [
  { k: 'id', t: 'colId', f: (r) => r.id },
  { k: 'd', t: 'colDate', f: (r) => fmtDate(r.d), s: (r) => r.d },
  { k: 'klant', t: 'lbl_klant', f: (r) => r.klantNaam || r.klant, s: (r) => r.klantNaam, clip: 1 },
  { k: 'prod', t: 'lbl_prod', f: (r) => r.prodNaam, clip: 1 },
  { k: 'ntr', t: 'colNtr', f: (r) => (r.ntr ? `${r.ntr} – ${r.ntrNaam}` : '–'), s: (r) => r.ntr, clip: 1 },
  { k: 'tree', t: 'colTree', f: (r) => (r.tree[0] ? r.tree[0].split('/').slice(1).join(' › ') : '') + (r.tree.length > 1 ? ` (+${r.tree.length - 1})` : ''), s: (r) => r.tree[0] || '', clip: 1 },
  { k: 'oms', t: 'colOms', f: (r) => r.oms.replace(/\s+/g, ' '), clip: 1 },
  { k: 'open', t: 'lbl_status', html: (r) => `<span class="pill ${r.open ? 'open' : 'closed'}" title="${esc(r.faseNaam)}">${esc(r.open ? t('open') : t('closed'))}</span>`, s: (r) => (r.open ? 1 : 0) },
  { k: 'cost', t: 'ntrCost', f: (r) => (r.cost ? eur.format(r.cost) : ''), s: (r) => r.cost, num: 1 },
];
const PAGE = 50;
function renderTable() {
  const col = COLS.find((c) => c.k === UI.sort.key) || COLS[1];
  const sv = col.s || col.f;
  const rows = filtered.slice().sort((a, b) => { const x = sv(a), y = sv(b); return (x < y ? -1 : x > y ? 1 : 0) * UI.sort.dir; });
  const pages = Math.max(1, Math.ceil(rows.length / PAGE));
  UI.page = Math.min(UI.page, pages - 1);
  const slice = rows.slice(UI.page * PAGE, UI.page * PAGE + PAGE);
  $('tbl').innerHTML = `<thead><tr>${COLS.map((c) => `<th data-k="${c.k}" class="${c.num ? 'num' : ''}">${esc(t(c.t))} <span class="arr">${UI.sort.key === c.k ? (UI.sort.dir > 0 ? '▲' : '▼') : ''}</span></th>`).join('')}</tr></thead>
    <tbody>${slice.map((r) => `<tr data-id="${r.id}">${COLS.map((c) => `<td class="${c.clip ? 'clip' : ''} ${c.num ? 'num' : ''}">${c.html ? c.html(r) : esc(c.f(r))}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${COLS.length}" class="empty">${esc(t('noRows'))}</td></tr>`}</tbody>`;
  $('tbl').querySelectorAll('th').forEach((th) => th.addEventListener('click', () => {
    const k = th.dataset.k; UI.sort = { key: k, dir: UI.sort.key === k ? -UI.sort.dir : (k === 'd' || k === 'cost' ? -1 : 1) }; renderTable();
  }));
  $('tbl').querySelectorAll('tbody tr[data-id]').forEach((tr) => tr.addEventListener('click', () => showDetail(+tr.dataset.id)));
  $('pager').innerHTML = `${esc(t('pager', { n: nf.format(rows.length), p: UI.page + 1, pp: pages }))}
    <button class="btn-outline" id="pPrev" ${UI.page === 0 ? 'disabled' : ''}>‹ ${esc(t('prev'))}</button>
    <button class="btn-outline" id="pNext" ${UI.page >= pages - 1 ? 'disabled' : ''}>${esc(t('next'))} ›</button>`;
  $('pPrev').onclick = () => { UI.page--; renderTable(); };
  $('pNext').onclick = () => { UI.page++; renderTable(); };
}

/* ───────────── Detail ───────────── */
function showDetail(id) {
  const r = DATA.rows.find((x) => x.id === id); if (!r) return;
  const kv = (l, v) => `<div><span>${l}</span><b>${esc(v || '–')}</b></div>`;
  $('modal').innerHTML = `
    <div style="display:flex;gap:10px;align-items:flex-start">
      <div style="flex:1"><h2>${esc(t('complaintN', { n: r.id }))}</h2>
      <div class="sub">${esc(r.klantNaam || '')} (${esc(FIRMA[r.firma] || r.firma)} ${r.klnr})${r.land ? ' · ' + esc(r.land) : ''} · ${fmtDate(r.d)} · <span class="pill ${r.open ? 'open' : 'closed'}">${esc(r.faseNaam)}</span></div></div>
      <a class="btn-outline" href="${WINPRO_KLACHT}${r.id}" target="_blank" rel="noopener">${esc(t('openWinpro'))} ↗</a>
      <button class="btn-outline" id="mClose" type="button">${esc(t('close'))}</button>
    </div>
    <div class="dl">
      ${kv(t('lbl_pg'), r.pgNaam)}${kv(t('lbl_prod'), `${r.prodNaam} (${r.prod})`)}${r.aVia ? kv(t('aVia'), t('aVia_' + r.aVia)) + (r.aVia !== 'product' ? kv(t('aEntered'), `${r.prodKNaam} (${r.prodK}) · ${r.pgKNaam}`) : '') : ''}${kv(t('order'), fmtOrder(r.order) || '')}
      ${r.oD != null ? kv(t('oDate'), fmtDate(r.oD)) + kv(t('oLev'), fmtDate(r.oDU || r.oDL)) + kv(t('oAge'), (r.oDU || r.oDL) && r.d >= (r.oDU || r.oDL) ? t('nDays', { n: r.d - (r.oDU || r.oDL) }) : '') + kv(t('oFakt'), r.oFakt ? eur.format(r.oFakt) : '') + kv(t('oPlk'), r.oPlk ? eur.format(r.oPlk) : '') + kv(t('oLp'), r.oLp === 2 ? t('actie2') : r.oLp === 1 ? t('actie1') : '') : ''}
      ${kv('NTR' + (r.ntr && !r.ntrBevestigd ? ' (' + t('presumed') + ')' : ''), r.ntr ? `${r.ntr} – ${r.ntrNaam}` : t('missing'))}${kv(t('lbl_fout'), r.fout ? `${r.fout} – ${r.foutNaam}` : '')}${kv(t('ntrOrder'), fmtOrder(r.ntrOrder) || '')}
      ${kv(t('costMat'), r.costMat ? eur.format(r.costMat) : '')}${kv(t('costPl'), r.costPl ? eur.format(r.costPl) : '')}${kv(t('handledBy'), r.beh)}${kv(t('lbl_actie'), actieNaam(r.actie))}${kv(t('iFlag'), r.intervFlag === '1' || r.intervFlag === 'I' ? t('yes') + ' (' + r.intervFlag + ')' : t('no'))}${kv(t('sVervolg'), r.vervolg ? t('vvYes', { d: r.vgap, how: vlinkLabel(r.vlink) }) : r.andereOpOrder ? t('vvNoOther') : t('no'))}
      ${kv(t('lbl_vert'), r.vertNaam ? `${r.vertNaam} (${r.vert})` : r.vert)}${kv(t('lbl_regio'), `${r.regio} ${r.regioNaam}`)}${kv(t('leadTime'), !r.open && r.dA >= r.d ? t('nDays', { n: r.dA - r.d }) : '')}
      ${kv(t('dB'), fmtDate(r.dB))}${kv(t('dM'), fmtDate(r.dM))}${kv(t('dD'), fmtDate(r.dD))}
      ${kv(t('dK'), fmtDate(r.dK))}${kv(t('dA'), fmtDate(r.dA))}
    </div>
    ${dossierHtml(r)}
    <h4>TREE</h4><div>${r.tree.length ? r.tree.map((p) => `<span class="tree-path">${esc(p.split('/').join(' › '))}</span>`).join('') : `<span class="empty">${esc(t('notFilled'))}</span>`}</div>
    <h4>${esc(t('omsTitle'))}</h4><div class="txt-block">${esc(r.oms) || '–'}</div>
    <h4>${esc(t('oplTitle'))}</h4><div class="txt-block">${esc(r.opl) || '–'}</div>
    ${r.oorz.map((o) => `<h4>${esc(t('cause'))}</h4><div class="txt-block">${esc(o.oorzaak) || '–'}</div><h4>${esc(t('corr'))}</h4><div class="txt-block">${esc(o.corr) || '–'}</div><h4>${esc(t('corrA'))}</h4><div class="txt-block">${esc(o.corrA) || '–'}</div>`).join('')}
  `;
  $('modalBg').classList.add('open');
  $('mClose').onclick = closeDetail;
  $('modal').querySelectorAll('.dos-item[data-id]').forEach((el) => el.addEventListener('click', () => showDetail(+el.dataset.id)));
}
const DOS_COLORS = ['#cfe3f7', '#f6d4e4', '#d4efdd', '#fde8b8', '#e3dcf7', '#f9d9c8'];
function dossierHtml(r) {
  const byId = new Map(DATA.rows.map((x) => [x.id, x]));
  const ids = r.rootOrder ? ORDERS.get(r.rootOrder) || [r.id] : DOSSIERS.get(r.dos) || [r.id];
  const list = ids.map((i) => byId.get(i)).filter(Boolean);
  if (list.length < 2) return `<h4>${esc(t('order'))} ${esc(fmtOrder(r.order) || '–')}</h4><div class="empty">${esc(t('noOtherOnOrder'))}</div>`;
  // groepeer per dossier (probleem), dossiers in volgorde van hun eerste klacht
  const groups = new Map();
  for (const x of list.sort((a, b) => a.d - b.d || a.id - b.id)) { let g = groups.get(x.dos); if (!g) groups.set(x.dos, (g = [])); g.push(x); }
  const cost = list.reduce((t, x) => t + x.cost, 0), nIv = list.filter((x) => x.interv).length;
  let gi = 0, html = '';
  for (const g of groups.values()) {
    const c = DOS_COLORS[gi % DOS_COLORS.length]; gi++;
    const gc = g.reduce((t, x) => t + x.cost, 0);
    html += `<div class="dos-group" style="border-left:6px solid ${c}"><div class="dos-head">${esc(t('problemN', { n: gi }))} · ${esc(t(g.length === 1 ? 'nComplaint1' : 'nComplaints', { n: g.length }))}${g.length > 1 ? ` (${esc(t('nFollow', { n: g.length - 1 }))})` : ''}${gc ? ' · ' + eur.format(gc) : ''}</div>` +
      g.map((x) => `<div class="dos-item${x.id === r.id ? ' cur' : ''}" data-id="${x.id}" title="${esc(t('complaintN', { n: x.id }))}"><span>${fmtDate(x.d)}</span><span>${x.id}</span><span class="act">${esc(actieNaam(x.actie))}</span><span class="txt">${x.vervolg ? '↳ ' : ''}${esc(x.oms.replace(/\s+/g, ' ') || x.faseNaam)}</span></div>`).join('') + '</div>';
  }
  const root = byId.get(r.dos) || r; const fakt = root.order === r.rootOrder ? root.oFakt : list.find((x) => x.order === r.rootOrder && x.oFakt)?.oFakt;
  return `<h4>${esc(t('orderSummary', { o: fmtOrder(r.rootOrder), n: list.length, p: groups.size, i: nIv, c: eur.format(cost) }))}${fakt ? esc(' · ' + t('oFaktPct', { f: eur.format(fakt), p: fpct((cost / fakt) * 100) })) : ''}</h4><div class="dos-list">${html}</div>`;
}
function closeDetail() { $('modalBg').classList.remove('open'); }
$('modalBg').addEventListener('click', (e) => { if (e.target === $('modalBg')) closeDetail(); });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeDetail(); });

/* ───────────── Export ───────────── */
function exportXlsx() {
  const X = [
    ['colId', (r) => r.id], ['colDate', (r) => fmtDate(r.d)], ['xFirma', (r) => r.firma], ['xKlnr', (r) => r.klnr], ['lbl_klant', (r) => r.klantNaam], ['xLand', (r) => r.land],
    ['lbl_pg', (r) => r.pgNaam], ['xProdnr', (r) => r.prod], ['lbl_prod', (r) => r.prodNaam], ['aVia', (r) => (r.aVia ? t('aVia_' + r.aVia) : '')], ['xPgK', (r) => r.pgKNaam ?? r.pgNaam], ['xProdK', (r) => r.prodKNaam ?? r.prodNaam], ['order', (r) => r.order || ''],
    ['colNtr', (r) => r.ntr || ''], ['xNtrOms', (r) => (r.ntr ? r.ntrNaam : t('ntrMissing'))], ['xNtrConf', (r) => yesNo(r.ntrBevestigd)], ['lbl_fout', (r) => r.fout || ''], ['xFout', (r) => r.foutNaam],
    ['TREE 1', (r) => r.tree[0] || ''], ['TREE 2', (r) => r.tree[1] || ''], ['TREE 3', (r) => r.tree[2] || ''],
    ['lbl_status', (r) => r.faseNaam], ['lbl_actie', (r) => actieNaam(r.actie)], ['xInterv', (r) => yesNo(r.interv)], ['xDossier', (r) => r.dos], ['xDosSize', (r) => r.dosSize],
    ['sVervolg', (r) => yesNo(r.vervolg)], ['xLink', (r) => vlinkLabel(r.vlink)], ['sAnder', (r) => yesNo(r.andereOpOrder)], ['xRootOrder', (r) => fmtOrder(r.rootOrder)], ['xGap', (r) => r.vgap ?? ''],
    ['open', (r) => yesNo(r.open)], ['dA', (r) => fmtDate(r.dA)], ['xLead', (r) => (!r.open && r.dA >= r.d ? r.dA - r.d : '')],
    ['costMat', (r) => r.costMat], ['costPl', (r) => r.costPl], ['lbl_vert', (r) => r.vertNaam], ['lbl_regio', (r) => `${r.regio} ${r.regioNaam}`.trim()], ['handledBy', (r) => r.beh],
    ['omsTitle', (r) => r.oms], ['oplTitle', (r) => r.opl], ['cause', (r) => r.oorz.map((o) => o.oorzaak).join(' | ')],
  ];
  const head = X.map(([k]) => (k.startsWith('TREE') ? k : t(k)));
  const ws = XLSX.utils.aoa_to_sheet([head, ...filtered.map((r) => X.map(([, f]) => f(r)))]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, t('sheetName'));
  XLSX.writeFile(wb, `${t('fileName')}_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* ───────────── Periode-snelkeuze ───────────── */
function quickPeriod(q) {
  const maxD = DATA.rows.reduce((m, r) => Math.max(m, r.d), 0);
  const minD = DATA.rows.reduce((m, r) => Math.min(m, r.d), Infinity);
  const end = new Date(maxD * 864e5);
  if (q === 'all') { F.from = minD; F.to = maxD; }
  else if (q === 'ytd') { F.from = dayNum(new Date(end.getUTCFullYear(), 0, 1)); F.to = maxD; }
  else { const m = +q; F.from = dayNum(new Date(end.getUTCFullYear(), end.getUTCMonth() - m + 1, 1)); F.to = maxD; }
  $('dFrom').value = isoFromDay(F.from); $('dTo').value = isoFromDay(F.to);
  document.querySelectorAll('#quick button').forEach((b) => b.classList.toggle('on', b.dataset.q === q));
}

/* ───────────── Events ───────────── */
$('reloadBtn').addEventListener('click', () => loadFromServer());
wireDrop('dataDrop', 'dataFile', prepareUpload);
$('uploadBtn').addEventListener('click', doUpload);
wireDrop('usersDrop', 'usersFile', prepareUsers);
$('usersUpdateBtn').addEventListener('click', doUsers);
$('usersListBtn').addEventListener('click', listUsers);
$('loginSubmitBtn').addEventListener('click', doLogin);
$('loginPassword').addEventListener('keydown', (e) => { if (e.key === 'Enter') doLogin(); });
$('loginUsername').addEventListener('keydown', (e) => { if (e.key === 'Enter') $('loginPassword').focus(); });
$('loginPwToggle').addEventListener('click', () => { const p = $('loginPassword'); p.type = p.type === 'password' ? 'text' : 'password'; });
$('logoutBtn').addEventListener('click', async () => { try { await apiRequest({ mode: 'logout' }); } catch { /* */ } await DB.clear(); showLogin(); });
document.querySelectorAll('select.ui-lang').forEach((sel) => sel.addEventListener('change', () => {
  saveUiLangSetting(sel.value); uiLangOverride = sel.value;
  applyI18n(); document.querySelectorAll('select.ui-lang').forEach(fillUiLangSelect);
  setNumberLocale();
  if (typeof refreshInfoButtons === 'function') refreshInfoButtons();
  if (ME) { $('topbarRole').textContent = roleLabel(ME.role); renderDataInfo(); if (ME.role === 'admin') showLastUpload(); }
  if (DATA) { buildFilters(); renderTabs(); update(); }
}));
$('dFrom').addEventListener('change', (e) => { F.from = dayFromIso(e.target.value); document.querySelectorAll('#quick button').forEach((b) => b.classList.remove('on')); UI.page = 0; update(); });
$('dTo').addEventListener('change', (e) => { F.to = dayFromIso(e.target.value); document.querySelectorAll('#quick button').forEach((b) => b.classList.remove('on')); UI.page = 0; update(); });
$('quick').addEventListener('click', (e) => { const q = e.target.dataset.q; if (q) { quickPeriod(q); UI.page = 0; update(); } });
$('fOpen').addEventListener('change', (e) => { F.open = e.target.value; UI.page = 0; update(); });
$('fInfo').addEventListener('change', (e) => { F.info = e.target.value; UI.page = 0; update(); });
$('fVervolg').addEventListener('change', (e) => { F.vervolg = e.target.value; UI.page = 0; update(); });
let tmr; $('fText').addEventListener('input', (e) => { clearTimeout(tmr); tmr = setTimeout(() => { F.text = e.target.value.trim(); F.textLc = F.text.toLowerCase(); F.textOrd = /^[\d\s.\-\/]+$/.test(F.text) && F.text.replace(/\D/g, '').length >= 4 ? F.text.replace(/\D/g, '') : ''; UI.page = 0; update(); }, 250); });
const seg = (id, key) => $(id).addEventListener('click', (e) => { const v = e.target.dataset.v; if (!v) return; UI[key] = v; $(id).querySelectorAll('button').forEach((b) => b.classList.toggle('on', b.dataset.v === v)); renderTrend(); });
seg('segGran', 'gran'); seg('segMeasure', 'measure'); seg('segCumGran', 'cumGran');
function syncModeCtrls() { document.querySelectorAll('[data-modes]').forEach((el) => { el.hidden = !el.dataset.modes.split(' ').includes(UI.mode); }); }
$('segMode').addEventListener('click', (e) => {
  const v = e.target.dataset.v; if (!v) return; UI.mode = v;
  $('segMode').querySelectorAll('button').forEach((b) => b.classList.toggle('on', b.dataset.v === v));
  syncModeCtrls();
  renderTrend();
});
$('targetPct').addEventListener('change', (e) => { UI.target = e.target.value; renderTrend(); });
$('splitBy').addEventListener('change', (e) => { UI.split = e.target.value; renderTrend(); });
$('treeLevel').addEventListener('change', (e) => { UI.treeLevel = +e.target.value; renderTops(); });
$('exportBtn').addEventListener('click', exportXlsx);

/* ───────────── Start ───────────── */
setNumberLocale();
applyI18n();
document.querySelectorAll('select.ui-lang').forEach(fillUiLangSelect);
tryRestoreSession();
