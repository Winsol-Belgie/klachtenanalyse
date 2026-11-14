/* Winsol Klachtenanalyse – Productinfo
 * Productkenmerken (AS400-specificaties, queries P01–P38) en verkochte stuks (verkoopstatistiek, query P00).
 *
 * Telregels (zie ook README, hoofdstuk 9):
 *  - Eén "verkochte eenheid" = één order × artikelgroep uit de verkoopstatistiek (P00): netto stuks (enkel het hoofdartikel),
 *    jaar/maand = eerste boeking in de statistiek. Enkel orders met netto stuks > 0 EN netto omzet > 0: zo vallen geannuleerde
 *    orders (netto 0) en correctie-/creditorders (stuks +1, omzet negatief) weg. Zo komen de aantallen overeen met WinPro.
 *  - De kenmerken komen uit de specificatietabel van dezelfde productfamilie, gekoppeld op ordernummer (eerste lijn van het order).
 *    Orders zonder specificatie tellen mee in de totalen met kenmerk "onbekend".
 *  - Motoren: omzettingslijst MOTOREN (PRODUCTINFO_OPZOEKLIJSTEN.xlsx), door Gwenn nagekeken. RTS telt als io.
 *  - Klacht ↔ order: via het oorspronkelijk order van het dossier (rootOrder), anders het order van de klacht.
 *
 * Opslag: één datablok per productfamilie onder de gereserveerde nummers 999900–999998 (zoals de verkoop per product onder 999999).
 * Verkopers en directeurs (rol sales/dir) krijgen die blokken niet: voor hen verschijnen de kaart en het tabblad niet.
 */
'use strict';

const PI_VT0 = 999900, PI_VT1 = 999998;
const PI_UNK = '?';               // waarde "onbekend" (getoond als t('pkUnknown'))
const PI_NA = '–';                // niet van toepassing
const isPiVt = (vt) => { const n = Number(vt); return n >= PI_VT0 && n <= PI_VT1; };

/* ───────────── Kenmerken ─────────────
 * lbl/grp = sleutel in i18n; shared = bruikbaar over productfamilies heen (kaart bij keuze van een productgroep). */
const PI_KEN = {
  prod:     { lbl: 'pkProd',     grp: 'pkGrpProduct', shared: true },
  uitv:     { lbl: 'pkUitv',     grp: 'pkGrpProduct' },
  type:     { lbl: 'pkType',     grp: 'pkGrpProduct' },
  pack:     { lbl: 'pkPack',     grp: 'pkGrpProduct' },
  kast:     { lbl: 'pkKast',     grp: 'pkGrpProduct' },
  montage:  { lbl: 'pkMontage',  grp: 'pkGrpProduct' },
  motor:    { lbl: 'pkMotor',    grp: 'pkGrpBed', shared: true, order: ['Somfy', 'Winsol', 'Ander merk', 'Manueel', PI_UNK] },
  motorDet: { lbl: 'pkMotorDet', grp: 'pkGrpBed', shared: true },
  aanst:    { lbl: 'pkAanst',    grp: 'pkGrpBed', shared: true },
  bedmodus: { lbl: 'pkBedModus', grp: 'pkGrpBed' },
  verl:     { lbl: 'pkVerl',     grp: 'pkGrpVerl' },
  dim:      { lbl: 'pkDim',      grp: 'pkGrpVerl', order: ['Ja', 'Nee', PI_NA, PI_UNK] },
  led:      { lbl: 'pkLed',      grp: 'pkGrpVerl' },
  solar:    { lbl: 'pkSolar',    grp: 'pkGrpVerl', order: ['Ja', 'Nee', PI_UNK] },
  solarUitv:{ lbl: 'pkSolarUitv',grp: 'pkGrpVerl' },
  doek:     { lbl: 'pkDoek',     grp: 'pkGrpKleur' },
  kleur:    { lbl: 'pkKleur',    grp: 'pkGrpKleur', shared: true },
  kleur2:   { lbl: 'pkKleur2',   grp: 'pkGrpKleur' },
  breedte:  { lbl: 'pkBreedte',  grp: 'pkGrpAfm' },
  uitval:   { lbl: 'pkUitval',   grp: 'pkGrpAfm' },
  hoogte:   { lbl: 'pkHoogte',   grp: 'pkGrpAfm' },
  lengte:   { lbl: 'pkLengte',   grp: 'pkGrpAfm' },
};

/* ───────────── Hulpfuncties voor het inlezen ───────────── */
const piS = (v) => String(v ?? '').replace(/\|/g, '!').trim();
const piN = (v) => { const n = Number(v); return Number.isFinite(n) ? n : 0; };
const piX = (v) => { const s = piS(v).toUpperCase(); return s !== '' && !['0', 'N', 'NEE', '-', 'NO'].includes(s); };
function piBin(mm, edges, unit) {
  // edges in m, bv. [3,4,5,6] → "< 3 m", "3–4 m", …, "≥ 6 m"
  const n = piN(mm); if (n <= 0) return PI_UNK;
  const m = n / 1000; const f = (x) => String(x).replace('.', ',');
  if (m < edges[0]) return `< ${f(edges[0])} ${unit}`;
  for (let i = 1; i < edges.length; i++) if (m < edges[i]) return `${f(edges[i - 1])}–${f(edges[i])} ${unit}`;
  return `≥ ${f(edges[edges.length - 1])} ${unit}`;
}

function piCtx(sheets) {
  // Opzoeklijsten: uit PRODUCTINFO_OPZOEKLIJSTEN.xlsx (MOTOREN, DOEKEN, LAND, PRODUCTEN, KLEUR_RAL) of de L-queries (L1, L2, L6, L8).
  const ctx = { mot: new Map(), doek: new Map(), land: {}, pn: {}, ral: new Set() };
  const rows = (s) => (s ? s.rows.map((r) => Object.fromEntries(s.hdr.map((h, i) => [h, r[i]]))) : []);
  for (const r of rows(sheets.MOTOREN)) {
    const n = piS(r.MOTORNAAM_AS400).toUpperCase(); if (!n) continue;
    ctx.mot.set(n, { g: piS(r.MERKGROEP), m: piS(r.MERK), a: piS(r.AANSTURING) });
  }
  for (const r of rows(sheets.DOEKEN || sheets.L1_DOEKEN)) {
    const nr = piS(r.DOEKNR).toUpperCase(); if (!nr) continue;
    const soort = piS(r.SOORT); const col = piS(r.COLLECTIE || r.KWALITEIT).replace(/^\*+|\*+$/g, '').replace(/^2ANCY/, 'FANCY').replace(/^2OLDIES/, 'GOLDIES');
    const cur = ctx.doek.get(nr);
    if (!cur || (soort === 'Z' && cur.soort !== 'Z')) ctx.doek.set(nr, { soort, col: col || PI_UNK });
  }
  for (const r of rows(sheets.LAND || sheets.L6_LAND)) { const c = piN(r.LANDCODE ?? r.KO_NAT); if (c) ctx.land[c] = piS(r.LAND_NL || r.OMS_NATN); }
  for (const r of rows(sheets.PRODUCTEN || sheets.L2_PRODUCTEN)) { const c = piN(r.PRODUCT_ID); if (c) ctx.pn[c] = piS(r.NAAM_NL); }
  for (const r of rows(sheets.KLEUR_RAL || sheets.L8_RAL)) { const c = piS(r.RAL || r.RALNR); if (/^\d{4}$/.test(c)) ctx.ral.add(c); }
  if (!ctx.ral.size) for (let i = 1000; i <= 9023; i++) ctx.ral.add(String(i));
  return ctx;
}

/* Kleur: RAL-nummer uit de vrije tekst (bv. "7016 ST AE0305770162", "R9010G", "RAL 9016"); de rest = "Andere kleur". */
function piRal(ctx, ...vals) {
  for (const v of vals) {
    const s = piS(v).toUpperCase(); if (!s || s === '0') continue;
    const m = s.match(/(?:^|[^\d])(\d{4})(?!\d)/);
    if (m && ctx.ral.has(m[1])) return 'RAL ' + m[1];
    return 'Andere kleur';
  }
  return PI_UNK;
}
/* Motor: omzettingslijst; zonder tekst de oude vinkjes. */
function piMotor(ctx, name, flags) {
  const n = piS(name).toUpperCase();
  if (n && n !== '0') {
    const m = ctx.mot.get(n);
    if (m) {
      if (m.g === 'n.v.t.') return {};
      const g = m.g === 'Manueel' || m.m === 'manueel' ? 'Manueel' : m.g || PI_UNK;
      const a = !m.a || /^onbekend/i.test(m.a) ? PI_UNK : m.a;
      return { motor: g, motorDet: g === 'Manueel' ? 'Manueel' : (m.m || PI_UNK), aanst: a || PI_UNK };
    }
    if (/STANG|LINT|KOORD|SCHIJF|VEER|MANUEEL/.test(n)) return { motor: 'Manueel', motorDet: 'Manueel', aanst: 'manueel' };
    if (/A-?OK|WINSOL|WINS\./.test(n)) return { motor: 'Winsol', motorDet: 'Winsol (A-OK)', aanst: PI_UNK };
    if (/SOMF|SUNEA|SUNILUS|ALTUS|OXIMO|OREA|ILMO|MAESTRIA|\bJET\b|METEOR|ATLAS|HELIOS|MARINER|GEMINI|APOLLO|ORION|VEGA|VECTRAN|SIRIUS|JUPITER|LT ?50|\bRTS\b/.test(n))
      return { motor: 'Somfy', motorDet: 'Somfy', aanst: /\bIO\b|\bRTS\b/.test(n) ? 'io' : /\bLT\b|\bWT\b|JET|METEOR|ATLAS|HELIOS|MARINER|VEGA|SIRIUS/.test(n) ? 'bekabeld' : PI_UNK };
    return { motor: PI_UNK, motorDet: PI_UNK, aanst: PI_UNK };
  }
  if (flags) {
    if (flags.win) return { motor: 'Winsol', motorDet: 'Winsol (A-OK)', aanst: PI_UNK };
    if (flags.rts || flags.io || flags.sunea) return { motor: 'Somfy', motorDet: 'Somfy', aanst: 'io' };
    if (flags.stang) return { motor: 'Manueel', motorDet: 'Manueel', aanst: 'manueel' };
  }
  return { motor: PI_UNK, motorDet: PI_UNK, aanst: PI_UNK };
}
const piCode = (v) => { const s = piS(v); return s && s !== '0' ? 'Code ' + s : PI_UNK; };
/* Kleur als naam (WIT, BRUIN, ANTRAC…) of RAL in de tekst. */
function piKleurNaam(ctx, v) {
  const s = piS(v).toUpperCase(); if (!s || s === '0') return PI_UNK;
  const m = s.match(/(\d{4})/); if (m && ctx.ral.has(m[1])) return 'RAL ' + m[1];
  return s.charAt(0) + s.slice(1).toLowerCase();
}
const piDoek = (ctx, nr) => { const k = piS(nr).toUpperCase(); if (!k || k === '0') return PI_UNK; const d = ctx.doek.get(k); return d ? d.col : PI_UNK; };

/* Verlichting luifels: nieuw veld (2024+: Geen / KAST / ARM / VOORL, DB = dimbaar), anders de oude vinkjes (kast, voorlijst; D = dimbaar).
 * Oude orders: verlichting in de armen bestond niet als vinkje → zonder vinkje = onbekend (geen of in de armen). */
function piVerl(newTxt, kas, vp, kvp) {
  const s = piS(newTxt).toUpperCase();
  if (s && s !== '0') {
    if (/GEEN/.test(s)) return { verl: 'Geen', dim: PI_NA };
    const parts = [/KAST/.test(s) && 'Kast', /VOORL/.test(s) && 'Voorlijst', /ARM/.test(s) && 'Armen'].filter(Boolean);
    return { verl: parts.join(' + ') || 'Andere', dim: /\bDB\b/.test(s) ? 'Ja' : 'Nee' };
  }
  const k = piS(kas).toUpperCase(), v = piS(vp).toUpperCase(), kv = piS(kvp).toUpperCase();
  const on = (x) => x !== '' && x !== '0';
  if (!on(k) && !on(v) && !on(kv)) return { verl: PI_UNK, dim: PI_UNK };
  const verl = on(kv) || (on(k) && on(v)) ? 'Kast + Voorlijst' : on(k) ? 'Kast' : 'Voorlijst';
  const fl = [k, v, kv];
  return { verl, dim: fl.includes('D') ? 'Ja' : fl.includes('X') ? 'Nee' : PI_UNK };
}

/* ───────────── Productfamilies ─────────────
 * ags = artikelgroepen in de verkoopstatistiek; sheets = specificatie-queries; ken = kenmerken; x = extractie uit één specificatielijn.
 * g(col) = waarde van die kolom op de specificatielijn; u = { prodName } van de verkochte eenheid. */
const LUIFEL_KEN = ['prod', 'verl', 'dim', 'solar', 'motor', 'motorDet', 'aanst', 'doek', 'kleur', 'breedte', 'uitval'];
function luifelX(newCol) {
  return (g, ctx, u) => ({
    ...piVerl(newCol ? g(newCol) : '', g('S2LKAS'), g('S2LVP'), g('S2LKVP')),
    solar: /SOLAR/i.test(u.prodName) || piX(g('LUSOLAR')) && /JA/i.test(piS(g('LUSOLAR'))) ? 'Ja' : 'Nee',
    ...piMotor(ctx, g('BEDIENING'), { rts: piX(g('S1RTS')), sunea: piX(g('S1MSUN')), stang: piX(g('S1TYST')) }),
    doek: piDoek(ctx, g('S1NRDO')),
    kleur: piRal(ctx, g('KLEUR'), g('S2KGAR')),
    breedte: piBin(g('S2BREO'), [3, 4, 5, 6], 'm'),
    uitval: piBin(g('S2ARML'), [2, 2.5, 3, 3.5], 'm'),
  });
}
const ROLLUIK_KEN = ['prod', 'solar', 'motor', 'motorDet', 'aanst', 'kleur', 'kleur2', 'breedte', 'hoogte'];
function rolluikX(bedCol, kastKleurCol, bladKleurCol, bCol, hCol) {
  return (g, ctx, u) => ({
    solar: /SOLAR/i.test(u.prodName) || piX(g('SOLARB')) ? 'Ja' : 'Nee',
    ...piMotor(ctx, g(bedCol), null),
    kleur: kastKleurCol ? piKleurNaam(ctx, g(kastKleurCol)) : PI_UNK,
    kleur2: piCode(g(bladKleurCol)),
    breedte: piBin(g(bCol), [1, 1.5, 2, 3], 'm'),
    hoogte: piBin(g(hCol), [1, 1.5, 2, 2.5], 'm'),
  });
}
const PI_FAMS = [
  { id: 'SOLFIX', label: 'Solfix / Solarfix', ags: [338, 297, 298, 300, 301, 302, 303, 304, 306, 307, 308], sheets: ['P10_SOLFIX', 'P11_SOLFIX'],
    ken: ['prod', 'uitv', 'kast', 'montage', 'solarUitv', 'motor', 'motorDet', 'aanst', 'doek', 'kleur', 'breedte', 'hoogte'],
    x: (g, ctx, u) => {
      const n = u.prodName.toUpperCase();
      return {
        uitv: /SOLARFIX/.test(n) || piX(g('SOLARF')) ? 'Solarfix' : 'Solfix',
        kast: (n.match(/\b(90|110|124|125|150)\b/) || [])[1] ? 'Kast ' + n.match(/\b(90|110|124|125|150)\b/)[1] : PI_UNK,
        montage: / RE\b/.test(n) ? 'RE' : / AR\b/.test(n) ? 'AR' : / IN\b/.test(n) ? 'IN' : PI_UNK,
        solarUitv: piS(g('SOLARFTY')) && piS(g('SOLARFTY')) !== '0' ? piS(g('SOLARFTY')) : PI_NA,
        ...piMotor(ctx, g('BEDIENING'), { rts: piX(g('SMRTS')) || piX(g('SMARTS')), win: piX(g('SMWIN')), sunea: piX(g('S1MSUN')) }),
        doek: piDoek(ctx, g('SDKNR')),
        kleur: piRal(ctx, g('KLEUR'), g('S2PDC'), g('SKRAL')),
        breedte: piBin(g('SBRET'), [1.5, 2.5, 3.5, 4.5], 'm'),
        hoogte: piBin(g('SHOOT'), [1.5, 2.5, 3.5], 'm'),
      };
    } },
  { id: 'SOLSCREEN', label: 'Solscreen AR', ags: [335], sheets: ['P12_SOLSCREEN_AR'],
    ken: ['prod', 'motor', 'motorDet', 'aanst', 'doek', 'kleur', 'breedte', 'hoogte'],
    x: (g, ctx) => ({
      ...piMotor(ctx, g('BEDIENING'), { rts: piX(g('SMRTS')) || piX(g('SMARTS')), win: piX(g('SMWIN')), sunea: piX(g('S1MSUN')) }),
      doek: piDoek(ctx, g('SDKNR')), kleur: piRal(ctx, g('S2PDC'), g('SKRAL')),
      breedte: piBin(g('SBRET'), [1.5, 2.5, 3.5, 4.5], 'm'), hoogte: piBin(g('SHOOT'), [1.5, 2.5, 3.5], 'm'),
    }) },
  { id: 'LUMISOL', label: 'Lumisol', ags: [546, 487], sheets: ['P16_LUMISOL'], ken: LUIFEL_KEN, x: luifelX('VERLICHT') },
  { id: 'LINASOL', label: 'Linasol', ags: [516, 488], sheets: ['P17_LINASOL'], ken: LUIFEL_KEN, x: luifelX('VERL06') },
  { id: 'SQUARO', label: 'Squaro', ags: [549], sheets: ['P18_SQUARO'], ken: LUIFEL_KEN, x: luifelX('VERL19') },
  { id: 'LUNO', label: 'Luno', ags: [547, 557, 609], sheets: ['P19_LUNO'], ken: LUIFEL_KEN, x: luifelX('VERL83') },
  { id: 'COMBISOL1200', label: 'Combisol 1200', ags: [536], sheets: ['P20_COMBISOL1200'], ken: LUIFEL_KEN, x: luifelX('VERLICHT') },
  { id: 'COMBISOL2500', label: 'Combisol 2500', ags: [427], sheets: ['P21_COMBISOL2500'], ken: LUIFEL_KEN, x: luifelX('') },
  { id: 'COMBISOL550', label: 'Combisol 550', ags: [539], sheets: ['P22_COMBISOL550'], ken: LUIFEL_KEN, x: luifelX('') },
  { id: 'LOFT', label: 'Loft', ags: [559], sheets: ['P23_LOFT'], ken: LUIFEL_KEN, x: luifelX('') },
  { id: 'COMBISOL3050', label: 'Combisol 3050', ags: [541], sheets: ['P27_COMBISOL3050'], ken: LUIFEL_KEN, x: luifelX('') },
  { id: 'SO', label: 'SO!', ags: [586, 587, 588, 589, 605, 606, 607, 608, 620, 625, 626, 627], sheets: ['P01_SO'],
    ken: ['prod', 'pack', 'bedmodus', 'led', 'kleur', 'kleur2', 'lengte', 'breedte'],
    x: (g, ctx, u) => {
      const p = piS(g('PACK')); const n = u.prodName.toUpperCase();
      return {
        pack: /^(COSY|CHIC|STAR)$/i.test(p) ? p[0].toUpperCase() + p.slice(1).toLowerCase() : /ORIG/.test(n) ? 'Cosy' : PI_UNK,
        bedmodus: piS(g('MOTCOSY')) && piS(g('MOTCOSY')) !== '0' ? piS(g('MOTCOSY')) : PI_UNK,
        led: piS(g('OREGDI')) && !['0', '-'].includes(piS(g('OREGDI'))) ? piS(g('OREGDI')) : piS(g('OREGDI')) === '-' ? 'Geen' : PI_UNK,
        kleur: piRal(ctx, g('S2PDC')), kleur2: piRal(ctx, g('S3PDC')),
        lengte: piBin(g('S2LENG'), [3, 4, 5, 6], 'm'), breedte: piBin(g('S2BREO'), [3, 4, 5, 6], 'm'),
      };
    } },
  { id: 'SO_RETROFIT', label: 'SO! Retrofit', ags: [590], sheets: ['P30_RETROFIT'], ken: ['prod'], x: () => ({}) },
  { id: 'ZIP', label: 'Z!P / Bienestar', ags: [603, 604], sheets: ['P02_ZP'], ken: ['prod', 'pack', 'led', 'kleur', 'breedte', 'uitval'],
    x: (g, ctx) => ({
      pack: piS(g('PACK')) && piS(g('PACK')) !== '0' ? piS(g('PACK')) : PI_UNK,
      led: piX(g('LEDKAST')) || piX(g('LEDGEL')) || piX(g('STRIPL23')) ? 'Ja' : 'Nee',
      kleur: piRal(ctx, g('S2PDC'), g('SKGAR')),
      breedte: piBin(g('BREEDTE'), [3, 4, 5, 6], 'm'), uitval: piBin(g('UITVAL'), [3, 4, 5, 6], 'm'),
    }) },
  { id: 'VERANDASOL', label: 'Verandasol ST2220 / ST2100', ags: [469, 466], sheets: ['P03_VERANDASOL'], ken: ['prod', 'motor', 'motorDet', 'aanst', 'doek', 'kleur', 'breedte'],
    x: (g, ctx) => ({ ...piMotor(ctx, piS(g('MOT1')).replace(/\s+(LI|RE|L|R)$/i, ''), { sunea: piX(g('S1MSUN')) }), doek: piDoek(ctx, g('SDKNR')), kleur: piRal(ctx, g('S2PDC'), g('SGARA')), breedte: piBin(g('STBVD'), [3, 4, 5, 6], 'm') }) },
  { id: 'ST2300', label: 'Verandasol ST2300', ags: [545], sheets: ['P04_ST2300'], ken: ['prod', 'motor', 'motorDet', 'aanst', 'doek', 'kleur', 'breedte'],
    x: (g, ctx) => ({ ...piMotor(ctx, '', { rts: piX(g('S1RTS')), sunea: piX(g('S1MSUN')) }), doek: piDoek(ctx, g('S1NRDO')), kleur: piRal(ctx, g('S2KGAR')), breedte: piBin(g('S2BREO'), [3, 4, 5, 6], 'm') }) },
  { id: 'WINCUBE', label: 'WinCube', ags: [540], sheets: ['P05_WINCUBE'], ken: ['prod', 'motor', 'aanst', 'verl', 'kleur', 'breedte'],
    x: (g, ctx) => ({ ...piMotor(ctx, '', { rts: piX(g('S2RTS')), io: piX(g('S2IO')) }), verl: piX(g('S2VERL')) ? 'Ja' : 'Nee', kleur: piRal(ctx, g('S2PDC'), g('S2KGAR')), breedte: piBin(g('S2BREO'), [3, 4, 5, 6], 'm') }) },
  { id: 'WINCUBE_XL', label: 'WinCube XL', ags: [583], sheets: ['P06_WINCUBE_XL'], ken: ['prod', 'motor', 'aanst', 'verl', 'kleur', 'breedte'],
    x: (g, ctx) => ({ ...piMotor(ctx, '', { rts: piX(g('S2RTS')), io: piX(g('S2IO')) }), verl: piX(g('S2VERL')) ? 'Ja' : 'Nee', kleur: piRal(ctx, g('S2PDC'), g('S2KGAR')), breedte: piBin(g('S2BREO'), [3, 4, 5, 6], 'm') }) },
  { id: 'ALUBOX', label: 'Alubox', ags: [447], sheets: ['P07_ALUBOX'], ken: ['prod', 'motor', 'aanst', 'doek', 'kleur', 'breedte'],
    x: (g, ctx) => ({ ...piMotor(ctx, '', { rts: piX(g('S1RTS')), sunea: piX(g('S1MSUN')), stang: piX(g('S1TYST')) }), doek: piDoek(ctx, g('S1NRDO')), kleur: piRal(ctx, g('S2KGAR')), breedte: piBin(g('S2BREO'), [3, 4, 5, 6], 'm') }) },
  { id: 'TERRASOL', label: 'Terrasol II', ags: [437], sheets: ['P08_TERRASOL'], ken: ['prod', 'doek', 'kleur', 'breedte'],
    x: (g, ctx) => ({ doek: piDoek(ctx, g('S1NRDO')), kleur: piRal(ctx, g('S2KGAR')), breedte: piBin(g('S2BREO'), [3, 4, 5, 6], 'm') }) },
  { id: 'PARAVENTO', label: 'Paravento', ags: [385], sheets: ['P28_PARAVENTO'], ken: ['prod'], x: () => ({}) },
  { id: 'PATIOLA', label: 'Patiola', ags: [468], sheets: ['P29_PATIOLA'], ken: ['prod'], x: () => ({}) },
  { id: 'POORTEN', label: 'Poorten', ags: [521, 519, 522, 523, 525], sheets: ['P09_POORTEN'], ken: ['prod', 'type', 'bedmodus', 'kleur', 'breedte'],
    x: (g, ctx) => ({ type: piS(g('ALTYPE')) || PI_UNK, bedmodus: (() => { const b = piS(g('ALBED')).toUpperCase(); return !b ? PI_UNK : /MANUEEL/.test(b) ? 'Manueel' : /^BERN/.test(b) ? 'Motor Bernal' : /SOMMER/.test(b) ? 'Motor Sommer' : /^MOTOR D/.test(b) ? 'Motor ' + b.replace(/^MOTOR\s*/, '') : 'Andere'; })(), kleur: piRal(ctx, g('ALKLEUR')), breedte: piBin(g('ALBRED'), [2.5, 3, 4, 5], 'm') }) },
  { id: 'COMBIFLEX', label: 'Voorzetrolluik (Combiflex)', ags: [314, 311], sheets: ['P32_COMBIFLEX'], ken: ROLLUIK_KEN, x: rolluikX('BEDNO', 'KASTK', 'SBLADK', 'SBREE', 'SHOOG') },
  { id: 'FUSION', label: 'Fusion / SolarFuse', ags: [325, 326], sheets: ['P33_FUSION'], ken: ROLLUIK_KEN, x: rolluikX('BEDNO', 'KASTK', 'SBLADK', 'SBREE', 'SHOOG') },
  { id: 'MINICAISSON', label: 'Mini Caisson', ags: [319], sheets: ['P37_MINICAISSON'], ken: ROLLUIK_KEN, x: rolluikX('BEDNO', '', 'SBLADK', 'SBREE', 'SHOOG') },
  { id: 'INBOUW', label: 'Inbouwrolluik', ags: [111, 116], sheets: ['P34_INBOUWROLLUIK'], ken: ['prod', 'motor', 'motorDet', 'aanst', 'kleur2', 'breedte', 'hoogte'], x: (g, ctx) => ({ ...piMotor(ctx, g('LINTS'), null), kleur2: piCode(g('SKLEUR')), breedte: piBin(g('SBR'), [1, 1.5, 2, 3], 'm'), hoogte: piBin(g('SHO'), [1, 1.5, 2, 2.5], 'm') }) },
  { id: 'OPBOUW', label: 'Opbouwrolluik', ags: [241, 162], sheets: ['P35_OPBOUWROLLUIK'], ken: ['prod', 'motor', 'motorDet', 'aanst', 'kleur2', 'breedte', 'hoogte'], x: (g, ctx) => ({ ...piMotor(ctx, g('AANGAR'), null), kleur2: piCode(g('SKLEUR')), breedte: piBin(g('SBR'), [1, 1.5, 2, 3], 'm'), hoogte: piBin(g('SHO'), [1, 1.5, 2, 2.5], 'm') }) },
  { id: 'VOLETTRADI', label: 'Volet Tradi', ags: [114], sheets: ['P38_VOLETTRADI'], ken: ['prod', 'motor', 'motorDet', 'aanst', 'kleur2', 'breedte', 'hoogte'], x: (g, ctx) => ({ ...piMotor(ctx, g('LINTS'), null), kleur2: piCode(g('SKLEUR')), breedte: piBin(g('SBR'), [1, 1.5, 2, 3], 'm'), hoogte: piBin(g('SHO'), [1, 1.5, 2, 2.5], 'm') }) },
];

/* ───────────── Inlezen (admin) ─────────────
 * sheets: { naam: { hdr: [...], rows: [[...]] } } uit één of meer Excel-bestanden.
 * Resultaat: { fams: [ payload per familie ], stats } */
const PI_SHEET_RX = /^(P\d\d_|L\d_|MOTOREN$|DOEKEN$|LAND$|PRODUCTEN$|KLEUR_RAL$|DEFINITIEF$)/i;
function piFindSheets(wb, XLSXlib, into) {
  const want = PI_SHEET_RX;
  for (const name of wb.SheetNames) {
    if (!want.test(name)) continue;
    const rows = XLSXlib.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: false });
    if (!rows.length) continue;
    const def = /^definitief$/i.test(name);
    const key = def ? 'MOTOREN' : name.toUpperCase();
    const hdr = rows[0].map((h) => String(h ?? '').trim().toUpperCase());
    if (key === 'MOTOREN') {
      // Omzettingslijst motoren: enkel met de definitieve kolommen (MERKGROEP…); het tabblad "Definitief" van
      // PRODUCTINFO_MOTOREN.xlsx gaat voor op het tabblad MOTOREN van PRODUCTINFO_OPZOEKLIJSTEN.xlsx.
      if (!hdr.includes('MERKGROEP')) continue;
      if (into.MOTOREN && into.MOTOREN.def && !def) continue;
    }
    into[key] = { hdr, rows: rows.slice(1), def };
  }
  return into;
}
function piBuild(sheets) {
  const P00 = sheets.P00_VERKOOPSTAT;
  if (!P00) throw new Error('Tabblad P00_VERKOOPSTAT niet gevonden (Verkoopstats.xlsm).');
  if (!sheets.MOTOREN) throw new Error('Tabblad MOTOREN niet gevonden (PRODUCTINFO_OPZOEKLIJSTEN.xlsx).');
  const ctx = piCtx(sheets);
  const agFam = new Map(); for (const f of PI_FAMS) for (const a of f.ags) agFam.set(a, f);
  const H = (s) => Object.fromEntries(s.hdr.map((h, i) => [h, i]));
  // 1. verkochte eenheden uit de statistiek (per order × artikelgroep), eerste bron W > B > H
  const hi = H(P00); const prio = { W: 1, B: 2, H: 3 };
  const units = new Map(); // key order|ag
  let skipped = { cancel: 0, credit: 0 };
  for (const r of P00.rows) {
    const ag = piN(r[hi.ARTGROEP]); const fam = agFam.get(ag); if (!fam) continue;
    const st = piN(r[hi.STUKS_NETTO]), om = piN(r[hi.OMZET_NETTO]);
    if (st <= 0) { skipped.cancel++; continue; }
    if (om <= 0) { skipped.credit++; continue; }
    const o = piN(r[hi.ORDERNR]); const pr = prio[piS(r[hi.BRON])] || 9; const k = o + '|' + ag;
    const cur = units.get(k); if (cur && cur.pr <= pr) continue;
    const d = piN(r[hi.EERSTE_BOEKING]); if (d < 20000101) continue;
    const pid = piN(r[hi.PRODUCT_ID]);
    units.set(k, { pr, o, ag, fam, st, ym: Math.floor(d / 100), pid, klnr: piN(r[hi.KLANTNR]) });
  }
  // 2. specificaties per familie: eerste lijn per order
  const spec = new Map(); // famId -> Map(order -> {g})
  const famStats = {};
  for (const f of PI_FAMS) {
    const m = new Map(); spec.set(f.id, m);
    for (const sn of f.sheets) {
      const s = sheets[sn]; if (!s) continue;
      const ix = H(s); const io = ix.ORDERNR, iv = ix.VOLGNR, ib = ix.BRON, il = ix.LAND;
      for (const r of s.rows) {
        const o = piN(r[io]); if (!o) continue;
        const cur = m.get(o); const v = piN(r[iv]);
        const pr = ({ W: 1, B: 2, H: 3, I: 4 })[piS(r[ib])] || 9;
        if (cur && (cur.pr < pr || (cur.pr === pr && cur.v <= v))) continue;
        m.set(o, { pr, v, r, ix, land: piN(r[il]) });
      }
    }
  }
  // 3. per familie: eenheden + kenmerken → compacte tabel
  const out = [];
  const byFam = new Map(); for (const u of units.values()) { let l = byFam.get(u.fam.id); if (!l) byFam.set(u.fam.id, (l = [])); l.push(u); }
  for (const f of PI_FAMS) {
    const list = byFam.get(f.id) || []; if (!list.length) continue;
    const sm = spec.get(f.id);
    const dict = Object.fromEntries(f.ken.map((k) => [k, new Map()]));
    const prods = new Map(); const lands = new Map();
    const idx = (map, v) => { let i = map.get(v); if (i == null) { i = map.size; map.set(v, i); } return i; };
    const rows = []; let withSpec = 0, stuks = 0, stuksSpec = 0;
    for (const u of list.sort((a, b) => a.ym - b.ym || a.o - b.o)) {
      const prodName = piS(ctx.pn[u.pid] || ctx.pn[u.ag] || ('Product ' + u.ag));
      const sp = sm.get(u.o);
      let vals = {};
      if (sp) {
        const g = (c) => (c && sp.ix[c] != null ? sp.r[sp.ix[c]] : null);
        try { vals = f.x(g, ctx, { prodName }) || {}; } catch (e) { vals = {}; }
        withSpec++; stuksSpec += u.st;
      }
      vals.prod = prodName;
      stuks += u.st;
      // rij: [order, jaarmaand, stuks, artikelgroep, product-ID, productnaam (index), landcode, met specificatie (0/1), kenmerken…]
      const row = [u.o, u.ym, u.st, u.ag, u.pid || u.ag, idx(prods, prodName), sp && sp.land ? sp.land : 0, sp ? 1 : 0];
      for (const k of f.ken) { const v = vals[k] == null || vals[k] === '' ? PI_UNK : String(vals[k]); row.push(idx(dict[k], v)); }
      rows.push(row);
    }
    famStats[f.id] = { units: rows.length, withSpec, stuks, stuksSpec };
    out.push({ pi: { v: 1, f: f.id, label: f.label, ags: f.ags, ks: f.ken, d: Object.fromEntries(f.ken.map((k) => [k, [...dict[k].keys()]])), p: [...prods.keys()], land: ctx.land, u: rows, n: famStats[f.id] } });
  }
  return { fams: out, stats: { fams: famStats, skipped, units: units.size, motoren: ctx.mot.size, doeken: ctx.doek.size } };
}

const PI_OFS = 8; // eerste kolom met kenmerken in een rij
if (typeof module !== 'undefined') module.exports = { piBuild, piFindSheets, PI_FAMS, PI_KEN, PI_OFS };

/* ═══════════════════════════════════════════════════════════════════════════════════════════
 * In de browser: kaart "Productkenmerken", klachtratio per kenmerk (tab Klachten), tab "Producten", opladen (admin).
 * ═══════════════════════════════════════════════════════════════════════════════════════════ */
if (typeof window !== 'undefined') (function () {

let PU = new WeakMap();          // klacht → [familie, eenheid] (cache)
let PKS = { fams: [], keys: [] }; // actieve families en zichtbare kenmerken
let prodChart = null;
Object.assign(UI, { pkOpen: false, pkSplit: '', pkOpenKey: '', prodTopKey: '', prodMeas: 'st' });
try { UI.pkOpen = localStorage.getItem('kl_pkOpen') === '1'; } catch { /* geen opslag */ }

const PK_VAL = { Ja: 'yes', Nee: 'no', Geen: 'pkvNone', Manueel: 'pkvManual', manueel: 'pkvManual', 'Ander merk': 'pkvOther', 'Andere kleur': 'pkvOtherColour',
  Kast: 'pkvKast', Voorlijst: 'pkvVoorl', Armen: 'pkvArmen', 'Kast + Voorlijst': 'pkvKastVoorl', 'Kast + Armen': 'pkvKastArmen', 'Voorlijst + Armen': 'pkvVoorlArmen', bedraad: 'pkvWired', bekabeld: 'pkvWired' };
const pkLabel = (v) => (v === PI_UNK ? t('pkUnknown') : v === PI_NA ? t('pkNA') : PK_VAL[v] ? t(PK_VAL[v]) : v);
const pkName = (k) => t(PI_KEN[k] ? PI_KEN[k].lbl : k);
const ymOf = (dn) => { const d = new Date(dn * 864e5); return d.getUTCFullYear() * 100 + d.getUTCMonth() + 1; };
const pct = (a, b) => (b ? Math.round((a / b) * 1000) / 10 : 0);
const fmtPct = (a, b) => (b ? nf.format(pct(a, b)) + '%' : '–');

/* ───────────── Data ───────────── */
window.setPiData = function (parts) {
  PU = new WeakMap();
  if (!parts || !parts.length) { window.PI = null; return; }
  const order = new Map(PI_FAMS.map((f, i) => [f.id, i]));
  parts.sort((a, b) => (order.get(a.f) ?? 99) - (order.get(b.f) ?? 99));
  const fams = new Map(), agFam = new Map(), byOrder = new Map();
  for (const p of parts) {
    const f = Object.assign({}, p, { col: Object.fromEntries(p.ks.map((k, i) => [k, PI_OFS + i])) });
    fams.set(p.f, f);
    for (const ag of p.ags) agFam.set(String(ag), f);
    for (const row of p.u) { let l = byOrder.get(row[0]); if (!l) byOrder.set(row[0], (l = [])); l.push([f, row]); }
  }
  window.PI = { fams, agFam, byOrder, agPg: null, at: parts.reduce((m, p) => (p.at && p.at > m ? p.at : m), '') };
  if (typeof window.showPiInfo === 'function') window.showPiInfo();
};
function agPg() {
  // artikelgroep → productgroep (verkoopgroep), zoals in de filter "Productgroep"
  if (PI.agPg) return PI.agPg;
  const cnt = new Map();
  const add = (ag, pg, n) => { if (!pg) return; let m = cnt.get(ag); if (!m) cnt.set(ag, (m = new Map())); m.set(pg, (m.get(pg) || 0) + n); };
  if (DATA.sales && DATA.sales.prod) for (const x of DATA.sales.prod) add(String(x[2]), x[3], 1);
  for (const r of DATA.rows) add(String(r.prod), r.pg, 1);
  PI.agPg = new Map([...cnt].map(([ag, m]) => [ag, [...m].sort((a, b) => b[1] - a[1])[0][0]]));
  return PI.agPg;
}
const piVal = (f, row, k) => { const c = f.col[k]; return c == null ? PI_NA : f.d[k][row[c]]; };
function pkUnit(r) {
  let u = PU.get(r); if (u) return u;
  const f = PI.agFam.get(String(r.prod)) || null; let row = null;
  if (f) for (const o of [r.rootOrder, r.order, r.ntrOrder]) {
    if (!o) continue; const l = PI.byOrder.get(Number(o)); if (!l) continue;
    const hit = l.find((x) => x[0] === f); if (hit) { row = hit[1]; break; }
  }
  u = [f, row]; PU.set(r, u); return u;
}
const pkValOf = (f, row, k) => (!f ? PI_NA : !row ? (f.col[k] == null ? PI_NA : PI_UNK) : piVal(f, row, k));

// Welke families en kenmerken horen bij de gekozen product(groep)en? Kenmerken die niet meer zichtbaar zijn, vallen weg.
window.pkSync = function () {
  if (!window.PI) { PKS = { fams: [], keys: [] }; PKF.clear(); return; }
  const all = [...PI.fams.values()];
  let fams = all;
  if (F.prod.size) fams = all.filter((f) => f.ags.some((a) => F.prod.has(String(a))));
  else if (F.pg.size) { const ap = agPg(); fams = all.filter((f) => f.ags.some((a) => F.pg.has(ap.get(String(a))))); }
  let keys;
  if (fams.length === 1) keys = fams[0].ks.slice();
  else { const s = new Set(); for (const f of fams) for (const k of f.ks) if (PI_KEN[k] && PI_KEN[k].shared) s.add(k); keys = Object.keys(PI_KEN).filter((k) => s.has(k)); }
  PKS = { fams, keys };
  for (const k of [...PKF.keys()]) if (!keys.includes(k)) PKF.delete(k);
  if (UI.pkSplit && !keys.includes(UI.pkSplit)) UI.pkSplit = '';
};
const PKF = new Map(); window.PKF = PKF;

window.pkMatch = function (r) {
  if (!window.PI || !PKF.size) return true;
  const [f, row] = pkUnit(r);
  for (const [k, set] of PKF) if (!set.has(pkValOf(f, row, k))) return false;
  return true;
};

// Verkochte eenheden volgens de filters (periode = datum eerste boeking). skipKey = dit kenmerk niet filteren (tellers in de keuzelijst).
function eachUnit(fn, opt) {
  const o = opt || {};
  const y0 = !o.noPeriod && F.from != null ? ymOf(F.from) : 0, y1 = !o.noPeriod && F.to != null ? ymOf(F.to) : 999999;
  const ap = F.pg.size && !F.prod.size ? agPg() : null;
  const kf = [...PKF].filter(([k]) => k !== o.skipKey);
  for (const f of PKS.fams) {
    const cols = kf.map(([k, set]) => [f.col[k], set, f.d[k]]);
    for (const row of f.u) {
      if (row[1] < y0 || row[1] > y1) continue;
      if (F.prod.size && !F.prod.has(String(row[3]))) continue;
      if (ap && !F.pg.has(ap.get(String(row[3])))) continue;
      let ok = true;
      for (const [c, set, d] of cols) if (!set.has(c == null ? PI_NA : d[row[c]])) { ok = false; break; }
      if (ok) fn(f, row);
    }
  }
}

/* ───────────── Kaart "Productkenmerken" ───────────── */
window.pkClear = function () { PKF.clear(); };
window.pkChips = function (add) {
  for (const [k, set] of PKF) {
    const vals = [...set].map(pkLabel);
    add(`${pkName(k)}: ${vals.length <= 2 ? vals.join(', ') : t('msChosen', { n: vals.length })}`, () => PKF.delete(k), vals.join('\n'));
  }
};
function setPk(k, vals) { if (vals && vals.size) PKF.set(k, vals); else PKF.delete(k); UI.page = 0; update(); }

window.renderPkCard = function () {
  const card = $('pkCard'); if (!card) return;
  card.hidden = !window.PI;
  const ban = $('pkBanner');
  ban.hidden = !window.PI || !PKF.size;
  if (!ban.hidden) ban.textContent = t('pkBanner');
  if (!window.PI) return;
  const { fams, keys } = PKS;
  $('pkSum').innerHTML = PKF.size ? [...PKF].map(([k, s]) => `<span class="pk-tag">${esc(pkName(k))}: ${esc([...s].map(pkLabel).join(', '))}</span>`).join(' ') : `<span class="muted">${esc(t('pkNoneActive') + (UI.pkOpen ? '' : ' ' + t('pkClickOpen')))}</span>`;
  $('pkClear').hidden = !PKF.size;
  card.classList.toggle('open', UI.pkOpen);
  $('pkToggle').setAttribute('aria-expanded', UI.pkOpen);
  $('pkBody').hidden = !UI.pkOpen;
  const total = PI.fams.size;
  $('pkScope').textContent = !fams.length ? t('pkNoFam') : fams.length === total ? t('pkScopeAll', { n: total }) + ' ' + t('pkScopeShared')
    : fams.length === 1 ? fams[0].label : t('pkScopeN', { n: fams.length, list: fams.slice(0, 6).map((f) => f.label).join(', ') + (fams.length > 6 ? ', …' : '') }) + ' ' + t('pkScopeShared');
  // splitsen
  const sp = $('pkSplit');
  sp.innerHTML = `<option value="">${esc(t('pkSplitNone'))}</option>` + keys.map((k) => `<option value="${k}"${UI.pkSplit === k ? ' selected' : ''}>${esc(pkName(k))}</option>`).join('');
  if (!UI.pkOpen) return;
  // per kenmerk: waarden met stuks (volgens de andere kenmerken) en dekking
  const grid = $('pkGrid');
  const prevQ = {}; grid.querySelectorAll('.ms input[type=search]').forEach((i) => { prevQ[i.dataset.k] = i.value; });
  const groups = new Map();
  for (const k of keys) {
    const cnt = new Map(); let known = 0, tot = 0;
    eachUnit((f, row) => { const v = piVal(f, row, k); cnt.set(v, (cnt.get(v) || 0) + row[2]); if (f.col[k] != null) { tot += row[2]; if (v !== PI_UNK) known += row[2]; } }, { skipKey: k });
    for (const v of PKF.get(k) || []) if (!cnt.has(v)) cnt.set(v, 0);
    const ord = PI_KEN[k].order;
    const items = [...cnt].sort((a, b) => {
      const sa = a[0] === PI_UNK || a[0] === PI_NA, sb = b[0] === PI_UNK || b[0] === PI_NA;
      if (sa !== sb) return sa ? 1 : -1;
      if (ord) { const ia = ord.indexOf(a[0]), ib = ord.indexOf(b[0]); if (ia !== ib) return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib); }
      return b[1] - a[1];
    });
    const g = PI_KEN[k].grp; if (!groups.has(g)) groups.set(g, []);
    groups.get(g).push({ k, items, dek: tot ? Math.round((known / tot) * 100) : null });
  }
  const GRP = ['pkGrpProduct', 'pkGrpBed', 'pkGrpVerl', 'pkGrpKleur', 'pkGrpAfm'];
  grid.innerHTML = [...groups].sort((a, b) => GRP.indexOf(a[0]) - GRP.indexOf(b[0])).map(([g, list]) => `<div class="pk-grp"><h4>${esc(t(g))}</h4><div class="pk-fields">${list.map(({ k, items, dek }) => {
    const sel = PKF.get(k) || new Set();
    const btn = sel.size === 0 ? t('optAll') : sel.size === 1 ? pkLabel([...sel][0]) : t('msChosen', { n: sel.size });
    return `<div class="pk-field"><label class="field-label">${esc(pkName(k))}${dek != null ? ` <span class="pk-dek${dek < 70 ? ' low' : ''}" title="${esc(t('pkDekkingTip'))}">${esc(t('pkDekking', { p: dek }))}</span>` : ''}</label>
      <div class="ms${UI.pkOpenKey === k ? ' open' : ''}" data-k="${k}"><button type="button" class="ms-btn${sel.size ? ' has' : ''}">${esc(btn)}</button>
      <div class="ms-panel">${items.length > 12 ? `<input type="search" data-k="${k}" placeholder="${esc(t('msSearch'))}" value="${esc(prevQ[k] || '')}">` : ''}<div class="ms-list">${items.map(([v, n]) =>
        `<label class="ms-item" data-s="${esc(pkLabel(v).toLowerCase())}"><input type="checkbox" value="${esc(v)}"${sel.has(v) ? ' checked' : ''}><span class="lbl">${esc(pkLabel(v))}</span><span class="cnt">${nf.format(n)}</span></label>`).join('')}</div>
      <div class="ms-foot"><span class="info">${esc(t('pkCntStuks'))}</span><span class="ms-acts"><button type="button" class="btn-outline pk-x">${esc(t('msClear'))}</button></span></div></div></div></div>`;
  }).join('')}</div></div>`).join('') || `<p class="muted">${esc(t('pkNoFam'))}</p>`;
  grid.querySelectorAll('.ms').forEach((ms) => {
    const k = ms.dataset.k;
    const q = ms.querySelector('input[type=search]');
    const filt = () => { const s = q ? q.value.trim().toLowerCase() : ''; ms.querySelectorAll('.ms-item').forEach((it) => { it.hidden = !!s && !it.dataset.s.includes(s); }); };
    if (q) { q.addEventListener('input', filt); filt(); }
    ms.querySelector('.ms-btn').addEventListener('click', (e) => {
      e.stopPropagation(); const open = ms.classList.contains('open');
      document.querySelectorAll('.ms.open').forEach((m) => m.classList.remove('open'));
      UI.pkOpenKey = open ? '' : k; if (!open) { ms.classList.add('open'); if (q) q.focus(); }
    });
    ms.querySelector('.ms-panel').addEventListener('click', (e) => e.stopPropagation());
    ms.querySelector('.ms-list').addEventListener('change', () => { UI.pkOpenKey = k; setPk(k, new Set([...ms.querySelectorAll('.ms-list input:checked')].map((i) => i.value))); });
    ms.querySelector('.pk-x').addEventListener('click', () => { UI.pkOpenKey = k; setPk(k, null); });
  });
};
document.addEventListener('click', () => { UI.pkOpenKey = ''; });

/* ───────────── Klachtratio per kenmerk (tab Klachten) ───────────── */
// Verkochte orders in de periode (volgens alle filters) en het deel daarvan met minstens één klacht (alle filters behalve de periode).
function ratioBy(key, cohortKeyFn) {
  const res = new Map(); const unitOf = new Map();
  eachUnit((f, row) => {
    const v = key ? cohortKeyFn(f, row) : '';
    let e = res.get(v); if (!e) res.set(v, (e = { v, n: 0, st: 0, k: 0 }));
    e.n++; e.st += row[2]; unitOf.set(row, e);
  });
  const seen = new Set();
  for (const r of DATA.rows) {
    if (!matches(r, 'period')) continue;
    const [f, row] = pkUnit(r); if (!row || seen.has(row)) continue;
    const e = unitOf.get(row); if (!e) continue;
    seen.add(row); e.k++;
  }
  return [...res.values()];
}
window.renderPkRatio = function () {
  const card = $('pkRatioCard'); if (!card) return;
  const k = UI.pkSplit;
  card.hidden = !window.PI || !k || !PKS.fams.length;
  if (card.hidden) return;
  $('pkRatioTitle').textContent = t('pkRatioTitle', { k: pkName(k) });
  const list = ratioBy(k, (f, row) => piVal(f, row, k)).sort((a, b) => b.st - a.st);
  const tot = list.reduce((s, e) => ({ n: s.n + e.n, st: s.st + e.st, k: s.k + e.k }), { n: 0, st: 0, k: 0 });
  const maxPct = Math.max(...list.filter((e) => e.n >= 30).map((e) => pct(e.k, e.n)), 0.1);
  const rowHtml = (e, lbl, cls) => `<tr class="${cls || ''}" ${e.v != null && !cls ? `data-v="${esc(e.v)}"` : ''}><td>${esc(lbl)}</td><td class="num">${nf.format(e.n)}</td><td class="num">${nf.format(e.st)}</td><td class="num">${nf.format(e.k)}</td>
    <td class="num pk-pct"><span class="pk-pbar" style="width:${Math.min(100, (pct(e.k, e.n) / maxPct) * 100)}%"></span><b>${fmtPct(e.k, e.n)}</b>${e.n < 30 ? ` <small title="${esc(t('pkFewTip'))}">*</small>` : ''}</td></tr>`;
  $('pkRatio').innerHTML = `<thead><tr><th>${esc(pkName(k))}</th><th class="num">${esc(t('pkColOrders'))}</th><th class="num">${esc(t('pkColStuks'))}</th><th class="num">${esc(t('pkColKlacht'))}</th><th class="num">${esc(t('pkColPct'))}</th></tr></thead><tbody>` +
    list.map((e) => rowHtml(e, pkLabel(e.v))).join('') + rowHtml(tot, t('pkTotal'), 'tot') + '</tbody>';
  $('pkRatio').querySelectorAll('tr[data-v]').forEach((tr) => tr.addEventListener('click', () => setPk(k, new Set([tr.dataset.v]))));
};

/* ───────────── Tab "Producten" ───────────── */
const KLACHT_ONLY = ['tree', 'ntr', 'fout', 'klant', 'status', 'vert', 'regio', 'actie'];
function prodGroupKey() {
  if (UI.pkSplit) return { key: UI.pkSplit, val: (f, row) => piVal(f, row, UI.pkSplit), lbl: pkLabel };
  if (PKS.fams.length === 1) return { key: 'prod', val: (f, row) => piVal(f, row, 'prod'), lbl: (v) => v };
  return { key: '_fam', val: (f) => f.label, lbl: (v) => v };
}
function renderProdRep() {
  const empty = $('prodEmpty'), body = $('prodBody');
  if (!window.PI || !PKS.fams.length) { empty.hidden = false; body.hidden = true; empty.textContent = window.PI ? t('pkNoFam') : t('prodNoData'); return; }
  empty.hidden = true; body.hidden = false;
  const excl = KLACHT_ONLY.filter((k) => F[k] && F[k].size).map((k) => t(k === 'tree' ? 'lbl_tree' : 'lbl_' + k));
  if (F.text) excl.push(t('lbl_zoek').split(' (')[0]);
  $('prodNote').textContent = t('prodNote') + (excl.length ? ' ' + t('prodNoteOnly', { list: excl.join(', ') }) : '');
  // kerncijfers
  let n = 0, st = 0, spec = 0;
  eachUnit((f, row) => { n++; st += row[2]; if (row[7]) spec++; });
  const all = ratioBy('', null)[0] || { n: 0, k: 0 };
  $('prodKpi').innerHTML = [
    [nf.format(n), t('prodKUnits')], [nf.format(st), t('prodKStuks')], [fmtPct(spec, n), t('prodKSpec')], [fmtPct(all.k, all.n), t('prodKRatio')],
  ].map(([b, s]) => `<div class="stat"><b>${b}</b><span>${esc(s)}</span></div>`).join('');
  renderProdChart();
  renderProdTop();
}
function renderProdChart() {
  const g = prodGroupKey();
  const byY = new Map(), tot = new Map();
  eachUnit((f, row) => {
    const y = Math.floor(row[1] / 100), v = g.val(f, row), w = UI.prodMeas === 'n' ? 1 : row[2];
    let m = byY.get(y); if (!m) byY.set(y, (m = new Map())); m.set(v, (m.get(v) || 0) + w); tot.set(v, (tot.get(v) || 0) + w);
  });
  const years = [...byY.keys()].sort();
  const top = [...tot].sort((a, b) => b[1] - a[1]).map((x) => x[0]);
  const shown = top.length > 8 ? top.slice(0, 7) : top; const rest = top.slice(shown.length);
  const series = shown.map((v, i) => ({ label: g.lbl(v), color: v === PI_UNK ? OTHER_COLOR : COLORS[i % COLORS.length], vals: years.map((y) => byY.get(y).get(v) || 0) }));
  if (rest.length) series.push({ label: t('prodOther'), color: '#d1d5db', vals: years.map((y) => rest.reduce((s, v) => s + (byY.get(y).get(v) || 0), 0)) });
  const pctMode = UI.prodMeas === 'pct';
  const yTot = years.map((y, i) => series.reduce((s, x) => s + x.vals[i], 0));
  if (prodChart) prodChart.destroy();
  prodChart = new Chart($('prodChart'), {
    type: 'bar',
    data: { labels: years.map(String), datasets: series.map((s) => ({ label: s.label, data: pctMode ? s.vals.map((v, i) => pct(v, yTot[i])) : s.vals, backgroundColor: s.color, borderRadius: 2, maxBarThickness: 70 })) },
    options: { responsive: true, maintainAspectRatio: false, animation: false,
      scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, max: pctMode ? 100 : undefined, ticks: { callback: (v) => nf.format(v) + (pctMode ? '%' : '') } } },
      plugins: { legend: { display: false }, tooltip: { backgroundColor: '#111', padding: 10, callbacks: { label: (c) => ` ${c.dataset.label}: ${nf.format(c.raw)}${pctMode ? '%' : ''}` } } } },
  });
  $('prodLegend').innerHTML = series.map((s) => `<span><i style="background:${s.color}"></i>${esc(s.label)}</span>`).join('');
  $('prodChartTitle').textContent = t('prodChartTitle', { k: g.key === '_fam' ? t('prodFamily') : pkName(g.key) });
  const lastY = years[years.length - 1];
  $('prodChartNote').textContent = lastY && F.to != null && ymOf(F.to) % 100 !== 12 && Math.floor(ymOf(F.to) / 100) === lastY ? t('prodPartialYear', { y: lastY }) : '';
}
function renderProdTop() {
  const keys = PKS.keys.filter((k) => k !== 'prod' || PKS.fams.length === 1);
  if (!keys.includes(UI.prodTopKey)) UI.prodTopKey = keys.includes('motor') ? 'motor' : keys[0] || 'prod';
  const sel = $('prodTopKey');
  sel.innerHTML = keys.map((k) => `<option value="${k}"${k === UI.prodTopKey ? ' selected' : ''}>${esc(pkName(k))}</option>`).join('');
  const k = UI.prodTopKey;
  const cnt = new Map(), byLand = new Map(), landTot = new Map(); let total = 0;
  eachUnit((f, row) => {
    const v = piVal(f, row, k), w = UI.prodMeas === 'n' ? 1 : row[2];
    cnt.set(v, (cnt.get(v) || 0) + w); total += w;
    const l = row[6] || 0; landTot.set(l, (landTot.get(l) || 0) + w);
    let m = byLand.get(l); if (!m) byLand.set(l, (m = new Map())); m.set(v, (m.get(v) || 0) + w);
  });
  const entries = [...cnt].sort((a, b) => b[1] - a[1]).slice(0, 10).map(([v, nn]) => ({ key: v, label: pkLabel(v), n: nn }));
  renderBars($('prodTop'), entries, (e) => setPk(k, new Set([e.key])), total);
  // heatmap: waarde × land (% van de stuks per land)
  const lands = [...landTot].filter(([l]) => l).sort((a, b) => b[1] - a[1]).slice(0, 8).map((x) => x[0]);
  if (landTot.get(0)) lands.push(0);
  const vals = entries.map((e) => e.key);
  const landName = (l) => (l ? (PKS.fams[0].land && PKS.fams[0].land[l]) || 'Land ' + l : t('pkUnknown'));
  $('prodHeat').innerHTML = !vals.length ? '' : `<thead><tr><th>${esc(pkName(k))}</th>${lands.map((l) => `<th class="num">${esc(landName(l))}<br><small>${nf.format(landTot.get(l))}</small></th>`).join('')}</tr></thead><tbody>` +
    vals.map((v) => `<tr><td>${esc(pkLabel(v))}</td>${lands.map((l) => { const p = pct((byLand.get(l) || new Map()).get(v) || 0, landTot.get(l)); return `<td class="num heat" style="background:rgba(42,120,214,${Math.min(0.85, p / 100 * 1.2).toFixed(2)});color:${p > 55 ? '#fff' : 'inherit'}">${!p ? '' : p < 1 ? '&lt;1%' : nf.format(Math.round(p)) + '%'}</td>`; }).join('')}</tr>`).join('') + '</tbody>';
}
REPORTS.push({ id: 'prod', label: 'repProd', render: renderProdRep, avail: () => !!window.PI });
try { if (localStorage.getItem('kl_report') === 'prod') UI.report = 'prod'; } catch { /* geen opslag */ }

/* ───────────── Bediening ───────────── */
$('pkToggle').addEventListener('click', () => { UI.pkOpen = !UI.pkOpen; try { localStorage.setItem('kl_pkOpen', UI.pkOpen ? '1' : '0'); } catch { /* */ } renderPkCard(); });
$('pkClear').addEventListener('click', (e) => { e.stopPropagation(); PKF.clear(); UI.page = 0; update(); });
$('pkSplit').addEventListener('change', (e) => { UI.pkSplit = e.target.value; update(); });
$('prodTopKey').addEventListener('change', (e) => { UI.prodTopKey = e.target.value; renderProdTop(); });
$('segProdMeas').addEventListener('click', (e) => {
  const v = e.target.dataset.v; if (!v) return; UI.prodMeas = v;
  $('segProdMeas').querySelectorAll('button').forEach((b) => b.classList.toggle('on', b.dataset.v === v));
  renderProdChart(); renderProdTop();
});

/* ───────────── Admin: productdata opladen ───────────── */
let PI_PENDING = null;
const PI_NEED = new Set(PI_FAMS.flatMap((f) => f.sheets).concat(['P00_VERKOOPSTAT', 'MOTOREN', 'DEFINITIEF', 'DOEKEN', 'L1_DOEKEN', 'LAND', 'L6_LAND', 'PRODUCTEN', 'L2_PRODUCTEN', 'KLEUR_RAL', 'L8_RAL']));
function showPiInfo() {
  const el = $('prodLast'); if (!el) return;
  const c = META && META.chunks ? Object.keys(META.chunks).filter(isPiVt) : [];
  setMsg('prodLast', c.length ? esc(t('prodLast', { n: c.length, date: window.PI && PI.at ? new Date(PI.at).toLocaleString(currentUiLang() + '-BE') : '?' })) : esc(t('prodNone')));
}
async function piPrepare(files) {
  PI_PENDING = null; $('prodUploadBtn').disabled = true; setMsg('prodStatus', '');
  $('prodDrop').classList.remove('has-file');
  try {
    const sheets = {};
    // Opzoeklijsten eerst, dan de motoren (tabblad Definitief gaat voor), dan de grote bestanden.
    const list = [...files].sort((a, b) => a.size - b.size);
    for (const file of list) {
      setLoading(true, t('prodReading', { file: file.name }));
      await new Promise((r) => setTimeout(r, 30));
      const buf = await file.arrayBuffer();
      const names = XLSX.read(buf, { type: 'array', bookSheets: true }).SheetNames.filter((n) => PI_NEED.has(n.toUpperCase()));
      if (!names.length) continue;
      const wb = XLSX.read(buf, { type: 'array', dense: true, sheets: names, cellFormula: false, cellHTML: false, cellText: false, cellStyles: false });
      piFindSheets(wb, XLSX, sheets);
    }
    const miss = [];
    if (!sheets.P00_VERKOOPSTAT) miss.push('P00_VERKOOPSTAT (Verkoopstats.xlsm)');
    if (!sheets.MOTOREN) miss.push('MOTOREN / Definitief (PRODUCTINFO_OPZOEKLIJSTEN.xlsx)');
    if (!PI_FAMS.some((f) => f.sheets.some((s) => sheets[s]))) miss.push('P01…P38 (PRODUCTINFO_DATA.xlsm)');
    if (miss.length) throw new Error(t('prodMissing', { list: miss.join(', ') }));
    setLoading(true, t('prodBuilding'));
    await new Promise((r) => setTimeout(r, 30));
    const built = piBuild(sheets);
    const at = new Date().toISOString();
    const oldChunks = (META && META.available && META.chunks) || {};
    const chunks = {}, changed = [];
    let units = 0, spec = 0;
    for (const p of built.fams) {
      const vt = String(PI_VT0 + PI_FAMS.findIndex((f) => f.id === p.pi.f));
      units += p.pi.n.units; spec += p.pi.n.withSpec;
      const h = await sha(JSON.stringify(p));      // hash zonder datum: ongewijzigde families niet opnieuw opladen
      p.pi.at = at; p.pi.files = list.map((f) => f.name);
      const old = oldChunks[vt];
      chunks[vt] = { n: 0, h, bytes: old && old.h === h ? old.bytes : 0 };
      if (!old || old.h !== h) changed.push({ vt, json: JSON.stringify(p), payload: p });
    }
    const removed = Object.keys(oldChunks).filter((vt) => isPiVt(vt) && !chunks[vt]);
    PI_PENDING = { chunks, changed, nothing: !changed.length && !removed.length };
    $('prodDrop').classList.add('has-file');
    const missSheets = PI_FAMS.filter((f) => !f.sheets.some((s) => sheets[s])).map((f) => f.label);
    setMsg('prodStatus', esc(t('prodBuilt', { units: nf.format(units), fams: built.fams.length, spec: fmtPct(spec, units), chunks: changed.length, total: built.fams.length })) +
      (missSheets.length ? '<br>' + esc(t('prodNoSheet', { list: missSheets.join(', ') })) : '') + (PI_PENDING.nothing ? '<br>' + esc(t('diffNothing')) : ''), PI_PENDING.nothing ? 'ok' : '');
    $('prodUploadBtn').disabled = PI_PENDING.nothing;
  } catch (e) {
    console.error(e);
    setMsg('prodStatus', esc(t('readError', { msg: e.message })), 'err');
  } finally { setLoading(false); }
}
async function piUpload() {
  if (!PI_PENDING) return;
  const P = PI_PENDING; const btn = $('prodUploadBtn'); btn.disabled = true;
  try {
    let done = 0; const total = P.changed.length;
    setMsg('prodStatus', esc(t('uploading', { done, total })));
    await pool(P.changed, 3, async (c) => {
      const gz = await gzip(c.json);
      P.chunks[c.vt].bytes = gz.byteLength;
      for (let attempt = 1; ; attempt++) {
        try { await apiBinary('chunk_put', { vt: c.vt, h: P.chunks[c.vt].h }, gz); break; }
        catch (e) { if (attempt >= 3 || e.status === 401 || e.status === 403) throw e; await new Promise((r) => setTimeout(r, 1000 * attempt)); }
      }
      done++; setMsg('prodStatus', esc(t('uploading', { done, total })));
    });
    // Nieuwe versie = de huidige klachtenblokken (ongewijzigd) + de nieuwe productblokken.
    const { ok: mOk, data: m } = await apiRequest({ mode: 'data_meta' });
    if (!mOk || !m.available) throw new Error(t('prodNeedKlachten'));
    const chunks = {};
    for (const [vt, c] of Object.entries(m.chunks)) if (!isPiVt(vt)) chunks[vt] = c;
    Object.assign(chunks, P.chunks);
    const md = m.diff || {};
    const { ok, status, data } = await apiRequest({ mode: 'data_commit', chunks, fileName: m.fileName, fileDate: m.fileDate, diff: { new: md.new || 0, changed: md.changed || 0, removed: md.removed || 0 } });
    if (handleAuthFailure(status)) return;
    if (!ok) throw new Error(data.error || t('unknownError'));
    for (const c of P.changed) await DB.set('c:' + c.vt, { h: P.chunks[c.vt].h, rows: c.payload });
    PI_PENDING = null; $('prodDrop').classList.remove('has-file');
    setMsg('prodStatus', esc(t('prodDone')), 'ok');
    await loadFromServer();
    showPiInfo();
  } catch (e) {
    console.error(e);
    if (e.status === 401) { showLogin(t('sessionExpired')); return; }
    setMsg('prodStatus', esc(t('uploadError', { msg: e.message })), 'err');
    btn.disabled = false;
  }
}
const drop = $('prodDrop'), inp = $('prodFile');
drop.addEventListener('click', () => inp.click());
drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('drag-over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('drag-over'));
drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('drag-over'); if (e.dataTransfer.files.length) piPrepare(e.dataTransfer.files); });
inp.addEventListener('change', () => { if (inp.files.length) piPrepare(inp.files); inp.value = ''; });
$('prodUploadBtn').addEventListener('click', piUpload);
window.showPiInfo = showPiInfo;
})();
