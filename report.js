/* Klachtenanalyse – PDF-rapport van de huidige selectie (filters zoals ze nu staan).
 * Gebruikt jsPDF (lib/jspdf.umd.min.js) en de grafieken van de pagina zelf.
 * Alles gebeurt in de browser: er wordt niets naar een server gestuurd.
 */
'use strict';

const PDF = { W: 210, H: 297, M: 14 }; // A4 staand, marges in mm
const PDF_YELLOW = [255, 212, 0];
const PDF_GREY = [107, 114, 128];

// Helvetica in jsPDF kent enkel WinAnsi-tekens: vervang wat het niet kan tonen.
function pdfTxt(s) {
  return String(s ?? '')
    .replace(/[−‒–—]/g, '-').replace(/→/g, '->').replace(/↳/g, '>')
    .replace(/[‘’]/g, "'").replace(/[   ]/g, ' ')
    .replace(/[^\u0000-ÿ€…“”„•›‹]/g, '?');
}

// Rendert de grafiek in een bepaalde modus, neemt er een beeld van en zet de weergave terug.
function chartImage(mode) {
  const keep = UI.mode;
  UI.mode = mode;
  renderTrend();
  const c = $('trend');
  // witte achtergrond (PNG is anders transparant)
  const tmp = document.createElement('canvas'); tmp.width = c.width; tmp.height = c.height;
  const g = tmp.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, tmp.width, tmp.height); g.drawImage(c, 0, 0);
  const lg = $('legend').cloneNode(true); lg.querySelectorAll('button, .yr-lbl, label.off').forEach((x) => x.remove());
  const legend = [...lg.childNodes].map((x) => x.textContent.trim()).filter(Boolean).join(' · ');
  const img = { data: tmp.toDataURL('image/png'), w: tmp.width, h: tmp.height, legend };
  UI.mode = keep;
  renderTrend();
  return img;
}

async function makePdfReport() {
  if (!DATA || !window.jspdf) return;
  const btn = $('pdfBtn'); btn.disabled = true; const label = btn.textContent; btn.textContent = t('pdfBusy');
  await new Promise((r) => setTimeout(r, 20));
  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: 'mm', format: 'a4' });
    const { W, H, M } = PDF; const CW = W - 2 * M;
    let y = M;
    const now = new Date();
    const loc = currentUiLang() + '-BE';
    const newPage = () => { doc.addPage(); y = M; };
    const need = (h) => { if (y + h > H - 16) newPage(); };
    const text = (s, x, yy, o) => doc.text(pdfTxt(s), x, yy, o);
    const h2 = (s) => {
      need(14);
      doc.setFillColor(...PDF_YELLOW); doc.circle(M + 2.2, y + 2.6, 2.2, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(12); doc.setTextColor(17); text(s, M + 7, y + 4);
      y += 9;
    };
    const para = (s, size = 9, color = 60) => {
      doc.setFont('helvetica', 'normal'); doc.setFontSize(size); doc.setTextColor(color);
      const lines = doc.splitTextToSize(pdfTxt(s), CW);
      need(lines.length * size * 0.42 + 1);
      doc.text(lines, M, y + 3); y += lines.length * size * 0.42 + 2;
    };

    /* ── Kop ── */
    doc.setFillColor(...PDF_YELLOW); doc.rect(0, 0, W, 3, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(20); doc.setTextColor(17);
    text('Winsol – ' + t('pdfTitle'), M, y + 8);
    y += 13;
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.setTextColor(...PDF_GREY);
    const maxD = DATA.rows.reduce((m, r) => Math.max(m, r.d), 0);
    text(t('pdfMeta', {
      date: now.toLocaleDateString(loc) + ' ' + now.toLocaleTimeString(loc, { hour: '2-digit', minute: '2-digit' }),
      user: ME ? ME.username : '', data: DATA.uploadedAt ? new Date(DATA.uploadedAt).toLocaleDateString(loc) : '–', last: fmtDate(maxD),
    }), M, y);
    y += 4;
    if (ME && ['sales', 'dir'].includes(ME.role)) { text(t('dataScope', { vts: ME.vts.join(', ') }), M, y); y += 4; }
    y += 3;

    /* ── Filters ── */
    h2(t('pdfFilters'));
    const chips = (typeof CHIPS !== 'undefined' ? CHIPS : []);
    doc.setFontSize(9);
    if (!chips.length) para(t('pdfNoFilters'));
    for (const c of chips) {
      doc.setFont('helvetica', 'bold'); doc.setTextColor(17);
      const lines = doc.splitTextToSize(pdfTxt('• ' + c.label), CW);
      need(lines.length * 4); doc.text(lines, M, y + 3); y += lines.length * 4;
      // volledige lijst als de samenvatting die niet toont (max. 12 waarden)
      let vals = c.tip ? c.tip.split('\n') : [], pre = '';
      if (vals[0] === t('chipExceptTip')) { pre = vals.shift() + ' '; }
      if (vals.length && !vals.every((v) => c.label.includes(v))) {
        const shown = pre + vals.slice(0, 12).join(', ') + (vals.length > 12 ? ', … (+' + (vals.length - 12) + ')' : '');
        doc.setFont('helvetica', 'normal'); doc.setTextColor(...PDF_GREY);
        const l2 = doc.splitTextToSize(pdfTxt(shown), CW - 4); need(l2.length * 3.6); doc.text(l2, M + 4, y + 2.6); y += l2.length * 3.6 + 0.6;
      }
    }
    y += 3;

    /* ── Kerncijfers ── */
    h2(t('pdfKpi'));
    const kpis = ['kTotal', 'kOpen', 'kCost', 'kLead', 'kTree', 'kDos', 'kInterv', 'kVervolg', 'kAnder'].map((k) => ({ v: $(k).innerText.replace(/\s+/g, ' ').trim(), l: t('lbl_' + k) }));
    const cols = 3, bw = (CW - (cols - 1) * 3) / cols, bh = 17;
    need(Math.ceil(kpis.length / cols) * (bh + 3));
    kpis.forEach((k, i) => {
      const x = M + (i % cols) * (bw + 3), yy = y + Math.floor(i / cols) * (bh + 3);
      doc.setDrawColor(225); doc.setFillColor(250, 250, 250); doc.roundedRect(x, yy, bw, bh, 2, 2, 'FD');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.setTextColor(17); text(k.v, x + 3, yy + 7);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(...PDF_GREY);
      doc.text(doc.splitTextToSize(pdfTxt(k.l), bw - 6), x + 3, yy + 11.5);
    });
    y += Math.ceil(kpis.length / cols) * (bh + 3) + 3;

    /* ── Grafieken ── */
    const charts = [
      { mode: 'time', title: t('modeTime') + (UI.split ? ' – ' + $('splitBy').selectedOptions[0].text : '') + ' (' + (UI.measure === 'cost' ? t('ntrCost') : t('count')) + ')' },
      { mode: 'qual', title: t('modeQual') },
      { mode: 'cumul', title: t('modeCumul') + ' (' + (UI.measure === 'cost' ? t('ntrCost') : t('count')) + ')' },
    ];
    let first = true;
    for (const c of charts) {
      const img = chartImage(c.mode);
      const ih = (CW * img.h) / img.w;
      if (first) { h2(t('lbl_evolutie')); first = false; }
      need(ih + 18);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(17); text(c.title, M, y + 3); y += 5;
      doc.addImage(img.data, 'PNG', M, y, CW, ih, undefined, 'FAST'); y += ih + 1;
      if (img.legend) { para(img.legend, 7.5, 110); }
      y += 3;
    }

    /* ── Top-lijsten ── */
    newPage();
    h2(t('lbl_top'));
    const total = filtered.length;
    const lv = UI.treeLevel;
    const lists = [
      { title: 'TREE (' + $('treeLevel').selectedOptions[0].text + ')', rows: topBy((r) => { const s = new Set(r.tree.map((p) => treeNodes(p)[Math.min(lv, treeNodes(p).length) - 1])); return [...s].map((n) => [n, n]); }, 12) },
      { title: t('topProd'), rows: topBy((r) => [[String(r.prod), r.prodNaam]], 12) },
      { title: t('topNtr'), rows: topBy((r) => [[String(r.ntr), ntrLabel(r)]], 12) },
      { title: t('topKlant'), rows: topBy((r) => [[r.klant, `${r.klantNaam || t('nameUnknown')} (${r.klnr})`]], 12) },
    ];
    const colW = (CW - 6) / 2;
    for (let i = 0; i < lists.length; i += 2) {
      const pair = lists.slice(i, i + 2);
      const rowsN = Math.max(...pair.map((l) => l.rows.length), 1);
      need(8 + rowsN * 5.4);
      const y0 = y;
      pair.forEach((l, j) => {
        const x = M + j * (colW + 6); let yy = y0;
        doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(17); text(l.title, x, yy + 3); yy += 6;
        const max = l.rows.length ? l.rows[0].n : 1;
        if (!l.rows.length) { doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(...PDF_GREY); text(t('noData'), x, yy + 3); }
        for (const r of l.rows) {
          const bwid = Math.max(1, ((colW - 22) * r.n) / max);
          doc.setFillColor(255, 236, 153); doc.rect(x, yy, bwid, 4.2, 'F');
          doc.setFont('helvetica', 'normal'); doc.setFontSize(7.2); doc.setTextColor(30);
          let lab = pdfTxt(r.label); while (doc.getTextWidth(lab) > colW - 24 && lab.length > 4) lab = lab.slice(0, -2);
          if (lab !== pdfTxt(r.label)) lab += '…';
          doc.text(lab, x + 1, yy + 3.1);
          doc.setTextColor(...PDF_GREY);
          doc.text(nf.format(r.n) + (total ? '  ' + Math.round((r.n / total) * 100) + '%' : ''), x + colW, yy + 3.1, { align: 'right' });
          yy += 5.4;
        }
      });
      y = y0 + 8 + rowsN * 5.4 + 4;
    }

    /* ── Klachtenlijst (eerste 40 volgens de huidige sortering) ── */
    const col = COLS.find((c) => c.k === UI.sort.key) || COLS[1];
    const sv = col.s || col.f;
    const sorted = filtered.slice().sort((a, b) => { const x = sv(a), z = sv(b); return (x < z ? -1 : x > z ? 1 : 0) * UI.sort.dir; });
    const N = Math.min(40, sorted.length);
    need(30);
    h2(t('lbl_lijst'));
    para(t('pdfListNote', { n: N, total: nf.format(sorted.length) }), 8, 110);
    const tc = [
      { h: t('colId'), w: 16, f: (r) => r.id },
      { h: t('colDate'), w: 17, f: (r) => fmtDate(r.d) },
      { h: t('lbl_klant'), w: 38, f: (r) => r.klantNaam || r.klant },
      { h: t('lbl_prod'), w: 36, f: (r) => r.prodNaam },
      { h: 'NTR', w: 30, f: (r) => ntrLabel(r) },
      { h: t('lbl_status'), w: 18, f: (r) => (r.open ? t('open') : t('closed')) },
      { h: t('ntrCost'), w: 0, f: (r) => (r.cost ? eur.format(r.cost) : ''), right: true },
    ];
    tc[tc.length - 1].w = CW - tc.slice(0, -1).reduce((s, c) => s + c.w, 0);
    const header = () => {
      doc.setFillColor(245, 245, 246); doc.rect(M, y, CW, 5.5, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(7.5); doc.setTextColor(60);
      let x = M; for (const c of tc) { doc.text(pdfTxt(c.h), c.right ? x + c.w - 1 : x + 1, y + 3.8, c.right ? { align: 'right' } : undefined); x += c.w; }
      y += 6;
    };
    header();
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7.2);
    for (const r of sorted.slice(0, N)) {
      if (y + 5 > H - 16) { newPage(); header(); doc.setFont('helvetica', 'normal'); doc.setFontSize(7.2); }
      let x = M; doc.setTextColor(30);
      for (const c of tc) {
        let s = pdfTxt(c.f(r)); while (doc.getTextWidth(s) > c.w - 2 && s.length > 3) s = s.slice(0, -2);
        doc.text(s, c.right ? x + c.w - 1 : x + 1, y + 3.4, c.right ? { align: 'right' } : undefined); x += c.w;
      }
      doc.setDrawColor(235); doc.line(M, y + 5, M + CW, y + 5);
      y += 5;
    }

    /* ── Voettekst op elke pagina ── */
    const n = doc.getNumberOfPages();
    for (let i = 1; i <= n; i++) {
      doc.setPage(i);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(...PDF_GREY);
      text(t('pdfSource'), M, H - 8);
      text(t('pdfPage', { p: i, n }), W - M, H - 8, { align: 'right' });
    }
    doc.save(`${t('pdfFile')}_${now.toISOString().slice(0, 10)}.pdf`);
  } catch (e) {
    console.error(e);
    setStatus(t('pdfError', { msg: e.message }), true);
  } finally {
    btn.disabled = false; btn.textContent = label;
  }
}
$('pdfBtn').addEventListener('click', makePdfReport);
