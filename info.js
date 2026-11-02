/* Klachtenanalyse – uitleg per veld: wat betekent het en waar komt het vandaan.
 * Bron van alle data: AS400 (IBM i), ODBC-verbinding S4411711, via KLACHTEN_DATA.xlsx (opgeladen door een admin).
 */
'use strict';

const SRC = 'OORDLIB.PFKLACHTEN';
const INFO = {
  filters: {
    t: 'Filters',
    wat: 'Alle filters werken samen (EN): een klacht moet aan elk ingevuld filter voldoen. Binnen één filter met meerdere keuzes volstaat één ervan (OF). De aantallen in de keuzelijsten tellen over álle geladen klachten, niet enkel over de huidige selectie.',
    bron: ['Eén rij per klacht uit <code>' + SRC + '</code> (de klachtenmodule van WinPro, op de AS400).', 'Enkel klachten vanaf 1/1/2020 worden geladen (<code>KDATJJ &gt;= 2020</code>).'],
  },
  periode: {
    t: 'Periode (datum klacht)',
    wat: 'Toont klachten waarvan de aanmaakdatum tussen beide datums valt (grenzen inbegrepen). De snelkeuzes rekenen terug vanaf de jongste klacht in de data.',
    bron: ['<code>KDATE</code> in <code>' + SRC + '</code>: "datum ingelogd", de dag waarop de klacht aangemaakt werd (formaat JJJJMMDD).'],
    let: 'Dit is de datum van de klacht, niet van levering of productie van het product.',
  },
  pg: {
    t: 'Productgroep',
    wat: 'De productgroep die op de klacht ingevuld werd, bv. SO!, Luifels, Screens, Rolluiken, PVC-ramen. De knoppen Winsol, WinVista en WinBalu onderaan de lijst zetten in één klik een vaste set productgroepen aan.',
    bron: ['Code: <code>KPROD</code> in <code>' + SRC + '</code> (bv. 4S, 4A, 2A).', 'Omschrijving: <code>WINSOLPRO.PFPRODGRO</code> (<code>KO_PG</code> → <code>OMS_PGN</code>).'],
    let: 'Op de AS400 staat "|" waar het merk "!" schrijft; de tool toont SO! en Z!P.',
  },
  prod: {
    t: 'Product',
    wat: 'Het product (artikel) waarop de klacht betrekking heeft, bv. SO!Classic, Squaro, Solarfix.',
    bron: ['Code: <code>KARTNR</code> in <code>' + SRC + '</code> (artikelnummer, 3 cijfers).', 'Naam: <code>OARTLIB.PFPROD</code> (<code>PROD</code> → <code>PRODBN</code>, Nederlandse benaming).'],
    let: 'Klachten via de webshop of op losse onderdelen staan vaak op een verzamelproduct zoals "WEBSHOP ARTIKEL HELIOS". Het eigenlijke product is dan niet gekend; gebruik TREE of de omschrijving om verder te zoeken.',
  },
  tree: {
    t: 'TREE (defect)',
    wat: 'De keuzeboom die in WinPro (tabblad "Tree" van de klacht) ingevuld wordt: product › soort probleem › onderdeel › detail, bv. SO! › WERKING › SENSOREN › Sensor defect › Temperatuursensor. Kies je een hoger niveau, dan zitten alle onderliggende takken mee in de selectie.',
    bron: ['<code>OORDLIB.PFKLACHTT</code>: <code>TVOLGN</code> (klachtnummer), <code>TTREEID</code> (id van de tak), <code>TTREEOM</code> (volledig pad als tekst).', 'Deze tabel wordt elke nacht vanuit WinPro naar de AS400 gekopieerd.'],
    let: 'Een klacht kan meerdere TREE-paden hebben en telt dan in elk pad mee. TREE wordt pas echt gebruikt sinds 2024 (ca. 35–40% van de klachten); vóór 2024 is het grotendeels leeg. De niveaus 1 tot 3 zijn betrouwbaar; alles daarna wordt als één detailtekst getoond, omdat sommige takken zelf een "/" bevatten.',
  },
  ntr: {
    t: 'NTR-code',
    wat: 'De NTR-categorie waaronder de klacht geboekt werd: de oorzaak of verantwoordelijke, bv. 13 Garantie, 8 Productie, 7 Diversen regio, 9 Inschrijver, 12 Expeditie.',
    bron: ['<code>KNNTR</code> in <code>' + SRC + '</code>: de bevestigde NTR-code.', 'Is er nog geen bevestigde code, dan gebruikt de tool <code>KNTR</code> (vermoedelijke NTR-code). In het detail van een klacht staat dan "(vermoedelijk)".', 'Omschrijving: <code>OORDLIB.PFKLACHTN</code> (<code>NNTR</code> → <code>NOMSC</code>).'],
    let: '"NTR ontbreekt" = geen bevestigde én geen vermoedelijke code ingevuld.',
  },
  fout: {
    t: 'Foutcode',
    wat: 'Fijnere indeling binnen de NTR-code, bv. 1302 Garantie Winsol, 1301 Garantie leverancier. De eerste cijfers verwijzen naar de NTR-code.',
    bron: ['<code>KNFOUT</code> in <code>' + SRC + '</code> (bevestigd), anders <code>KFOUT</code> (vermoedelijk).', 'Omschrijving: <code>OORDLIB.PFKLACHTFO</code> en <code>OORDLIB.PFKLACHTF2</code> (<code>UFOUT</code> → <code>UOMSC</code>).'],
    let: 'Toont een foutcode "Fout 904" zonder naam, dan staat die code in geen van beide omschrijvingstabellen.',
  },
  klant: {
    t: 'Klant',
    wat: 'De klant op wiens naam de klacht staat (dealer, schrijnwerker, particulier…). Zoeken kan op naam of nummer.',
    bron: ['<code>KFIRMA</code> + <code>KKLNR</code> in <code>' + SRC + '</code>.', 'Naam en land uit het klantenbestand van die firma: W = <code>WVDJ.PFKLAN</code> (Winsol), B = <code>BADJ.PFKLAN</code> (Winbalu), H = <code>HEDJ.PFKLAN</code> (Helios); velden <code>KNAAM1</code>, <code>KNAT</code>.'],
    let: 'Een klantnummer is enkel uniek samen met de firma. Bijna alle klachten staan op firma W, ook die op Helios-producten.',
  },
  status: {
    t: 'Status',
    wat: 'De fase waarin de klacht zich nu bevindt, bv. in behandeling, materiaal klaar in magazijn, bij plaatsingsdienst, afgesloten.',
    bron: ['Code: <code>KFASE</code> in <code>' + SRC + '</code>.', 'Omschrijving: <code>OORDLIB.PFKLACHTFA</code> (<code>OFASE</code> → <code>OFASEOM</code>).'],
    let: 'Dit is de status van vandaag (op het moment van de laatste verversing), niet de status op de klachtdatum.',
  },
  vert: {
    t: 'Vertegenwoordiger',
    wat: 'De vertegenwoordiger die aan de klacht gekoppeld is.',
    bron: ['Nummer: <code>KVTNR</code> in <code>' + SRC + '</code>.', 'Naam: <code>ODIVERS.PFVERT</code> (<code>VERTNR</code> → <code>VNAAM</code>).'],
  },
  regio: {
    t: 'Regio',
    wat: 'De verkoopsregio van de klacht.',
    bron: ['Nummer: <code>KREGIO</code> in <code>' + SRC + '</code>.', 'Naam: <code>ODIVERS.PFREGIO</code> (<code>REGIO</code> → <code>REGPNN</code>).'],
    let: 'Enkele regionummers hebben geen naam in de opzoektabel; die tonen enkel het nummer.',
  },
  open: {
    t: 'Open / afgesloten',
    wat: 'Een klacht telt als <b>afgesloten</b> als de status één van deze is: A Klacht afgesloten, W Klacht geweigerd, I Klacht ter info, O Klacht ter info (afgesloten), Y Interventie geweigerd. Alle andere statussen tellen als <b>open</b>.',
    bron: ['Afgeleid van <code>KFASE</code> in <code>' + SRC + '</code>.'],
    let: 'Deze regel is in de tool zelf vastgelegd (niet in de AS400). "Klacht geweigerd door fabriek" (R) telt als open, zoals in de rapportering van Controlling.',
  },
  terinfo: {
    t: 'Klachten "ter info"',
    wat: 'Klachten met status I (Klacht ter info) of O (Klacht ter info, afgesloten) zijn meldingen zonder echte klachtbehandeling. Standaard tellen ze niet mee, net zoals in de klachtenrapportering van Controlling.',
    bron: ['Afgeleid van <code>KFASE</code> in <code>' + SRC + '</code>.'],
  },
  actie: {
    t: 'Uit te voeren actie',
    wat: 'De keuze "Uit te voeren actie?" op de klacht in WinPro: wat er moet gebeuren om de klacht op te lossen. Een klacht met actie <b>Interventie</b> telt als interventie: iemand gaat ter plaatse herstellen. <b>Plaatsen</b> is ook ter plaatse, maar om nieuw materiaal te plaatsen.',
    bron: ['<code>KLEVPL</code> in <code>' + SRC + '</code>: 1 Leveren, 2 Plaatsen, 3 Afhalen, 4 Interventie, 5 Ter info, 0 niet ingevuld.'],
    let: 'De (I) naast het klachtnummer in WinPro (veld <code>IINTVEN</code>, 1 = intern, I = via Extranet) betekent "interventie gevraagd" en staat als extra info in het detail van de klacht. Voor de cijfers telt de actie, omdat die zegt wat er gebeurt.',
  },
  vervolg: {
    t: 'Vervolg',
    wat: 'Een <b>dossier</b> is één probleem. Een <b>vervolgklacht</b> is een volgende klacht over hetzelfde probleem, binnen 1 jaar na de vorige: het probleem was dus niet in één keer opgelost. Dat is zo als de klacht (1) op het herstelorder (NTR-order) van een eerdere klacht staat, of (2) op hetzelfde order staat én dezelfde TREE heeft (niveau 3). Een <b>andere klacht op hetzelfde order</b> is een nieuw probleem op een order dat al een klacht had (ander of onbekend onderwerp): een eigen dossier.',
    bron: ['Order van de klacht: <code>KORNR</code>; herstelorder: <code>KNORDER</code>; volgorde: <code>KDATE</code>, in <code>' + SRC + '</code>.', 'Onderwerp: TREE niveau 3 uit <code>OORDLIB.PFKLACHTT</code>.'],
    let: 'Zonder TREE is het onderwerp onbekend; zo\'n klacht op hetzelfde order telt voorzichtig als "andere klacht", niet als vervolg. Korte ordernummers (1, 12, 43…) zijn regionummers en tellen niet als order. Vervolg wordt berekend over alle klachten, ook die buiten de huidige filters.',
  },
  kDos: {
    t: 'Dossiers',
    wat: 'Het aantal verschillende dossiers (unieke problemen) waartoe de klachten in de selectie behoren. Zie de uitleg bij de filter "Vervolg".',
    bron: ['Berekend in de tool uit <code>KORNR</code>, <code>KNORDER</code> en <code>KDATE</code>.'],
  },
  kAnder: {
    t: 'Andere klacht op hetzelfde order',
    wat: 'Klachten in de selectie op een order dat al een eerdere klacht had, maar over een ander (of onbekend) probleem. Ze zijn geen vervolgklacht en vormen een eigen dossier.',
    bron: ['Berekend in de tool; zie de uitleg bij de filter "Vervolg".'],
  },
  kInterv: {
    t: 'Met interventie',
    wat: 'Het aantal klachten in de selectie met uit te voeren actie "Interventie", en hun aandeel in alle klachten van de selectie.',
    bron: ['<code>KLEVPL</code> = 4 in <code>' + SRC + '</code>.'],
  },
  kVervolg: {
    t: 'Vervolgklachten',
    wat: 'Het aantal klachten in de selectie dat een vervolg is op een eerdere klacht in hetzelfde dossier (binnen 1 jaar), en hun aandeel.',
    bron: ['Berekend in de tool; zie de uitleg bij de filter "Vervolg".'],
    let: 'Recente klachten kunnen later nog een vervolg krijgen; het aandeel voor de laatste maanden stijgt daardoor nog.',
  },
  zoek: {
    t: 'Zoekterm',
    wat: 'Zoekt de ingegeven tekst (hoofdletterongevoelig, ook midden in een woord) in de vrije tekst van de klacht.',
    bron: ['<code>KOPM</code> (omschrijving klacht) en <code>KOPL</code> (verwachte oplossing) in <code>' + SRC + '</code>.', 'Oorzaak, corrigerende en correctieve maatregel uit <code>OORDLIB.PFKLACHTBH</code> (<code>BOORZA</code>, <code>BCORRI</code>, <code>BCORRE</code>).'],
    let: 'Klachten worden in het Nederlands, Frans, Duits of Engels ingegeven: zoek eventueel ook op de Franse term (bv. "moteur" naast "motor").',
  },
  kTotal: {
    t: 'Klachten',
    wat: 'Het aantal klachten dat aan alle filters voldoet. Elke klacht telt één keer, ook als ze meerdere TREE-paden heeft.',
    bron: ['Aantal unieke klachtnummers <code>KVOLGN</code> in <code>' + SRC + '</code>.'],
  },
  kOpen: {
    t: 'Open',
    wat: 'Het aantal klachten in de selectie met een open status (zie de uitleg bij "Open / afgesloten").',
    bron: ['Afgeleid van <code>KFASE</code> in <code>' + SRC + '</code>.'],
  },
  kCost: {
    t: 'NTR-kost (materiaal + plaatsing)',
    wat: 'De som van de kosten die op de klachten geboekt werden: het bedrag voor materiaal plus het bedrag voor plaatsing, in euro.',
    bron: ['<code>KNBDRB</code> ("bedrag fout", materiaal) + <code>KNWPL</code> ("plaatsingsbedrag") in <code>' + SRC + '</code>.', 'Dit zijn de bedragen van de bevestigde NTR-boeking (dezelfde bedragen staan per boeking in <code>OORDLIB.PFKLACHTNT</code>).'],
    let: 'Bedragen boven € 100.000 of negatieve bedragen tellen niet mee: dat zijn invoerfouten (bv. een ordernummer dat in het bedragveld getypt werd). Of de bedragen met of zonder btw zijn, is nog na te vragen bij Controlling.',
  },
  kLead: {
    t: 'Gemiddelde doorlooptijd tot afsluiten',
    wat: 'Voor de afgesloten klachten in de selectie: het gemiddelde aantal kalenderdagen tussen het aanmaken en het afsluiten van de klacht. Open klachten tellen niet mee.',
    bron: ['<code>KDATA</code> ("6. klacht afgesloten") − <code>KDATE</code> (datum ingelogd), beide in <code>' + SRC + '</code>.'],
    let: 'Een paar klachten in de selectie met een lange doorlooptijd kunnen het gemiddelde sterk beïnvloeden.',
  },
  kTree: {
    t: 'Met TREE ingevuld',
    wat: 'Het aandeel van de klachten in de selectie dat minstens één TREE-pad heeft.',
    bron: ['Klachten met minstens één rij in <code>OORDLIB.PFKLACHTT</code>.'],
    let: 'Hoe lager dit percentage, hoe minder representatief de TREE-grafieken en -lijsten zijn.',
  },
  evolutie: {
    t: 'Evolutie',
    wat: 'Het aantal klachten (of de NTR-kost) per maand, kwartaal of jaar, volgens de datum van de klacht. Alle filters zijn toegepast.',
    bron: ['Periode: <code>KDATE</code> in <code>' + SRC + '</code>.', 'NTR-kost: zie de uitleg bij de kerncijfers.'],
    let: 'De laatste periode is meestal nog niet volledig en ligt daarom lager.',
  },
  mode: {
    t: 'Grafiek: tijdlijn of jaren cumulatief',
    wat: '<b>Tijdlijn</b>: de gekozen periode van links naar rechts, met het aantal (of de kost) per maand, kwartaal of jaar.<br><b>Interventie &amp; vervolg</b>: per maand, kwartaal of jaar het aandeel klachten met interventie en het aandeel vervolgklachten.<br><b>Jaren cumulatief</b>: één lijn per jaar, van januari tot december, waarbij elke week (of maand) wordt opgeteld bij de vorige. Zo zie je in één oogopslag of een jaar sneller of trager oploopt dan de vorige jaren.',
    bron: ['Datum van de klacht: <code>KDATE</code> in <code>' + SRC + '</code>.'],
    let: 'In "Jaren cumulatief" wordt de periodefilter genegeerd (alle andere filters gelden wel), want elk jaar wordt volledig getoond. Het lopende jaar stopt bij de laatste klacht in de data. Weken tellen vanaf 1 januari (dag 1–7 = week 1); dit zijn geen ISO-weken. Met de vinkjes onder de grafiek kies je welke jaren getoond worden; je keuze wordt in deze browser onthouden.',
  },
  target: {
    t: 'Doel huidig jaar',
    wat: 'Tekent een gestippelde doellijn voor het lopende jaar: de cumulatieve curve van vorig jaar, verminderd met het gekozen percentage. Liggen de klachten van dit jaar onder die lijn, dan zit je op koers.',
    bron: ['Berekend uit de lijn van vorig jaar in dezelfde selectie.'],
  },
  split: {
    t: 'Uitsplitsen op',
    wat: 'Toont een lijn per waarde in plaats van één totaal. De 7 grootste waarden in de gekozen periode krijgen een eigen lijn, de rest wordt samengeteld in "Overig".',
    bron: ['Dezelfde velden als de filters met dezelfde naam.'],
    let: 'Bij TREE tellen klachten zonder TREE niet mee, en een klacht met meerdere paden telt in elk pad.',
  },
  top: {
    t: 'Waar zitten de klachten?',
    wat: 'De meest voorkomende TREE-takken, producten, NTR-codes en klanten binnen de huidige selectie. Het percentage is het aandeel in het totaal aantal klachten van de selectie. Klik op een balk om erop te filteren.',
    bron: ['TREE: <code>OORDLIB.PFKLACHTT</code>. Producten, NTR-codes en klanten: zie de filters.'],
    let: 'Bij TREE kan een klacht in meerdere takken voorkomen; de percentages tellen dan samen op tot meer dan 100%.',
  },
  lijst: {
    t: 'Klachten',
    wat: 'De klachten uit de selectie, standaard de jongste eerst. Klik op een kolomtitel om te sorteren, en op een klacht voor alle details. "Exporteren naar Excel" geeft de volledige selectie (niet enkel de getoonde pagina).',
    bron: ['Detail uit <code>' + SRC + '</code>: order (<code>KORNR</code>), NTR-order (<code>KNORDER</code>), behandeld door (<code>KBEHA</code>), datums per fase: in behandeling (<code>KDATB</code>), in magazijn (<code>KDATM</code>), bij plaatsingsdienst (<code>KDATD</code>), opgelost voor klant (<code>KDATK</code>), afgesloten (<code>KDATA</code>).', 'Oorzaak en maatregelen: <code>OORDLIB.PFKLACHTBH</code>.', '"Open in WinPro" opent de klacht in WinPro (enkel binnen het Winsol-netwerk).'],
  },
};

/* ───────────── (i)-knoppen en popover ───────────── */
const pop = document.createElement('div');
pop.id = 'infoPop'; pop.setAttribute('role', 'dialog');
document.body.appendChild(pop);
let popBtn = null;

function renderInfo(key) {
  const i = INFO[key];
  return `<h4>${i.t}<button type="button" class="x" aria-label="Sluiten">×</button></h4>
    <p>${i.wat}</p>
    ${i.bron ? `<span class="lbl">Bron</span><ul>${i.bron.map((b) => `<li>${b}</li>`).join('')}</ul>` : ''}
    ${i.let ? `<div class="warn"><b>Let op:</b> ${i.let}</div>` : ''}`;
}
function closePop() { pop.classList.remove('open'); if (popBtn) popBtn.classList.remove('on'); popBtn = null; }
function openPop(btn, key) {
  if (popBtn === btn) { closePop(); return; }
  closePop();
  pop.innerHTML = renderInfo(key);
  pop.classList.add('open'); btn.classList.add('on'); popBtn = btn;
  const r = btn.getBoundingClientRect(); const w = pop.offsetWidth;
  let left = r.left + window.scrollX - 12;
  left = Math.max(window.scrollX + 16, Math.min(left, window.scrollX + document.documentElement.clientWidth - w - 16));
  pop.style.left = left + 'px';
  pop.style.top = (r.bottom + window.scrollY + 8) + 'px';
  pop.querySelector('.x').addEventListener('click', closePop);
}
document.querySelectorAll('[data-info]').forEach((el) => {
  const key = el.dataset.info; if (!INFO[key]) return;
  const b = document.createElement('button');
  b.type = 'button'; b.className = 'ibtn'; b.textContent = 'i';
  b.title = 'Uitleg: ' + INFO[key].t; b.setAttribute('aria-label', b.title);
  b.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); openPop(b, key); });
  el.appendChild(b);
});
pop.addEventListener('click', (e) => e.stopPropagation());
document.addEventListener('click', closePop);
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closePop(); });
window.addEventListener('resize', closePop);

/* ───────────── Algemene uitleg: waar komt de data vandaan? ───────────── */
const TABLES = [
  ['OORDLIB.PFKLACHTEN', 'Klachten: 1 rij per klacht (klant, product, NTR, status, datums, kosten, teksten)', 'KLACHTEN'],
  ['OORDLIB.PFKLACHTT', 'TREE-paden per klacht (nachtelijke kopie uit WinPro)', 'TREE'],
  ['OORDLIB.PFKLACHTBH', 'Oorzaak, corrigerende en correctieve maatregel', 'OORZAAK'],
  ['OORDLIB.PFKLACHTN', 'Omschrijving NTR-codes', 'LOOKUP'],
  ['OORDLIB.PFKLACHTFO / PFKLACHTF2', 'Omschrijving foutcodes', 'LOOKUP'],
  ['OORDLIB.PFKLACHTFA', 'Omschrijving statussen', 'LOOKUP'],
  ['WINSOLPRO.PFPRODGRO', 'Omschrijving productgroepen', 'LOOKUP'],
  ['OARTLIB.PFPROD', 'Productnamen', 'LOOKUP'],
  ['ODIVERS.PFVERT', 'Namen vertegenwoordigers', 'LOOKUP'],
  ['ODIVERS.PFREGIO', 'Namen regio\'s', 'LOOKUP'],
  ['WVDJ / BADJ / HEDJ .PFKLAN', 'Klantnamen en land (Winsol / Winbalu / Helios)', 'LOOKUP'],
];
function showHelp() {
  closePop();
  $('modal').innerHTML = `<div class="help-modal">
    <div style="display:flex;align-items:flex-start;gap:10px"><h2 style="flex:1;margin:0">Waar komt de data vandaan?</h2>
      <button class="btn-outline" type="button" id="hClose">Sluiten</button></div>
    <h3>De keten</h3>
    <div class="flow">
      <span class="box"><b>WinPro</b><br>klachtenmodule</span><span class="arr">→</span>
      <span class="box"><b>AS400</b> (IBM i)<br>bibliotheek OORDLIB e.a.</span><span class="arr">→ ODBC <code>S4411711</code> →</span>
      <span class="box"><b>KLACHTEN_DATA.xlsx</b><br>4 queries (Power Query)</span><span class="arr">→</span>
      <span class="box"><b>Admin</b><br>laadt het bestand op in de tool</span><span class="arr">→</span>
      <span class="box"><b>Cloudflare</b><br>beveiligde opslag</span><span class="arr">→</span>
      <span class="box"><b>Deze pagina</b></span>
    </div>
    <p style="font-size:0.82rem;color:#555;margin-top:10px">Klachten worden in WinPro ingegeven en opgeslagen op de AS400. Bij "Alles vernieuwen" in KLACHTEN_DATA.xlsx haalt Excel de gegevens via ODBC op (met de eigen AS400-login van de admin). De admin laadt dat bestand wekelijks op in deze tool; enkel nieuwe en gewijzigde klachten worden bijgewerkt. De data staat afgeschermd in Cloudflare en is enkel na aanmelden zichtbaar; verkopers zien enkel de klachten van hun eigen VT-nummer(s). Een kopie wordt in de browser van deze pc bewaard, zodat de tool de volgende keer sneller opent. De SQL van de 4 queries staat in <b>DATA_OPHALEN.md</b>.</p>
    <h3>Gebruikte tabellen</h3>
    <table><tr><th>AS400-tabel</th><th>Inhoud</th><th>Query</th></tr>
      ${TABLES.map((t) => `<tr><td><code>${t[0]}</code></td><td>${t[1]}</td><td>${t[2]}</td></tr>`).join('')}
    </table>
    <h3>Regels die in de tool zelf vastliggen</h3>
    <ul style="font-size:0.82rem;margin:0;padding-left:18px">
      <li>Enkel klachten vanaf 1/1/2020.</li>
      <li>Open / afgesloten volgens de status (afgesloten = A, W, I, O, Y).</li>
      <li>Klachten "ter info" (status I en O) tellen standaard niet mee.</li>
      <li>Interventie = uit te voeren actie "Interventie" (KLEVPL 4).</li>
      <li>Vervolgklacht = zelfde probleem binnen 1 jaar: op het herstelorder van een vorige klacht, of op hetzelfde order met dezelfde TREE (niveau 3). Anders: andere klacht op hetzelfde order.</li>
      <li>NTR- en foutcode: de bevestigde code, anders de vermoedelijke.</li>
      <li>Bedragen boven € 100.000 of negatief tellen niet mee (invoerfouten).</li>
      <li>TREE: niveau 1 tot 3, de rest als detailtekst; dubbele paden per klacht worden weggelaten.</li>
    </ul>
    <p style="font-size:0.78rem;color:#6b7280;margin-top:14px">Klik op een <span class="ibtn" style="cursor:default">i</span> naast een veld voor de uitleg van dat veld.</p>
  </div>`;
  $('modalBg').classList.add('open');
  $('hClose').onclick = closeDetail;
}
$('helpBtn').addEventListener('click', showHelp);
