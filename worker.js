// ============================================================
// Winsol Klachtenanalyse — Cloudflare Worker
// ============================================================
//
// Zelfde patroon als de Presentatievertaler en Feedbackloop: login en rollen
// in Workers KV, de frontend (GitHub Pages) doet het rekenwerk in de browser.
//
// De klachtendata wordt door een admin opgeladen: de browser van de admin leest
// KLACHTEN_DATA.xlsx, berekent alle velden (incl. dossiers/vervolgklachten op de
// VOLLEDIGE data), vergelijkt met wat al online staat en stuurt enkel de
// gewijzigde stukken. De data staat in KV, opgesplitst per vertegenwoordiger
// (één gzip-blok per VT), zodat:
//   - een verkoper (rol sales/dir) enkel de blokken van zijn eigen VT-nummers krijgt;
//   - de Worker de data nooit hoeft te ontleden (weinig rekentijd, gratis plan volstaat).
//
// KV-sleutels:
//   user:<username>     -> { username, passwordHash, role, vts:[..], regio }
//   session:<token>     -> { username }                       (met vervaldatum)
//   fail:user:<naam> / fail:ip:<ip>                           (login-blokkering)
//   data:meta           -> { version, uploadedAt, uploadedBy, fileName, fileDate,
//                            rows, chunks:{ <vt>: { n, h, bytes } }, diff:{new,changed,removed} }
//   data:c:<vt>:<hash>  -> gzip(JSON-array met de klachten van die VT)   (binair)
//                          actief zodra data:meta naar die hash verwijst
//   data:diff           -> gzip(JSON { at, new:[ids], changed:[ids], removed:[ids] })
//                          = nieuwe/gewijzigde klachten van de laatste upload (input voor de AI-analyse, later)
//
// Secrets (npx wrangler secret put ...):
//   PASSWORD_PEPPER     — willekeurige lange tekst, voor de wachtwoord-hashes
//   ANTHROPIC_API_KEY   — (later, voor de AI-analyse; nu nog niet nodig)
// Optionele variabele (wrangler.toml [vars]):
//   ALLOWED_ORIGIN      — standaard https://winsol-belgie.github.io ("*" om lokaal te testen)

const DEFAULT_ALLOWED_ORIGIN = 'https://winsol-belgie.github.io';
let ALLOWED_ORIGIN = DEFAULT_ALLOWED_ORIGIN;

const ROLES = ['admin', 'user', 'sales', 'dir'];
const SCOPED_ROLES = ['sales', 'dir']; // zien enkel de klachten van hun VT-nummers
const DELETE_ROLES = ['delete', 'verwijderen', 'verwijder', 'supprimer', 'löschen', 'eliminar', 'elimina'];
const MAX_CHUNK_BYTES = 24 * 1024 * 1024; // KV-limiet is 25 MiB per waarde

export default {
  async fetch(request, env) {
    ALLOWED_ORIGIN = env.ALLOWED_ORIGIN || DEFAULT_ALLOWED_ORIGIN;
    if (request.method === 'OPTIONS') return new Response(null, { headers: corsHeaders() });
    if (request.method !== 'POST') return jsonResponse({ error: 'Method not allowed' }, 405);

    const url = new URL(request.url);
    const binMode = url.searchParams.get('mode');

    // ---- Binaire routes (datablokken): body = gzip-bytes, de Worker ontleedt ze niet ----
    if (binMode === 'chunk_get' || binMode === 'chunk_put' || binMode === 'diff_put') {
      const session = await getSession(request, env);
      if (!session) return jsonResponse({ error: 'Niet ingelogd of sessie verlopen.' }, 401);

      if (binMode === 'chunk_get') {
        const vt = cleanVt(url.searchParams.get('vt'));
        if (vt === null) return jsonResponse({ error: 'Ongeldig VT-nummer.' }, 400);
        if (!mayReadVt(session, vt)) return jsonResponse({ error: 'Geen toegang tot deze data.' }, 403);
        const meta = await env.KLACHTEN_KV.get('data:meta', 'json');
        const c = meta && meta.chunks && meta.chunks[vt];
        const stream = c ? await env.KLACHTEN_KV.get(chunkKey(vt, c.h), 'stream') : null;
        if (!stream) return jsonResponse({ error: 'Datablok niet gevonden.' }, 404);
        return new Response(stream, { headers: { 'content-type': 'application/octet-stream', 'cache-control': 'no-store', ...corsHeaders() } });
      }

      if (session.role !== 'admin') return jsonResponse({ error: 'Enkel toegankelijk voor admins.' }, 403);
      const buf = await request.arrayBuffer();
      if (!buf.byteLength || buf.byteLength > MAX_CHUNK_BYTES) return jsonResponse({ error: 'Datablok leeg of te groot.' }, 413);
      if (binMode === 'diff_put') {
        // lijst nieuwe/gewijzigde klachten (gzip-JSON), input voor de latere AI-analyse
        await env.KLACHTEN_KV.put('data:diff', buf);
        return jsonResponse({ ok: true });
      }
      // chunk_put: opgeslagen onder VT + inhouds-hash. Pas actief zodra data_commit
      // de nieuwe hash in data:meta zet; tot dan blijft de vorige versie zichtbaar.
      const vt = cleanVt(url.searchParams.get('vt'));
      const h = String(url.searchParams.get('h') || '');
      if (vt === null || !/^[0-9a-f]{8,64}$/.test(h)) return jsonResponse({ error: 'Ongeldig VT-nummer of hash.' }, 400);
      await env.KLACHTEN_KV.put(chunkKey(vt, h), buf);
      return jsonResponse({ ok: true, vt, bytes: buf.byteLength });
    }

    // ---- JSON-routes ----
    let body;
    try { body = await request.json(); } catch { return jsonResponse({ error: 'Ongeldige request body.' }, 400); }

    if (body.mode === 'login') return handleLogin(body, env, request);
    if (body.mode === 'logout') return handleLogout(request, env);

    const session = await getSession(request, env);
    if (!session) return jsonResponse({ error: 'Niet ingelogd of sessie verlopen.' }, 401);

    if (body.mode === 'auth_me') return jsonResponse(publicUser(session));

    if (body.mode === 'data_meta') {
      const meta = await env.KLACHTEN_KV.get('data:meta', 'json');
      if (!meta) return jsonResponse({ available: false });
      const chunks = {};
      for (const [vt, c] of Object.entries(meta.chunks || {})) if (mayReadVt(session, vt)) chunks[vt] = c;
      return jsonResponse({
        available: true, version: meta.version, uploadedAt: meta.uploadedAt, fileName: meta.fileName, fileDate: meta.fileDate,
        rows: Object.values(chunks).reduce((t, c) => t + c.n, 0), totalRows: meta.rows, chunks,
        ...(session.role === 'admin' ? { uploadedBy: meta.uploadedBy, diff: meta.diff } : {}),
      });
    }

    // --- Admin-only ---
    if (session.role !== 'admin') return jsonResponse({ error: 'Enkel toegankelijk voor admins.' }, 403);

    if (body.mode === 'data_commit') return handleDataCommit(body, env, session);
    if (body.mode === 'admin_upsert_users') return handleAdminUpsertUsers(body, env, session);
    if (body.mode === 'admin_list_users') return handleListUsers(env);

    return jsonResponse({ error: 'Onbekende actie.' }, 400);
  },
};

// ============================================================
// Data
// ============================================================

function cleanVt(v) {
  const s = String(v ?? '').trim();
  return /^\d{1,6}$/.test(s) ? String(Number(s)) : null;
}

function mayReadVt(session, vt) {
  if (!SCOPED_ROLES.includes(session.role)) return true;
  return (session.vts || []).includes(String(vt));
}

const chunkKey = (vt, h) => `data:c:${vt}:${h}`;

// Na het opladen van alle gewijzigde blokken maakt "commit" de nieuwe versie in
// één keer actief (één schrijfactie op data:meta) en ruimt oude blokken op.
async function handleDataCommit(body, env, session) {
  const chunks = body.chunks && typeof body.chunks === 'object' ? body.chunks : null;
  if (!chunks || !Object.keys(chunks).length) return jsonResponse({ error: 'Geen datablokken opgegeven.' }, 400);
  const clean = {};
  for (const [k, c] of Object.entries(chunks)) {
    const vt = cleanVt(k);
    const h = String((c && c.h) || '');
    if (vt === null || !/^[0-9a-f]{8,64}$/.test(h)) return jsonResponse({ error: `Ongeldig datablok (${k}).` }, 400);
    clean[vt] = { n: Math.max(0, Number(c.n) || 0), h, bytes: Math.max(0, Number(c.bytes) || 0) };
  }
  // Bestaan alle blokken? (één list-aanroep per 1000 sleutels)
  const existing = new Set();
  let cursor;
  do {
    const page = await env.KLACHTEN_KV.list({ prefix: 'data:c:', cursor });
    for (const k of page.keys) existing.add(k.name);
    cursor = page.list_complete ? null : page.cursor;
  } while (cursor);
  const missing = Object.entries(clean).filter(([vt, c]) => !existing.has(chunkKey(vt, c.h))).map(([vt]) => vt);
  if (missing.length) return jsonResponse({ error: `Datablokken ontbreken (VT ${missing.slice(0, 10).join(', ')}). Laad opnieuw op.` }, 409);

  const d = body.diff || {};
  const cnt = (v) => Math.max(0, Number(v) || 0);
  const meta = {
    version: Date.now(),
    uploadedAt: new Date().toISOString(),
    uploadedBy: session.username,
    fileName: String(body.fileName || '').slice(0, 200),
    fileDate: Number(body.fileDate) || null,
    rows: Object.values(clean).reduce((t, c) => t + c.n, 0),
    chunks: clean,
    diff: { new: cnt(d.new), changed: cnt(d.changed), removed: cnt(d.removed) },
  };
  await env.KLACHTEN_KV.put('data:meta', JSON.stringify(meta));

  // Oude blokken (niet meer in gebruik) verwijderen.
  const keep = new Set(Object.entries(clean).map(([vt, c]) => chunkKey(vt, c.h)));
  let removed = 0;
  for (const k of existing) {
    if (!keep.has(k) && removed < 800) { await env.KLACHTEN_KV.delete(k); removed++; }
  }
  return jsonResponse({ ok: true, version: meta.version, rows: meta.rows, cleaned: removed, diff: meta.diff });
}

// ============================================================
// Auth — zelfde werking als Presentatievertaler / Feedbackloop
// ============================================================
// Wachtwoorden worden nooit leesbaar bewaard: sha256(wachtwoord + PASSWORD_PEPPER).

const SESSION_TTL_REMEMBER = 60 * 60 * 24 * 30; // 30 dagen ("blijf aangemeld")
const SESSION_TTL_DEFAULT = 60 * 60 * 12; // 12 uur
const MAX_FAILS_USER = 5;
const MAX_FAILS_IP = 20;
const FAIL_WINDOW = 60 * 15; // seconden
const MIN_PASSWORD_LENGTH = 8;

async function sha256Hex(str) {
  const data = new TextEncoder().encode(str);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  return [...new Uint8Array(hashBuffer)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

function randomToken() {
  const bytes = crypto.getRandomValues(new Uint8Array(32));
  return [...bytes].map((b) => b.toString(16).padStart(2, '0')).join('');
}

function timingSafeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string' || a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function publicUser(u) {
  return { username: u.username, role: u.role, vts: u.vts || [], regio: u.regio || '' };
}

async function getSession(request, env) {
  const token = request.headers.get('X-Auth-Token') || '';
  if (!token) return null;
  const session = await env.KLACHTEN_KV.get(`session:${token}`, 'json');
  if (!session) return null;
  // Rol en VT-nummers altijd uit het actuele gebruikersrecord: wijzigingen gelden meteen.
  const user = await env.KLACHTEN_KV.get(`user:${session.username}`, 'json');
  if (!user) { await env.KLACHTEN_KV.delete(`session:${token}`); return null; }
  return publicUser(user);
}

async function readFails(env, key) {
  const v = await env.KLACHTEN_KV.get(key);
  return v ? parseInt(v, 10) || 0 : 0;
}
async function addFail(env, key) {
  const n = (await readFails(env, key)) + 1;
  await env.KLACHTEN_KV.put(key, String(n), { expirationTtl: FAIL_WINDOW });
  return n;
}

async function handleLogin(body, env, request) {
  const username = String(body.username || '').trim();
  const password = String(body.password || '');
  const remember = !!body.remember;
  if (!username || !password) return jsonResponse({ error: 'Gebruikersnaam en wachtwoord verplicht.' }, 400);
  if (!env.PASSWORD_PEPPER) return jsonResponse({ error: 'PASSWORD_PEPPER ontbreekt (wrangler secret put PASSWORD_PEPPER).' }, 500);

  const ip = request.headers.get('CF-Connecting-IP') || 'onbekend';
  const failUserKey = `fail:user:${username.toLowerCase()}`;
  const failIpKey = `fail:ip:${ip}`;
  const [failsUser, failsIp] = await Promise.all([readFails(env, failUserKey), readFails(env, failIpKey)]);
  if (failsUser >= MAX_FAILS_USER || failsIp >= MAX_FAILS_IP) {
    return jsonResponse({ error: 'Te veel mislukte pogingen. Probeer het binnen 15 minuten opnieuw.', code: 'too_many_attempts' }, 429);
  }
  const user = await env.KLACHTEN_KV.get(`user:${username}`, 'json');
  const hash = await sha256Hex(password + env.PASSWORD_PEPPER); // ook bij onbekende gebruiker (zelfde responstijd)
  if (!user || !timingSafeEqual(hash, user.passwordHash)) {
    await Promise.all([addFail(env, failUserKey), addFail(env, failIpKey)]);
    return jsonResponse({ error: 'Onbekende gebruikersnaam of fout wachtwoord.' }, 401);
  }
  if (failsUser) await env.KLACHTEN_KV.delete(failUserKey);
  const token = randomToken();
  await env.KLACHTEN_KV.put(`session:${token}`, JSON.stringify({ username: user.username }), { expirationTtl: remember ? SESSION_TTL_REMEMBER : SESSION_TTL_DEFAULT });
  return jsonResponse({ token, ...publicUser(user) });
}

async function handleLogout(request, env) {
  const token = request.headers.get('X-Auth-Token') || '';
  if (token) await env.KLACHTEN_KV.delete(`session:${token}`);
  return jsonResponse({ ok: true });
}

// Admin laadt de gebruikers-Excel op. Kolommen: USER, PW, ROLE, VTNR, Regio.
//   ROLE admin / user   -> ziet alle data (admin kan ook opladen)
//   ROLE sales / dir    -> ziet enkel klachten van de VT-nummers in VTNR (bv. "525" of "525;755")
//   ROLE delete         -> gebruiker verwijderen
// Gebruikers die niet in het bestand staan, blijven ongewijzigd.
async function handleAdminUpsertUsers(body, env, session) {
  if (!env.PASSWORD_PEPPER) return jsonResponse({ error: 'PASSWORD_PEPPER ontbreekt (wrangler secret put PASSWORD_PEPPER).' }, 500);
  const users = Array.isArray(body.users) ? body.users : [];
  if (!users.length) return jsonResponse({ error: 'Geen gebruikers ontvangen.' }, 400);

  let updated = 0, deleted = 0;
  const errors = [];
  for (const raw of users) {
    const username = String(raw.username || '').trim();
    const password = String(raw.password || '').trim();
    const role = String(raw.role || '').trim().toLowerCase();
    const vts = String(raw.vtnr ?? '').split(/[;,\s/]+/).map(cleanVt).filter((v) => v !== null);
    const regio = String(raw.regio || '').trim().slice(0, 40);
    if (!username) { errors.push('(geen gebruikersnaam) — overgeslagen.'); continue; }
    if (DELETE_ROLES.includes(role)) {
      if (username === session.username) { errors.push(`${username}: je kan jezelf niet verwijderen — overgeslagen.`); continue; }
      await env.KLACHTEN_KV.delete(`user:${username}`);
      deleted++;
      continue;
    }
    if (!ROLES.includes(role)) { errors.push(`${username}: onbekende rol "${raw.role || ''}" (verwacht admin, user, sales, dir of delete) — overgeslagen.`); continue; }
    if (SCOPED_ROLES.includes(role) && !vts.length) { errors.push(`${username}: rol ${role} heeft minstens één VT-nummer nodig in kolom VTNR — overgeslagen.`); continue; }
    if (password.length < MIN_PASSWORD_LENGTH) { errors.push(`${username}: wachtwoord ontbreekt of is korter dan ${MIN_PASSWORD_LENGTH} tekens — overgeslagen.`); continue; }
    if (username === session.username && role !== 'admin') { errors.push(`${username}: je kan je eigen admin-rechten niet afnemen — overgeslagen.`); continue; }
    const passwordHash = await sha256Hex(password + env.PASSWORD_PEPPER);
    await env.KLACHTEN_KV.put(`user:${username}`, JSON.stringify({ username, passwordHash, role, vts: SCOPED_ROLES.includes(role) ? vts : [], regio }));
    updated++;
  }
  return jsonResponse({ updated, deleted, total: users.length, errors });
}

async function handleListUsers(env) {
  const out = [];
  let cursor;
  do {
    const page = await env.KLACHTEN_KV.list({ prefix: 'user:', cursor });
    for (const k of page.keys) {
      const u = await env.KLACHTEN_KV.get(k.name, 'json');
      if (u) out.push(publicUser(u));
    }
    cursor = page.list_complete ? null : page.cursor;
  } while (cursor);
  return jsonResponse({ users: out.sort((a, b) => a.username.localeCompare(b.username)) });
}

// ============================================================
// Helpers
// ============================================================
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, X-Auth-Token',
  };
}
function jsonResponse(obj, status = 200) {
  return new Response(JSON.stringify(obj), { status, headers: { 'content-type': 'application/json', ...corsHeaders() } });
}
