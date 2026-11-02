// ============================================================
// Winsol Klachtenanalyse — taal van de app (UI)
// ============================================================
// Zelfde werkwijze als de Presentatievertaler en Socrates: een keuzelijst met
// "Automatisch" (= taal van de browser) en zes talen; keuze per browser onthouden.
//
// Gebruik in index.html:  data-i18n="key" (textContent), data-i18n-html="key",
//                         data-i18n-placeholder="key", data-i18n-title="key"
// Gebruik in app.js:      t('key', { var: waarde })
//
// Ontbreekt een sleutel in een taal, dan toont de app de Nederlandse tekst.
// Inhoud uit de AS400 (productnamen, NTR-omschrijvingen, TREE…) wordt niet vertaald.

const UI_LANGS = { nl: 'Nederlands', fr: 'Français', en: 'English', de: 'Deutsch', es: 'Español', it: 'Italiano' };
const UILANG_KEY = 'klachtenanalyse_uilang';

const I18N = {
  nl: {
    appTitle: 'Klachtenanalyse',
    uiLanguage: 'Taal van de app', auto: 'Automatisch',
    loginSub: 'Log in om verder te gaan.', username: 'Gebruikersnaam', password: 'Wachtwoord', pwToggle: 'Wachtwoord tonen/verbergen',
    remember: 'Blijf aangemeld', loginBtn: 'Aanmelden', loggingIn: 'Aanmelden…', loginMissing: 'Vul gebruikersnaam en wachtwoord in.',
    loginFailed: 'Onbekende gebruikersnaam of fout wachtwoord.', loginGenericFail: 'Aanmelden mislukt.', tooManyAttempts: 'Te veel mislukte pogingen. Probeer het binnen 15 minuten opnieuw.',
    noConnection: 'Kon niet verbinden met de server.', noConnectionRelogin: 'Geen verbinding met de server. Meld opnieuw aan.',
    sessionExpired: 'Je sessie is verlopen. Meld opnieuw aan.', noAccess: 'Geen toegang.',
    loggedInAs: 'Ingelogd als', logout: 'Uitloggen',
    role_admin: 'admin', role_user: 'gebruiker', role_sales: 'verkoper', role_dir: 'directeur',
    helpBtn: '? Uitleg data', reloadBtn: 'Data vernieuwen',
    dataNone: 'Er staat nog geen klachtendata online.', dataNoneAdmin: 'Er staat nog geen klachtendata online. Laad KLACHTEN_DATA.xlsx op in de admin-sectie onderaan.',
    dataLoading: 'Data ophalen…', dataLoadingN: 'Data ophalen… ({done}/{total})', dataError: 'Kon de data niet ophalen: {msg}',
    dataInfo: '<b>{n}</b> klachten · bijgewerkt op <b>{date}</b> · laatste klacht {last}',
    dataScope: 'Je ziet de klachten van vertegenwoordiger {vts}.',
    adminDataTitle: 'Klachtendata opladen', adminDataSub: 'Laad <b>KLACHTEN_DATA.xlsx</b> op (zie DATA_OPHALEN.md). De tool vergelijkt met de data die nu online staat en laadt enkel de gewijzigde delen op.',
    dropData: '<b>Sleep KLACHTEN_DATA.xlsx hierheen</b> of klik om het bestand te kiezen.',
    reading: 'Bestand inlezen… (dit kan enkele seconden duren)', readError: 'Kon het bestand niet lezen: {msg}',
    diffSummary: '<b>{n}</b> klachten in het bestand: <b>{new}</b> nieuw, <b>{changed}</b> gewijzigd, <b>{removed}</b> niet meer aanwezig, {same} ongewijzigd. {chunks} van {total} datablokken op te laden.',
    diffNothing: 'Geen wijzigingen ten opzichte van de data die online staat.',
    uploadBtn: 'Opladen', uploading: 'Opladen… ({done}/{total})', uploadDone: 'Opgeladen: {n} klachten online ({new} nieuw, {changed} gewijzigd, {removed} verwijderd).',
    uploadError: 'Opladen mislukt: {msg}',
    lastUpload: 'Laatste upload: {date} door {by} ({file}) · {new} nieuw, {changed} gewijzigd, {removed} verwijderd.',
    usersTitle: 'Gebruikers beheren', usersSub: 'Excel met kolommen USER, PW, ROLE (admin, user, sales, dir of delete), VTNR (VT-nummer(s), gescheiden door ; — verplicht voor sales en dir) en Regio. Een dir ziet zijn eigen VT-nummer(s) plus die van alle sales- en dir-gebruikers met dezelfde Regio. Gebruikers die niet in het bestand staan, blijven ongewijzigd.',
    dropUsers: '<b>Sleep de gebruikers-Excel hierheen</b> of klik om te kiezen.', usersUpdateBtn: 'Gebruikers bijwerken',
    usersDone: '{updated} bijgewerkt, {deleted} verwijderd.', usersNoRows: 'Geen gebruikers gevonden. Verwacht kolommen USER, PW, ROLE (en VTNR, Regio).',
    usersListBtn: 'Huidige gebruikers tonen',
    unknownError: 'Onbekende fout.',
  },
  fr: {
    appTitle: 'Analyse des plaintes',
    uiLanguage: "Langue de l'application", auto: 'Automatique',
    loginSub: 'Connectez-vous pour continuer.', username: "Nom d'utilisateur", password: 'Mot de passe', pwToggle: 'Afficher/masquer le mot de passe',
    remember: 'Rester connecté', loginBtn: 'Se connecter', loggingIn: 'Connexion…', loginMissing: "Saisissez le nom d'utilisateur et le mot de passe.",
    loginFailed: "Nom d'utilisateur inconnu ou mot de passe incorrect.", loginGenericFail: 'Échec de la connexion.', tooManyAttempts: 'Trop de tentatives échouées. Réessayez dans 15 minutes.',
    noConnection: 'Impossible de joindre le serveur.', noConnectionRelogin: 'Pas de connexion au serveur. Reconnectez-vous.',
    sessionExpired: 'Votre session a expiré. Reconnectez-vous.', noAccess: 'Accès refusé.',
    loggedInAs: 'Connecté en tant que', logout: 'Se déconnecter',
    role_admin: 'admin', role_user: 'utilisateur', role_sales: 'commercial', role_dir: 'directeur',
    helpBtn: '? Explication des données', reloadBtn: 'Actualiser les données',
    dataNone: "Aucune donnée de plaintes n'est encore en ligne.", dataNoneAdmin: "Aucune donnée de plaintes n'est encore en ligne. Chargez KLACHTEN_DATA.xlsx dans la section admin en bas.",
    dataLoading: 'Chargement des données…', dataLoadingN: 'Chargement des données… ({done}/{total})', dataError: 'Impossible de charger les données : {msg}',
    dataInfo: '<b>{n}</b> plaintes · mises à jour le <b>{date}</b> · dernière plainte {last}',
    dataScope: 'Vous voyez les plaintes du représentant {vts}.',
    unknownError: 'Erreur inconnue.',
  },
  en: {
    appTitle: 'Complaints analysis',
    uiLanguage: 'App language', auto: 'Automatic',
    loginSub: 'Log in to continue.', username: 'Username', password: 'Password', pwToggle: 'Show/hide password',
    remember: 'Stay logged in', loginBtn: 'Log in', loggingIn: 'Logging in…', loginMissing: 'Enter username and password.',
    loginFailed: 'Unknown username or wrong password.', loginGenericFail: 'Login failed.', tooManyAttempts: 'Too many failed attempts. Try again in 15 minutes.',
    noConnection: 'Could not connect to the server.', noConnectionRelogin: 'No connection to the server. Please log in again.',
    sessionExpired: 'Your session has expired. Please log in again.', noAccess: 'No access.',
    loggedInAs: 'Logged in as', logout: 'Log out',
    role_admin: 'admin', role_user: 'user', role_sales: 'sales', role_dir: 'director',
    helpBtn: '? About the data', reloadBtn: 'Reload data',
    dataNone: 'No complaints data is online yet.', dataNoneAdmin: 'No complaints data is online yet. Upload KLACHTEN_DATA.xlsx in the admin section below.',
    dataLoading: 'Loading data…', dataLoadingN: 'Loading data… ({done}/{total})', dataError: 'Could not load the data: {msg}',
    dataInfo: '<b>{n}</b> complaints · updated on <b>{date}</b> · latest complaint {last}',
    dataScope: 'You see the complaints of sales representative {vts}.',
    unknownError: 'Unknown error.',
  },
  de: {
    appTitle: 'Reklamationsanalyse',
    uiLanguage: 'Sprache der App', auto: 'Automatisch',
    loginSub: 'Melden Sie sich an, um fortzufahren.', username: 'Benutzername', password: 'Passwort', pwToggle: 'Passwort anzeigen/verbergen',
    remember: 'Angemeldet bleiben', loginBtn: 'Anmelden', loggingIn: 'Anmelden…', loginMissing: 'Benutzername und Passwort eingeben.',
    loginFailed: 'Unbekannter Benutzername oder falsches Passwort.', loginGenericFail: 'Anmeldung fehlgeschlagen.', tooManyAttempts: 'Zu viele Fehlversuche. Versuchen Sie es in 15 Minuten erneut.',
    noConnection: 'Keine Verbindung zum Server.', noConnectionRelogin: 'Keine Verbindung zum Server. Bitte erneut anmelden.',
    sessionExpired: 'Ihre Sitzung ist abgelaufen. Bitte erneut anmelden.', noAccess: 'Kein Zugriff.',
    loggedInAs: 'Angemeldet als', logout: 'Abmelden',
    role_admin: 'Admin', role_user: 'Benutzer', role_sales: 'Vertrieb', role_dir: 'Direktor',
    helpBtn: '? Erklärung der Daten', reloadBtn: 'Daten neu laden',
    dataNone: 'Es sind noch keine Reklamationsdaten online.', dataNoneAdmin: 'Es sind noch keine Reklamationsdaten online. Laden Sie KLACHTEN_DATA.xlsx im Admin-Bereich unten hoch.',
    dataLoading: 'Daten werden geladen…', dataLoadingN: 'Daten werden geladen… ({done}/{total})', dataError: 'Daten konnten nicht geladen werden: {msg}',
    dataInfo: '<b>{n}</b> Reklamationen · aktualisiert am <b>{date}</b> · letzte Reklamation {last}',
    dataScope: 'Sie sehen die Reklamationen von Vertreter {vts}.',
    unknownError: 'Unbekannter Fehler.',
  },
  es: {
    appTitle: 'Análisis de reclamaciones',
    uiLanguage: 'Idioma de la aplicación', auto: 'Automático',
    loginSub: 'Inicie sesión para continuar.', username: 'Usuario', password: 'Contraseña', pwToggle: 'Mostrar/ocultar contraseña',
    remember: 'Mantener la sesión', loginBtn: 'Iniciar sesión', loggingIn: 'Iniciando sesión…', loginMissing: 'Introduzca usuario y contraseña.',
    loginFailed: 'Usuario desconocido o contraseña incorrecta.', loginGenericFail: 'Error al iniciar sesión.', tooManyAttempts: 'Demasiados intentos fallidos. Inténtelo de nuevo en 15 minutos.',
    noConnection: 'No se pudo conectar con el servidor.', noConnectionRelogin: 'Sin conexión con el servidor. Vuelva a iniciar sesión.',
    sessionExpired: 'Su sesión ha caducado. Vuelva a iniciar sesión.', noAccess: 'Sin acceso.',
    loggedInAs: 'Sesión iniciada como', logout: 'Cerrar sesión',
    role_admin: 'admin', role_user: 'usuario', role_sales: 'comercial', role_dir: 'director',
    helpBtn: '? Explicación de los datos', reloadBtn: 'Recargar datos',
    dataNone: 'Todavía no hay datos de reclamaciones en línea.', dataNoneAdmin: 'Todavía no hay datos de reclamaciones en línea. Cargue KLACHTEN_DATA.xlsx en la sección de administración abajo.',
    dataLoading: 'Cargando datos…', dataLoadingN: 'Cargando datos… ({done}/{total})', dataError: 'No se pudieron cargar los datos: {msg}',
    dataInfo: '<b>{n}</b> reclamaciones · actualizado el <b>{date}</b> · última reclamación {last}',
    dataScope: 'Ve las reclamaciones del representante {vts}.',
    unknownError: 'Error desconocido.',
  },
  it: {
    appTitle: 'Analisi dei reclami',
    uiLanguage: "Lingua dell'app", auto: 'Automatico',
    loginSub: 'Accedi per continuare.', username: 'Nome utente', password: 'Password', pwToggle: 'Mostra/nascondi password',
    remember: 'Resta connesso', loginBtn: 'Accedi', loggingIn: 'Accesso…', loginMissing: 'Inserisci nome utente e password.',
    loginFailed: 'Nome utente sconosciuto o password errata.', loginGenericFail: 'Accesso non riuscito.', tooManyAttempts: 'Troppi tentativi falliti. Riprova tra 15 minuti.',
    noConnection: 'Impossibile connettersi al server.', noConnectionRelogin: 'Nessuna connessione al server. Accedi di nuovo.',
    sessionExpired: 'La sessione è scaduta. Accedi di nuovo.', noAccess: 'Accesso negato.',
    loggedInAs: 'Connesso come', logout: 'Esci',
    role_admin: 'admin', role_user: 'utente', role_sales: 'commerciale', role_dir: 'direttore',
    helpBtn: '? Spiegazione dei dati', reloadBtn: 'Ricarica dati',
    dataNone: 'Non ci sono ancora dati sui reclami online.', dataNoneAdmin: 'Non ci sono ancora dati sui reclami online. Carica KLACHTEN_DATA.xlsx nella sezione admin in basso.',
    dataLoading: 'Caricamento dati…', dataLoadingN: 'Caricamento dati… ({done}/{total})', dataError: 'Impossibile caricare i dati: {msg}',
    dataInfo: '<b>{n}</b> reclami · aggiornati il <b>{date}</b> · ultimo reclamo {last}',
    dataScope: 'Vedi i reclami del rappresentante {vts}.',
    unknownError: 'Errore sconosciuto.',
  },
};

// ---------------------------------------------------------- logica
function detectBrowserLang() {
  const list = navigator.languages && navigator.languages.length ? navigator.languages : [navigator.language || 'nl'];
  for (const l of list) { const code = String(l || '').slice(0, 2).toLowerCase(); if (I18N[code]) return code; }
  return 'nl';
}
function uiLangSetting() {
  try { const v = localStorage.getItem(UILANG_KEY); return v && (v === 'auto' || I18N[v]) ? v : 'auto'; } catch { return 'auto'; }
}
function saveUiLangSetting(v) { try { localStorage.setItem(UILANG_KEY, v); } catch { /* storage geblokkeerd */ } }
let uiLangOverride = null;
function currentUiLang() { const v = uiLangOverride || uiLangSetting(); return v === 'auto' ? detectBrowserLang() : v; }
function t(key, vars) {
  const dict = I18N[currentUiLang()] || I18N.nl;
  let s = dict[key] ?? I18N.nl[key] ?? key;
  if (vars) s = s.replace(/\{(\w+)\}/g, (m, k) => (k in vars ? String(vars[k]) : m));
  return s;
}
function applyI18n(root) {
  const r = root || document;
  r.querySelectorAll('[data-i18n]').forEach((el) => { el.textContent = t(el.dataset.i18n); });
  r.querySelectorAll('[data-i18n-html]').forEach((el) => { el.innerHTML = t(el.dataset.i18nHtml); });
  r.querySelectorAll('[data-i18n-placeholder]').forEach((el) => { el.placeholder = t(el.dataset.i18nPlaceholder); });
  r.querySelectorAll('[data-i18n-title]').forEach((el) => { el.title = t(el.dataset.i18nTitle); });
  document.documentElement.lang = currentUiLang();
  document.title = 'Winsol ' + t('appTitle');
}
function fillUiLangSelect(sel) {
  sel.innerHTML = '';
  sel.add(new Option(t('auto'), 'auto'));
  for (const [code, name] of Object.entries(UI_LANGS)) sel.add(new Option(name, code));
  sel.value = uiLangOverride || uiLangSetting();
}
