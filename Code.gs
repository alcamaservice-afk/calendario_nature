// ============================================================
//  CALENDARIO PULIZIE – Google Apps Script Backend v3.0
//  Sistema con account collaboratori, autenticazione,
//  contabilità storica e gestione admin avanzata.
// ============================================================

// ══ CONFIGURAZIONE ══════════════════════════════════════════
var SPREADSHEET_NAME = "Calendario Pulizie - App";

var MESES_ES = [
  "ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO",
  "JULIO","AGOSTO","SEPTIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE"
];

// Columnas del spreadsheet calendario (índices 0-based)
var COL = {
  ID: 0, APARTAMENTO: 1, FECHA: 2, HORARIO: 3,
  PREPARAR: 4, ENTRADA: 5, NOTAS: 6, DIAS_LIBRES: 7,
  PLAT: 8, ENCARGADO: 9, NOTA_ENCARGADO: 10, TIMESTAMP: 11
};
var NUM_COLS = 12;

// Columnas del foglio UTENTI
var COL_U = {
  EMAIL: 0, PASSWORD_HASH: 1, NOMBRE: 2, ENCARGADO_KEY: 3,
  ROLE: 4, ACTIVO: 5, BLOQUEADO: 6, CREADO: 7, ULTIMO_ACCESO: 8
};
var NUM_COLS_U = 9;

// Columnas del foglio CONTABILIDAD
var COL_C = {
  ID: 0, ENCARGADO: 1, EMAIL: 2, MES: 3, ANNO: 4,
  N_LIMPIEZAS: 5, PRECIO_LIMP: 6, GASTOS: 7, TOTAL: 8,
  ESTADO: 9, FECHA_PAGO: 10, NOTAS: 11, CREADO: 12
};
var NUM_COLS_C = 13;

var DEFAULT_ENCARGADOS = [
  { nombre: "Daniela", tel: "34643270991" },
  { nombre: "Lina", tel: "34600873148" },
  { nombre: "Levys", tel: "34643270991" },
  { nombre: "Alejandra", tel: "" },
  { nombre: "Marcello", tel: "" }
];

var DEFAULT_APARTAMENTOS = [
  "Juan Llorens",
  "Palleter",
  "San Ignacio"
];

// ══ URL WEB APP ══════════════════════════════════════════════
function getDeployedUrl() {
  try {
    var props = PropertiesService.getScriptProperties();
    var custom = props.getProperty('WEB_APP_URL');
    if (custom && custom.indexOf('http') === 0 && custom.indexOf('TU_ID') === -1) {
      return custom.split('?')[0].split('#')[0];
    }
    var url = ScriptApp.getService().getUrl();
    if (url && url.indexOf('TU_ID') === -1) {
      return url.split('?')[0].split('#')[0];
    }
  } catch (e) {}
  return "https://script.google.com/macros/s/TU_ID/exec";
}

function guardarWebAppUrl(url) {
  try {
    if (url && url.indexOf('http') === 0 && url.indexOf('TU_ID') === -1) {
      var clean = url.split('?')[0].split('#')[0];
      PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', clean);
      return { success: true, url: clean };
    }
    return { success: false, error: 'URL non valida' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ══ ROUTING ═════════════════════════════════════════════════
function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var req = JSON.parse(raw);
    var action = req.action || "";
    var res = executeAction(action, req);
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // Supporto API GET
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
  if (action) {
    var res = executeAction(action, e.parameter);
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Visualizzazione interfaccia Web App
  try {
    return HtmlService.createHtmlOutputFromFile("Index")
      .setTitle("Calendario Pulizie")
      .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(e1) {
    try {
      return HtmlService.createHtmlOutputFromFile("index")
        .setTitle("Calendario Pulizie")
        .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch(e2) {
      return HtmlService.createHtmlOutput(getStandaloneAppHtml())
        .setTitle("Calendario Pulizie")
        .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch(e) {
    return "";
  }
}

function executeAction(action, req) {
  switch (action) {
    case "ping":
      return { success: true, message: "pong", timestamp: new Date().toISOString() };

    case "login":
      return login(req.email, req.password);

    case "verificarSesion":
      return verificarSesion(req.token);

    case "logout":
      return logout(req.token);

    case "cambiarPassword":
      return cambiarPassword(req.token, req.oldPassword, req.newPassword);

    case "getUsuarios":
      return getUsuarios(req.token);

    case "crearUsuario":
      return crearUsuario(req.token, req.userData || req);

    case "actualizarUsuario":
      return actualizarUsuario(req.token, req.id || req.email, req.userData || req);

    case "resetPasswordUsuario":
      return resetPasswordUsuario(req.token, req.id || req.email, req.newPassword);

    case "toggleBloquearUsuario":
      return toggleBloquearUsuario(req.token, req.id || req.email);

    case "eliminarUsuario":
      return eliminarUsuario(req.token, req.id || req.email);

    case "getCalendario":
      return getCalendario(req.token, req.mes, req.anno, req.userFilter, req.aptFilter);

    case "saveRiga":
      return saveRiga(req.token, req.rowData || req.data, req.mes, req.anno);

    case "deleteRiga":
      return deleteRiga(req.token, req.id, req.mes, req.anno);

    case "saveNotaEncargado":
      return saveNotaEncargado(req.token, req.id, req.mes, req.anno, req.nota);

    case "getContabilidad":
      return getContabilidad(req.token, req.filtros || req);

    case "saveContabilidad":
      return saveContabilidad(req.token, req.contData || req);

    case "updateEstadoContabilidad":
      return updateEstadoContabilidad(req.token, req.contId || req.id, req.estado || req.stato, req.fechaPago, req.notas);

    case "calcolarCompensacion":
      return calcolarCompensacion(req.token, req.encargado, req.email, req.mes, req.anno, req.precioPorLimpieza);

    case "getAppConfig":
      return getAppConfig(req.token);

    case "saveAppConfig":
      return saveAppConfig(req.token, req.encargados, req.apartamentos);

    case "getDatosPDF":
      return getDatosPDF(req.token, req.mes, req.anno);

    case "initializeSystem":
      return initializeSystem();

    default:
      return { success: false, error: "Azione non riconosciuta: " + action };
  }
}

function getStandaloneAppHtml() {
  return "<!DOCTYPE html>\n<html lang=\"it\">\n<head>\n  <meta charset=\"UTF-8\">\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no\">\n  <title>Calendario Pulizie</title>\n  \n  <!-- PWA & Mobile Meta Tags -->\n  <meta name=\"theme-color\" content=\"#0f172a\">\n  <meta name=\"apple-mobile-web-app-capable\" content=\"yes\">\n  <meta name=\"apple-mobile-web-app-status-bar-style\" content=\"black-translucent\">\n  <meta name=\"apple-mobile-web-app-title\" content=\"Pulizie\">\n  <link rel=\"manifest\" href=\"manifest.json\">\n  <link rel=\"apple-touch-icon\" href=\"icons/icon-192.png\">\n  <link rel=\"icon\" type=\"image/svg+xml\" href=\"icons/icon-192.svg\">\n  <link rel=\"alternate icon\" type=\"image/png\" href=\"icons/icon-192.png\">\n\n  <!-- Google Fonts: Inter -->\n  <link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n  <link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>\n  <link href=\"https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap\" rel=\"stylesheet\">\n\n  <style>\n    :root {\n      --primary: #14b8a6;\n      --primary-hover: #0d9488;\n      --primary-glow: rgba(20, 184, 166, 0.25);\n      --bg: #0f172a;\n      --surface: #1e293b;\n      --surface-hover: #273549;\n      --surface-border: #334155;\n      --text: #f8fafc;\n      --text-muted: #94a3b8;\n      --danger: #ef4444;\n      --danger-hover: #dc2626;\n      --success: #22c55e;\n      --warning: #f59e0b;\n      --blue: #3b82f6;\n      --purple: #8b5cf6;\n      --radius: 12px;\n      --shadow: 0 4px 20px -2px rgba(0, 0, 0, 0.4);\n    }\n\n    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Inter', -apple-system, sans-serif; }\n    body { background-color: var(--bg); color: var(--text); min-height: 100vh; display: flex; flex-direction: column; overflow-x: hidden; }\n\n    /* Top Bar */\n    header {\n      background: rgba(30, 41, 59, 0.85);\n      backdrop-filter: blur(12px);\n      border-bottom: 1px solid var(--surface-border);\n      position: sticky;\n      top: 0;\n      z-index: 100;\n      padding: 12px 20px;\n      display: flex;\n      align-items: center;\n      justify-content: space-between;\n    }\n    .brand {\n      display: flex;\n      align-items: center;\n      gap: 12px;\n      text-decoration: none;\n      color: inherit;\n    }\n    .brand-logo {\n      width: 38px;\n      height: 38px;\n      border-radius: 10px;\n      background: linear-gradient(135deg, #14b8a6, #0284c7);\n      display: flex;\n      align-items: center;\n      justify-content: center;\n      box-shadow: 0 2px 10px var(--primary-glow);\n    }\n    .brand-logo svg { width: 22px; height: 22px; fill: #ffffff; }\n    .brand-title { font-weight: 700; font-size: 1.15rem; letter-spacing: -0.02em; }\n    .header-actions { display: flex; align-items: center; gap: 12px; }\n    .user-pill {\n      background: var(--surface);\n      border: 1px solid var(--surface-border);\n      padding: 6px 14px;\n      border-radius: 20px;\n      font-size: 0.82rem;\n      display: flex;\n      align-items: center;\n      gap: 8px;\n    }\n    .status-dot {\n      width: 8px;\n      height: 8px;\n      border-radius: 50%;\n      background-color: var(--success);\n      display: inline-block;\n      box-shadow: 0 0 8px var(--success);\n    }\n    .status-dot.syncing { background-color: var(--warning); animation: pulse 1s infinite; }\n\n    /* PWA Banner */\n    #pwa-install-banner {\n      background: linear-gradient(90deg, #14b8a6, #0284c7);\n      color: white;\n      padding: 10px 16px;\n      display: none;\n      align-items: center;\n      justify-content: space-between;\n      font-size: 0.88rem;\n      font-weight: 500;\n      box-shadow: 0 2px 12px rgba(2, 132, 199, 0.3);\n      position: relative;\n      z-index: 105;\n    }\n    #pwa-install-banner button {\n      background: white;\n      color: #0f172a;\n      border: none;\n      padding: 6px 14px;\n      border-radius: 20px;\n      font-weight: 600;\n      font-size: 0.8rem;\n      cursor: pointer;\n    }\n\n    /* Container */\n    main { flex: 1; padding: 20px; max-width: 1280px; margin: 0 auto; width: 100%; }\n\n    /* Nav Tabs */\n    .tabs-nav {\n      display: flex;\n      gap: 8px;\n      background: var(--surface);\n      padding: 6px;\n      border-radius: var(--radius);\n      border: 1px solid var(--surface-border);\n      margin-bottom: 20px;\n      overflow-x: auto;\n      scrollbar-width: none;\n    }\n    .tabs-nav::-webkit-scrollbar { display: none; }\n    .tab-btn {\n      background: transparent;\n      border: none;\n      color: var(--text-muted);\n      padding: 10px 18px;\n      border-radius: 8px;\n      font-weight: 600;\n      font-size: 0.88rem;\n      cursor: pointer;\n      display: flex;\n      align-items: center;\n      gap: 8px;\n      white-space: nowrap;\n      transition: all 0.2s ease;\n    }\n    .tab-btn.active {\n      background: var(--primary);\n      color: #0f172a;\n      box-shadow: 0 2px 10px var(--primary-glow);\n    }\n\n    /* Views */\n    .view-section { display: none; animation: fadeIn 0.25s ease-out forwards; }\n    .view-section.active { display: block; }\n\n    /* Cards */\n    .card {\n      background: var(--surface);\n      border: 1px solid var(--surface-border);\n      border-radius: var(--radius);\n      padding: 20px;\n      margin-bottom: 20px;\n      box-shadow: var(--shadow);\n    }\n    .card-title {\n      font-size: 1.1rem;\n      font-weight: 600;\n      margin-bottom: 16px;\n      display: flex;\n      align-items: center;\n      justify-content: space-between;\n    }\n\n    /* Summary Grid */\n    .stats-grid {\n      display: grid;\n      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));\n      gap: 16px;\n      margin-bottom: 24px;\n    }\n    .stat-card {\n      background: var(--surface);\n      border: 1px solid var(--surface-border);\n      border-radius: var(--radius);\n      padding: 16px;\n      display: flex;\n      flex-direction: column;\n      gap: 6px;\n    }\n    .stat-label { font-size: 0.78rem; font-weight: 500; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; }\n    .stat-value { font-size: 1.6rem; font-weight: 700; color: var(--text); }\n    .stat-value.primary { color: var(--primary); }\n    .stat-value.success { color: var(--success); }\n    .stat-value.warning { color: var(--warning); }\n    .stat-value.blue { color: var(--blue); }\n\n    /* Filter Bar */\n    .filter-bar {\n      display: flex;\n      flex-wrap: wrap;\n      align-items: center;\n      gap: 12px;\n      background: var(--surface);\n      padding: 14px 18px;\n      border-radius: var(--radius);\n      border: 1px solid var(--surface-border);\n      margin-bottom: 20px;\n    }\n    .filter-group { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }\n    .filter-label { font-size: 0.82rem; font-weight: 600; color: var(--text-muted); }\n\n    /* Forms & Inputs */\n    select, input[type=\"text\"], input[type=\"email\"], input[type=\"password\"], input[type=\"date\"], input[type=\"number\"], textarea {\n      background: #0f172a;\n      border: 1px solid var(--surface-border);\n      color: var(--text);\n      padding: 10px 14px;\n      border-radius: 8px;\n      font-size: 0.9rem;\n      outline: none;\n      transition: border-color 0.2s;\n    }\n    select:focus, input:focus, textarea:focus { border-color: var(--primary); box-shadow: 0 0 0 2px var(--primary-glow); }\n    .form-group { margin-bottom: 16px; display: flex; flex-direction: column; gap: 6px; }\n    .form-group label { font-size: 0.84rem; font-weight: 500; color: var(--text-muted); }\n\n    /* Buttons */\n    .btn {\n      background: var(--surface-hover);\n      border: 1px solid var(--surface-border);\n      color: var(--text);\n      padding: 9px 18px;\n      border-radius: 8px;\n      font-weight: 600;\n      font-size: 0.88rem;\n      cursor: pointer;\n      display: inline-flex;\n      align-items: center;\n      justify-content: center;\n      gap: 8px;\n      transition: all 0.2s;\n    }\n    .btn:hover { background: #334155; }\n    .btn-primary { background: var(--primary); color: #0f172a; border-color: transparent; }\n    .btn-primary:hover { background: var(--primary-hover); }\n    .btn-danger { background: rgba(239, 68, 68, 0.15); color: #f87171; border-color: rgba(239, 68, 68, 0.3); }\n    .btn-danger:hover { background: var(--danger); color: white; }\n    .btn-success { background: rgba(34, 197, 94, 0.15); color: #4ade80; border-color: rgba(34, 197, 94, 0.3); }\n    .btn-success:hover { background: var(--success); color: white; }\n    .btn-sm { padding: 6px 12px; font-size: 0.78rem; border-radius: 6px; }\n\n    /* Tables */\n    .table-container { overflow-x: auto; border-radius: var(--radius); border: 1px solid var(--surface-border); }\n    table { width: 100%; border-collapse: collapse; text-align: left; font-size: 0.86rem; }\n    th { background: #162032; padding: 12px 14px; font-weight: 600; color: var(--text-muted); border-bottom: 1px solid var(--surface-border); white-space: nowrap; }\n    td { padding: 12px 14px; border-bottom: 1px solid rgba(51, 65, 85, 0.5); vertical-align: middle; }\n    tr:last-child td { border-bottom: none; }\n    tr:hover td { background: rgba(39, 53, 73, 0.5); }\n\n    /* Badges */\n    .badge {\n      display: inline-flex;\n      align-items: center;\n      padding: 4px 10px;\n      border-radius: 20px;\n      font-size: 0.74rem;\n      font-weight: 600;\n      white-space: nowrap;\n    }\n    .badge-apt { background: rgba(59, 130, 246, 0.15); color: #60a5fa; border: 1px solid rgba(59, 130, 246, 0.3); }\n    .badge-enc { background: rgba(139, 92, 246, 0.15); color: #a78bfa; border: 1px solid rgba(139, 92, 246, 0.3); }\n    .badge-success { background: rgba(34, 197, 94, 0.15); color: #4ade80; border: 1px solid rgba(34, 197, 94, 0.3); }\n    .badge-warning { background: rgba(245, 158, 11, 0.15); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.3); }\n    .badge-danger { background: rgba(239, 68, 68, 0.15); color: #f87171; border: 1px solid rgba(239, 68, 68, 0.3); }\n\n    /* Login Box */\n    .login-wrapper {\n      min-height: 80vh;\n      display: flex;\n      align-items: center;\n      justify-content: center;\n      padding: 20px;\n    }\n    .login-card {\n      background: var(--surface);\n      border: 1px solid var(--surface-border);\n      border-radius: 16px;\n      padding: 32px 28px;\n      max-width: 400px;\n      width: 100%;\n      box-shadow: 0 10px 40px rgba(0, 0, 0, 0.5);\n      text-align: center;\n    }\n    .login-logo {\n      width: 64px;\n      height: 64px;\n      border-radius: 16px;\n      background: linear-gradient(135deg, #14b8a6, #0284c7);\n      display: inline-flex;\n      align-items: center;\n      justify-content: center;\n      margin-bottom: 20px;\n      box-shadow: 0 4px 20px var(--primary-glow);\n    }\n    .login-logo svg { width: 36px; height: 36px; fill: white; }\n\n    /* Modals */\n    .modal-overlay {\n      position: fixed;\n      inset: 0;\n      background: rgba(0, 0, 0, 0.7);\n      backdrop-filter: blur(4px);\n      display: none;\n      align-items: center;\n      justify-content: center;\n      z-index: 200;\n      padding: 20px;\n    }\n    .modal-overlay.active { display: flex; animation: fadeIn 0.2s ease-out; }\n    .modal {\n      background: var(--surface);\n      border: 1px solid var(--surface-border);\n      border-radius: 16px;\n      padding: 24px;\n      max-width: 500px;\n      width: 100%;\n      max-height: 90vh;\n      overflow-y: auto;\n      box-shadow: 0 10px 40px rgba(0, 0, 0, 0.6);\n    }\n    .modal-header {\n      display: flex;\n      align-items: center;\n      justify-content: space-between;\n      margin-bottom: 20px;\n    }\n    .modal-title { font-size: 1.15rem; font-weight: 700; }\n    .modal-close {\n      background: transparent;\n      border: none;\n      color: var(--text-muted);\n      font-size: 1.4rem;\n      cursor: pointer;\n    }\n\n    /* Toast */\n    #toast-container {\n      position: fixed;\n      bottom: 24px;\n      right: 24px;\n      z-index: 300;\n      display: flex;\n      flex-direction: column;\n      gap: 10px;\n    }\n    .toast {\n      background: #1e293b;\n      border: 1px solid var(--surface-border);\n      color: white;\n      padding: 12px 18px;\n      border-radius: 10px;\n      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.4);\n      display: flex;\n      align-items: center;\n      gap: 10px;\n      font-size: 0.88rem;\n      font-weight: 500;\n      animation: slideIn 0.25s ease-out;\n    }\n    .toast.success { border-left: 4px solid var(--success); }\n    .toast.error { border-left: 4px solid var(--danger); }\n    .toast.info { border-left: 4px solid var(--primary); }\n\n    /* Animations */\n    @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }\n    @keyframes slideIn { from { transform: translateX(100%); opacity: 0; } to { transform: translateX(0); opacity: 1; } }\n    @keyframes pulse { 0% { opacity: 1; } 50% { opacity: 0.4; } 100% { opacity: 1; } }\n\n    /* Responsive */\n    @media (max-width: 768px) {\n      header { padding: 10px 14px; }\n      main { padding: 12px; }\n      .brand-title { font-size: 1rem; }\n      .tab-btn { padding: 8px 12px; font-size: 0.8rem; }\n    }\n  </style>\n</head>\n<body>\n\n  <!-- PWA Install Banner -->\n  <div id=\"pwa-install-banner\">\n    <div style=\"display:flex; align-items:center; gap:10px;\">\n      <span>🧹</span>\n      <span id=\"pwa-prompt-text\">Installa l'App Calendario Pulizie sulla tua schermata home</span>\n    </div>\n    <div style=\"display:flex; align-items:center; gap:8px;\">\n      <button id=\"btn-install-pwa\">Installa</button>\n      <button style=\"background:transparent; color:white; padding:4px;\" onclick=\"dismissPwaBanner()\">✕</button>\n    </div>\n  </div>\n\n  <!-- Header -->\n  <header>\n    <a href=\"#\" class=\"brand\">\n      <div class=\"brand-logo\">\n        <svg viewBox=\"0 0 24 24\"><path d=\"M19.36 2.72l1.42 1.42-3.88 3.88-1.42-1.42 3.88-3.88m-5.88 5.88l1.42 1.42-7.07 7.07c-.39.39-1.02.39-1.41 0l-1.42-1.42c-.39-.39-.39-1.02 0-1.41l7.07-7.07m-5.65 9.9l1.41 1.41-1.41 1.42c-.78.78-2.05.78-2.83 0s-.78-2.05 0-2.83l1.41-1.41z\"/></svg>\n      </div>\n      <div>\n        <div class=\"brand-title\">Calendario Pulizie</div>\n      </div>\n    </a>\n    <div class=\"header-actions\">\n      <div class=\"user-pill\" id=\"user-status-pill\" style=\"display:none;\">\n        <span class=\"status-dot\" id=\"sync-dot\"></span>\n        <span id=\"header-user-name\">Utente</span>\n        <span class=\"badge badge-apt\" id=\"header-user-role\" style=\"font-size:0.7rem; padding:2px 6px;\">Admin</span>\n      </div>\n      <button class=\"btn btn-sm btn-danger\" id=\"btn-logout\" style=\"display:none;\" onclick=\"appLogout()\">Esci</button>\n    </div>\n  </header>\n\n  <!-- Main Container -->\n  <main>\n\n    <!-- ==================== VIEW: LOGIN ==================== -->\n    <div id=\"view-login\" class=\"view-section active\">\n      <div class=\"login-wrapper\">\n        <div class=\"login-card\">\n          <div class=\"login-logo\">\n            <svg viewBox=\"0 0 24 24\"><path d=\"M19.36 2.72l1.42 1.42-3.88 3.88-1.42-1.42 3.88-3.88m-5.88 5.88l1.42 1.42-7.07 7.07c-.39.39-1.02.39-1.41 0l-1.42-1.42c-.39-.39-.39-1.02 0-1.41l7.07-7.07m-5.65 9.9l1.41 1.41-1.41 1.42c-.78.78-2.05.78-2.83 0s-.78-2.05 0-2.83l1.41-1.41z\"/></svg>\n          </div>\n          <h2 style=\"font-size: 1.4rem; font-weight:700; margin-bottom:6px;\">Bentornato</h2>\n          <p style=\"color:var(--text-muted); font-size:0.86rem; margin-bottom:24px;\">Accedi al tuo account per gestire pulizie e contabilità</p>\n\n          <form id=\"form-login\" onsubmit=\"handleLoginSubmit(event)\">\n            <div class=\"form-group\" style=\"text-align:left;\">\n              <label>Email</label>\n              <input type=\"email\" id=\"login-email\" required placeholder=\"tua@email.com\">\n            </div>\n            <div class=\"form-group\" style=\"text-align:left; margin-bottom:24px;\">\n              <label>Password</label>\n              <input type=\"password\" id=\"login-password\" required placeholder=\"••••••••\">\n            </div>\n            <button type=\"submit\" class=\"btn btn-primary\" id=\"btn-login-submit\" style=\"width:100%; padding:12px; font-size:0.95rem;\">\n              Accedi\n            </button>\n          </form>\n\n          <div style=\"margin-top:20px; font-size:0.78rem; color:var(--text-muted);\">\n            <a href=\"#\" onclick=\"openModal('modal-script-url')\" style=\"color:var(--text-muted); text-decoration:none;\">⚙️ Impostazioni Connessione</a>\n          </div>\n        </div>\n      </div>\n    </div>\n\n    <!-- ==================== VIEW: ADMIN ==================== -->\n    <div id=\"view-admin\" class=\"view-section\">\n      <div class=\"tabs-nav\">\n        <button class=\"tab-btn active\" onclick=\"switchAdminTab('calendario')\">📅 Calendario</button>\n        <button class=\"tab-btn\" onclick=\"switchAdminTab('collaboratori')\">👥 Collaboratori</button>\n        <button class=\"tab-btn\" onclick=\"switchAdminTab('contabilita')\">💰 Contabilità</button>\n        <button class=\"tab-btn\" onclick=\"switchAdminTab('configurazione')\">⚙️ Configurazione</button>\n        <button class=\"tab-btn\" onclick=\"switchAdminTab('profilo')\">👤 Profilo</button>\n      </div>\n\n      <!-- Tab Admin: Calendario -->\n      <div id=\"admin-tab-calendario\" class=\"admin-subview\">\n        <div class=\"filter-bar\">\n          <div class=\"filter-group\">\n            <span class=\"filter-label\">Mese:</span>\n            <select id=\"admin-filter-mes\" onchange=\"loadAdminCalendario()\">\n              <option value=\"1\">Gennaio</option><option value=\"2\">Febbraio</option><option value=\"3\">Marzo</option>\n              <option value=\"4\">Aprile</option><option value=\"5\">Maggio</option><option value=\"6\">Giugno</option>\n              <option value=\"7\">Luglio</option><option value=\"8\">Agosto</option><option value=\"9\">Settembre</option>\n              <option value=\"10\">Ottobre</option><option value=\"11\">Novembre</option><option value=\"12\">Dicembre</option>\n            </select>\n            <select id=\"admin-filter-anno\" onchange=\"loadAdminCalendario()\">\n              <option value=\"2025\">2025</option><option value=\"2026\">2026</option><option value=\"2027\">2027</option>\n            </select>\n          </div>\n          <div class=\"filter-group\">\n            <span class=\"filter-label\">Collaboratore:</span>\n            <select id=\"admin-filter-enc\" onchange=\"loadAdminCalendario()\">\n              <option value=\"\">Tutti</option>\n            </select>\n          </div>\n          <div class=\"filter-group\">\n            <span class=\"filter-label\">Appartamento:</span>\n            <select id=\"admin-filter-apt\" onchange=\"loadAdminCalendario()\">\n              <option value=\"\">Tutti</option>\n            </select>\n          </div>\n          <div style=\"margin-left:auto;\">\n            <button class=\"btn btn-primary\" onclick=\"openModalAddRiga()\">➕ Nuova Pulizia</button>\n          </div>\n        </div>\n\n        <div class=\"card\">\n          <div class=\"table-container\">\n            <table id=\"table-admin-calendario\">\n              <thead>\n                <tr>\n                  <th>Appartamento</th>\n                  <th>Data Pulizia</th>\n                  <th>Orario</th>\n                  <th>Preparare Per</th>\n                  <th>Entrata Prossimi</th>\n                  <th>Note Admin</th>\n                  <th>Incaricato/a</th>\n                  <th>Note Incaricato</th>\n                  <th style=\"text-align:right;\">Azioni</th>\n                </tr>\n              </thead>\n              <tbody id=\"tbody-admin-calendario\">\n                <tr><td colspan=\"9\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento dati...</td></tr>\n              </tbody>\n            </table>\n          </div>\n        </div>\n      </div>\n\n      <!-- Tab Admin: Collaboratori -->\n      <div id=\"admin-tab-collaboratori\" class=\"admin-subview\" style=\"display:none;\">\n        <div style=\"display:flex; justify-content:space-between; align-items:center; margin-bottom:16px;\">\n          <h3 style=\"font-size:1.1rem; font-weight:700;\">Gestione Account Collaboratori</h3>\n          <button class=\"btn btn-primary\" onclick=\"openModalAddUser()\">➕ Nuovo Collaboratore</button>\n        </div>\n\n        <div class=\"card\">\n          <div class=\"table-container\">\n            <table>\n              <thead>\n                <tr>\n                  <th>Nome</th>\n                  <th>Email</th>\n                  <th>Incaricato Key</th>\n                  <th>Ruolo</th>\n                  <th>Stato</th>\n                  <th>Ultimo Accesso</th>\n                  <th style=\"text-align:right;\">Azioni</th>\n                </tr>\n              </thead>\n              <tbody id=\"tbody-admin-utenti\">\n                <tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento collaboratori...</td></tr>\n              </tbody>\n            </table>\n          </div>\n        </div>\n      </div>\n\n      <!-- Tab Admin: Contabilità -->\n      <div id=\"admin-tab-contabilita\" class=\"admin-subview\" style=\"display:none;\">\n        <div class=\"stats-grid\">\n          <div class=\"stat-card\">\n            <span class=\"stat-label\">Totale Compensi</span>\n            <span class=\"stat-value blue\" id=\"admin-cont-totale\">€ 0.00</span>\n          </div>\n          <div class=\"stat-card\">\n            <span class=\"stat-label\">Già Pagato</span>\n            <span class=\"stat-value success\" id=\"admin-cont-pagato\">€ 0.00</span>\n          </div>\n          <div class=\"stat-card\">\n            <span class=\"stat-label\">In Attesa</span>\n            <span class=\"stat-value warning\" id=\"admin-cont-attesa\">€ 0.00</span>\n          </div>\n          <div class=\"stat-card\">\n            <span class=\"stat-label\">Totale Pulizie</span>\n            <span class=\"stat-value primary\" id=\"admin-cont-pulizie\">0</span>\n          </div>\n        </div>\n\n        <div class=\"filter-bar\">\n          <div class=\"filter-group\">\n            <span class=\"filter-label\">Mese:</span>\n            <select id=\"admin-cont-filter-mes\" onchange=\"loadAdminContabilita()\">\n              <option value=\"\">Tutti i mesi</option>\n              <option value=\"1\">Gennaio</option><option value=\"2\">Febbraio</option><option value=\"3\">Marzo</option>\n              <option value=\"4\">Aprile</option><option value=\"5\">Maggio</option><option value=\"6\">Giugno</option>\n              <option value=\"7\">Luglio</option><option value=\"8\">Agosto</option><option value=\"9\">Settembre</option>\n              <option value=\"10\">Ottobre</option><option value=\"11\">Novembre</option><option value=\"12\">Dicembre</option>\n            </select>\n            <select id=\"admin-cont-filter-anno\" onchange=\"loadAdminContabilita()\">\n              <option value=\"2025\">2025</option><option value=\"2026\" selected>2026</option><option value=\"2027\">2027</option>\n            </select>\n          </div>\n          <div class=\"filter-group\">\n            <span class=\"filter-label\">Collaboratore:</span>\n            <select id=\"admin-cont-filter-enc\" onchange=\"loadAdminContabilita()\">\n              <option value=\"\">Tutti</option>\n            </select>\n          </div>\n          <div style=\"margin-left:auto; display:flex; gap:8px;\">\n            <button class=\"btn btn-primary\" onclick=\"openModalCalcolaComp()\">⚡ Calcola da Calendario</button>\n            <button class=\"btn btn-secondary\" onclick=\"openModalAddCont()\">➕ Aggiungi Record</button>\n          </div>\n        </div>\n\n        <div class=\"card\">\n          <div class=\"table-container\">\n            <table>\n              <thead>\n                <tr>\n                  <th>Periodo</th>\n                  <th>Collaboratore</th>\n                  <th>N. Pulizie</th>\n                  <th>Prezzo/cad</th>\n                  <th>Spese Extra</th>\n                  <th>Totale</th>\n                  <th>Stato</th>\n                  <th>Data Pagamento</th>\n                  <th>Note</th>\n                  <th style=\"text-align:right;\">Azioni</th>\n                </tr>\n              </thead>\n              <tbody id=\"tbody-admin-contabilita\">\n                <tr><td colspan=\"10\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento contabilità...</td></tr>\n              </tbody>\n            </table>\n          </div>\n        </div>\n      </div>\n\n      <!-- Tab Admin: Configurazione -->\n      <div id=\"admin-tab-configurazione\" class=\"admin-subview\" style=\"display:none;\">\n        <div style=\"display:grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap:20px;\">\n          <div class=\"card\">\n            <div class=\"card-title\">🏢 Appartamenti Gestiti</div>\n            <p style=\"color:var(--text-muted); font-size:0.84rem; margin-bottom:12px;\">Un appartamento per riga</p>\n            <textarea id=\"cfg-apartamentos\" rows=\"8\" style=\"width:100%; margin-bottom:14px;\"></textarea>\n            <button class=\"btn btn-primary\" onclick=\"salvaConfigurazioneApp()\">Salva Appartamenti</button>\n          </div>\n\n          <div class=\"card\">\n            <div class=\"card-title\">👥 Nomi Incaricati</div>\n            <p style=\"color:var(--text-muted); font-size:0.84rem; margin-bottom:12px;\">Un incaricato per riga (es: DANIELA, LINA)</p>\n            <textarea id=\"cfg-encargados\" rows=\"8\" style=\"width:100%; margin-bottom:14px;\"></textarea>\n            <button class=\"btn btn-primary\" onclick=\"salvaConfigurazioneApp()\">Salva Incaricati</button>\n          </div>\n        </div>\n      </div>\n\n      <!-- Tab Admin: Profilo -->\n      <div id=\"admin-tab-profilo\" class=\"admin-subview\" style=\"display:none;\">\n        <div class=\"card\" style=\"max-width:500px; margin:0 auto;\">\n          <div class=\"card-title\">👤 Profilo Amministratore</div>\n          <form onsubmit=\"handleCambiaPassword(event)\">\n            <div class=\"form-group\">\n              <label>Vecchia Password</label>\n              <input type=\"password\" id=\"admin-old-pwd\" required>\n            </div>\n            <div class=\"form-group\">\n              <label>Nuova Password</label>\n              <input type=\"password\" id=\"admin-new-pwd\" required minlength=\"6\">\n            </div>\n            <div class=\"form-group\" style=\"margin-bottom:20px;\">\n              <label>Conferma Nuova Password</label>\n              <input type=\"password\" id=\"admin-conf-pwd\" required minlength=\"6\">\n            </div>\n            <button type=\"submit\" class=\"btn btn-primary\" style=\"width:100%;\">Aggiorna Password</button>\n          </form>\n        </div>\n      </div>\n    </div>\n\n    <!-- ==================== VIEW: COLLABORATORE ==================== -->\n    <div id=\"view-collaborator\" class=\"view-section\">\n      <div class=\"tabs-nav\">\n        <button class=\"tab-btn active\" onclick=\"switchCollabTab('calendario')\">📅 Il Mio Calendario</button>\n        <button class=\"tab-btn\" onclick=\"switchCollabTab('compensi')\">💰 I Miei Compensi</button>\n        <button class=\"tab-btn\" onclick=\"switchCollabTab('profilo')\">👤 Profilo</button>\n      </div>\n\n      <!-- Tab Collab: Calendario Personale -->\n      <div id=\"collab-tab-calendario\" class=\"collab-subview\">\n        <div class=\"filter-bar\">\n          <div class=\"filter-group\">\n            <span class=\"filter-label\">Mese:</span>\n            <select id=\"collab-filter-mes\" onchange=\"loadCollabCalendario()\">\n              <option value=\"1\">Gennaio</option><option value=\"2\">Febbraio</option><option value=\"3\">Marzo</option>\n              <option value=\"4\">Aprile</option><option value=\"5\">Maggio</option><option value=\"6\">Giugno</option>\n              <option value=\"7\">Luglio</option><option value=\"8\">Agosto</option><option value=\"9\">Settembre</option>\n              <option value=\"10\">Ottobre</option><option value=\"11\">Novembre</option><option value=\"12\">Dicembre</option>\n            </select>\n            <select id=\"collab-filter-anno\" onchange=\"loadCollabCalendario()\">\n              <option value=\"2025\">2025</option><option value=\"2026\">2026</option><option value=\"2027\">2027</option>\n            </select>\n          </div>\n          <div style=\"margin-left:auto;\">\n            <span class=\"badge badge-enc\" id=\"collab-count-badge\">0 Pulizie</span>\n          </div>\n        </div>\n\n        <div class=\"card\">\n          <div class=\"table-container\">\n            <table>\n              <thead>\n                <tr>\n                  <th>Appartamento</th>\n                  <th>Data Pulizia</th>\n                  <th>Orario</th>\n                  <th>Preparare Per</th>\n                  <th>Entrata Prossimi</th>\n                  <th>Note Admin</th>\n                  <th>La Mia Nota</th>\n                </tr>\n              </thead>\n              <tbody id=\"tbody-collab-calendario\">\n                <tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento del tuo calendario...</td></tr>\n              </tbody>\n            </table>\n          </div>\n        </div>\n      </div>\n\n      <!-- Tab Collab: Compensi Personali -->\n      <div id=\"collab-tab-compensi\" class=\"collab-subview\" style=\"display:none;\">\n        <div class=\"stats-grid\">\n          <div class=\"stat-card\">\n            <span class=\"stat-label\">Pulizie Totali</span>\n            <span class=\"stat-value primary\" id=\"collab-stat-pulizie\">0</span>\n          </div>\n          <div class=\"stat-card\">\n            <span class=\"stat-label\">Compenso Totale</span>\n            <span class=\"stat-value blue\" id=\"collab-stat-totale\">€ 0.00</span>\n          </div>\n          <div class=\"stat-card\">\n            <span class=\"stat-label\">Già Saldato</span>\n            <span class=\"stat-value success\" id=\"collab-stat-pagato\">€ 0.00</span>\n          </div>\n          <div class=\"stat-card\">\n            <span class=\"stat-label\">In Attesa</span>\n            <span class=\"stat-value warning\" id=\"collab-stat-attesa\">€ 0.00</span>\n          </div>\n        </div>\n\n        <div class=\"card\">\n          <div class=\"card-title\">Storico Compensi e Pagamenti</div>\n          <div class=\"table-container\">\n            <table>\n              <thead>\n                <tr>\n                  <th>Periodo</th>\n                  <th>N. Pulizie</th>\n                  <th>Tariffa</th>\n                  <th>Spese Extra</th>\n                  <th>Totale</th>\n                  <th>Stato</th>\n                  <th>Data Pagamento</th>\n                  <th>Note</th>\n                </tr>\n              </thead>\n              <tbody id=\"tbody-collab-compensi\">\n                <tr><td colspan=\"8\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento compensi...</td></tr>\n              </tbody>\n            </table>\n          </div>\n        </div>\n      </div>\n\n      <!-- Tab Collab: Profilo -->\n      <div id=\"collab-tab-profilo\" class=\"collab-subview\" style=\"display:none;\">\n        <div class=\"card\" style=\"max-width:500px; margin:0 auto;\">\n          <div class=\"card-title\">👤 Il Mio Profilo</div>\n          <div style=\"margin-bottom:20px; font-size:0.9rem; line-height:1.6;\">\n            <div><strong>Nome:</strong> <span id=\"collab-info-nome\">-</span></div>\n            <div><strong>Email:</strong> <span id=\"collab-info-email\">-</span></div>\n            <div><strong>Incaricato:</strong> <span class=\"badge badge-enc\" id=\"collab-info-enc\">-</span></div>\n          </div>\n          <hr style=\"border:none; border-top:1px solid var(--surface-border); margin-bottom:20px;\">\n          <h4 style=\"font-size:0.95rem; margin-bottom:12px;\">Cambia Password</h4>\n          <form onsubmit=\"handleCambiaPassword(event)\">\n            <div class=\"form-group\">\n              <label>Vecchia Password</label>\n              <input type=\"password\" id=\"collab-old-pwd\" required>\n            </div>\n            <div class=\"form-group\">\n              <label>Nuova Password</label>\n              <input type=\"password\" id=\"collab-new-pwd\" required minlength=\"6\">\n            </div>\n            <div class=\"form-group\" style=\"margin-bottom:20px;\">\n              <label>Conferma Nuova Password</label>\n              <input type=\"password\" id=\"collab-conf-pwd\" required minlength=\"6\">\n            </div>\n            <button type=\"submit\" class=\"btn btn-primary\" style=\"width:100%;\">Aggiorna Password</button>\n          </form>\n        </div>\n      </div>\n    </div>\n\n  </main>\n\n  <!-- ==================== MODALS ==================== -->\n\n  <!-- Modal: Nuova/Modifica Pulizia (Admin) -->\n  <div class=\"modal-overlay\" id=\"modal-riga\">\n    <div class=\"modal\">\n      <div class=\"modal-header\">\n        <div class=\"modal-title\" id=\"modal-riga-title\">Nuova Pulizia</div>\n        <button class=\"modal-close\" onclick=\"closeModal('modal-riga')\">✕</button>\n      </div>\n      <form onsubmit=\"handleSaveRiga(event)\">\n        <input type=\"hidden\" id=\"riga-id\">\n        <div class=\"form-group\">\n          <label>Appartamento</label>\n          <select id=\"riga-apt\" required></select>\n        </div>\n        <div class=\"form-group\">\n          <label>Data Pulizia</label>\n          <input type=\"date\" id=\"riga-fecha\" required>\n        </div>\n        <div class=\"form-group\">\n          <label>Orario</label>\n          <input type=\"text\" id=\"riga-horario\" placeholder=\"Es: 11:00\">\n        </div>\n        <div class=\"form-group\">\n          <label>Preparare Per</label>\n          <input type=\"text\" id=\"riga-preparar\" placeholder=\"Es: 2 PAX (Matrimoniale)\">\n        </div>\n        <div class=\"form-group\">\n          <label>Entrata Prossimi Ospiti</label>\n          <input type=\"text\" id=\"riga-entrada\" placeholder=\"Es: 15:00\">\n        </div>\n        <div class=\"form-group\">\n          <label>Note Admin</label>\n          <input type=\"text\" id=\"riga-notas\" placeholder=\"Es: Cambio lenzuola completo\">\n        </div>\n        <div class=\"form-group\" style=\"margin-bottom:20px;\">\n          <label>Incaricato/a</label>\n          <select id=\"riga-encargado\" required></select>\n        </div>\n        <div style=\"display:flex; justify-content:flex-end; gap:10px;\">\n          <button type=\"button\" class=\"btn\" onclick=\"closeModal('modal-riga')\">Annulla</button>\n          <button type=\"submit\" class=\"btn btn-primary\" id=\"btn-save-riga\">Salva Pulizia</button>\n        </div>\n      </form>\n    </div>\n  </div>\n\n  <!-- Modal: Nota Incaricato -->\n  <div class=\"modal-overlay\" id=\"modal-nota-enc\">\n    <div class=\"modal\">\n      <div class=\"modal-header\">\n        <div class=\"modal-title\">Nota Incaricato</div>\n        <button class=\"modal-close\" onclick=\"closeModal('modal-nota-enc')\">✕</button>\n      </div>\n      <form onsubmit=\"handleSaveNotaEnc(event)\">\n        <input type=\"hidden\" id=\"nota-enc-id\">\n        <div class=\"form-group\" style=\"margin-bottom:20px;\">\n          <label>Note sulla pulizia (es: segnalazioni, guasti, rifornimenti)</label>\n          <textarea id=\"nota-enc-text\" rows=\"4\" style=\"width:100%;\" placeholder=\"Inserisci qui eventuali note...\"></textarea>\n        </div>\n        <div style=\"display:flex; justify-content:flex-end; gap:10px;\">\n          <button type=\"button\" class=\"btn\" onclick=\"closeModal('modal-nota-enc')\">Annulla</button>\n          <button type=\"submit\" class=\"btn btn-primary\">Salva Nota</button>\n        </div>\n      </form>\n    </div>\n  </div>\n\n  <!-- Modal: Nuovo Collaboratore (Admin) -->\n  <div class=\"modal-overlay\" id=\"modal-user\">\n    <div class=\"modal\">\n      <div class=\"modal-header\">\n        <div class=\"modal-title\">Nuovo Collaboratore</div>\n        <button class=\"modal-close\" onclick=\"closeModal('modal-user')\">✕</button>\n      </div>\n      <form onsubmit=\"handleCreateUser(event)\">\n        <div class=\"form-group\">\n          <label>Email</label>\n          <input type=\"email\" id=\"user-email\" required placeholder=\"collaboratore@pulizie.it\">\n        </div>\n        <div class=\"form-group\">\n          <label>Nome Completo</label>\n          <input type=\"text\" id=\"user-nombre\" required placeholder=\"Es: Maria Rossi\">\n        </div>\n        <div class=\"form-group\">\n          <label>Chiave Incaricato (Nome nel Calendario)</label>\n          <select id=\"user-enckey\" required></select>\n        </div>\n        <div class=\"form-group\" style=\"margin-bottom:20px;\">\n          <label>Password Iniziale</label>\n          <div style=\"display:flex; gap:8px;\">\n            <input type=\"text\" id=\"user-password\" required minlength=\"6\" style=\"flex:1;\">\n            <button type=\"button\" class=\"btn btn-sm\" onclick=\"generaPasswordCasuale('user-password')\">🎲 Genera</button>\n          </div>\n        </div>\n        <div style=\"display:flex; justify-content:flex-end; gap:10px;\">\n          <button type=\"button\" class=\"btn\" onclick=\"closeModal('modal-user')\">Annulla</button>\n          <button type=\"submit\" class=\"btn btn-primary\">Crea Account</button>\n        </div>\n      </form>\n    </div>\n  </div>\n\n  <!-- Modal: Reset Password (Admin) -->\n  <div class=\"modal-overlay\" id=\"modal-reset-pwd\">\n    <div class=\"modal\">\n      <div class=\"modal-header\">\n        <div class=\"modal-title\">Reset Password Collaboratore</div>\n        <button class=\"modal-close\" onclick=\"closeModal('modal-reset-pwd')\">✕</button>\n      </div>\n      <form onsubmit=\"handleResetPasswordSubmit(event)\">\n        <input type=\"hidden\" id=\"reset-user-email\">\n        <p style=\"color:var(--text-muted); font-size:0.86rem; margin-bottom:14px;\">Reimposta la password per: <strong id=\"reset-user-email-label\"></strong></p>\n        <div class=\"form-group\" style=\"margin-bottom:20px;\">\n          <label>Nuova Password</label>\n          <div style=\"display:flex; gap:8px;\">\n            <input type=\"text\" id=\"reset-new-password\" required minlength=\"6\" style=\"flex:1;\">\n            <button type=\"button\" class=\"btn btn-sm\" onclick=\"generaPasswordCasuale('reset-new-password')\">🎲 Genera</button>\n          </div>\n        </div>\n        <div style=\"display:flex; justify-content:flex-end; gap:10px;\">\n          <button type=\"button\" class=\"btn\" onclick=\"closeModal('modal-reset-pwd')\">Annulla</button>\n          <button type=\"submit\" class=\"btn btn-primary\">Salva Nuova Password</button>\n        </div>\n      </form>\n    </div>\n  </div>\n\n  <!-- Modal: Calcola Compenso da Calendario -->\n  <div class=\"modal-overlay\" id=\"modal-calc-comp\">\n    <div class=\"modal\">\n      <div class=\"modal-header\">\n        <div class=\"modal-title\">⚡ Calcola Compenso da Calendario</div>\n        <button class=\"modal-close\" onclick=\"closeModal('modal-calc-comp')\">✕</button>\n      </div>\n      <form onsubmit=\"handleExecCalcolaComp(event)\">\n        <div class=\"form-group\">\n          <label>Collaboratore</label>\n          <select id=\"calc-encargado\" required></select>\n        </div>\n        <div style=\"display:grid; grid-template-columns: 1fr 1fr; gap:12px;\">\n          <div class=\"form-group\">\n            <label>Mese</label>\n            <select id=\"calc-mes\" required>\n              <option value=\"1\">Gennaio</option><option value=\"2\">Febbraio</option><option value=\"3\">Marzo</option>\n              <option value=\"4\">Aprile</option><option value=\"5\">Maggio</option><option value=\"6\">Giugno</option>\n              <option value=\"7\">Luglio</option><option value=\"8\">Agosto</option><option value=\"9\">Settembre</option>\n              <option value=\"10\">Ottobre</option><option value=\"11\">Novembre</option><option value=\"12\">Dicembre</option>\n            </select>\n          </div>\n          <div class=\"form-group\">\n            <label>Anno</label>\n            <select id=\"calc-anno\" required>\n              <option value=\"2025\">2025</option><option value=\"2026\" selected>2026</option><option value=\"2027\">2027</option>\n            </select>\n          </div>\n        </div>\n        <div class=\"form-group\">\n          <label>Tariffa per Pulizia (€)</label>\n          <input type=\"number\" step=\"0.5\" id=\"calc-tarifa\" value=\"40\" required>\n        </div>\n        <div id=\"calc-result-box\" style=\"display:none; background:rgba(20,184,166,0.1); border:1px solid var(--primary); border-radius:8px; padding:12px; margin-bottom:16px;\">\n          <div style=\"font-size:0.88rem; font-weight:600; color:var(--primary);\">Pulizie trovate: <span id=\"calc-count\">0</span></div>\n          <div style=\"font-size:1.1rem; font-weight:700; margin-top:4px;\">Totale Calcolato: <span id=\"calc-totale\">€ 0.00</span></div>\n        </div>\n        <div style=\"display:flex; justify-content:flex-end; gap:10px;\">\n          <button type=\"button\" class=\"btn\" onclick=\"closeModal('modal-calc-comp')\">Chiudi</button>\n          <button type=\"button\" class=\"btn\" onclick=\"anteprimaCalcolo()\">Calcola Anteprima</button>\n          <button type=\"submit\" class=\"btn btn-primary\" id=\"btn-salva-calc-comp\" style=\"display:none;\">Salva in Contabilità</button>\n        </div>\n      </form>\n    </div>\n  </div>\n\n  <!-- Modal: Aggiorna Stato Pagamento -->\n  <div class=\"modal-overlay\" id=\"modal-stato-cont\">\n    <div class=\"modal\">\n      <div class=\"modal-header\">\n        <div class=\"modal-title\">Aggiorna Stato Pagamento</div>\n        <button class=\"modal-close\" onclick=\"closeModal('modal-stato-cont')\">✕</button>\n      </div>\n      <form onsubmit=\"handleUpdateStatoCont(event)\">\n        <input type=\"hidden\" id=\"stato-cont-id\">\n        <div class=\"form-group\">\n          <label>Stato Pagamento</label>\n          <select id=\"stato-cont-estado\">\n            <option value=\"pendiente\">🟡 In Attesa</option>\n            <option value=\"pagato\">🟢 Pagato</option>\n          </select>\n        </div>\n        <div class=\"form-group\">\n          <label>Data Pagamento</label>\n          <input type=\"date\" id=\"stato-cont-fecha\">\n        </div>\n        <div class=\"form-group\" style=\"margin-bottom:20px;\">\n          <label>Note Pagamento</label>\n          <input type=\"text\" id=\"stato-cont-notas\" placeholder=\"Es: Bonifico effettuato il 05/11\">\n        </div>\n        <div style=\"display:flex; justify-content:flex-end; gap:10px;\">\n          <button type=\"button\" class=\"btn\" onclick=\"closeModal('modal-stato-cont')\">Annulla</button>\n          <button type=\"submit\" class=\"btn btn-primary\">Salva Stato</button>\n        </div>\n      </form>\n    </div>\n  </div>\n\n  <!-- Modal: Configurazione Script URL -->\n  <div class=\"modal-overlay\" id=\"modal-script-url\">\n    <div class=\"modal\">\n      <div class=\"modal-header\">\n        <div class=\"modal-title\">⚙️ Connessione Backend Google Apps Script</div>\n        <button class=\"modal-close\" onclick=\"closeModal('modal-script-url')\">✕</button>\n      </div>\n      <div class=\"form-group\" style=\"margin-bottom:20px;\">\n        <label>URL Web App Apps Script</label>\n        <input type=\"text\" id=\"cfg-script-url\" style=\"width:100%;\">\n        <span style=\"font-size:0.75rem; color:var(--text-muted); margin-top:4px;\">Inserisci l'URL del deployment Web App che termina con /exec</span>\n      </div>\n      <div style=\"display:flex; justify-content:flex-end; gap:10px;\">\n        <button type=\"button\" class=\"btn\" onclick=\"closeModal('modal-script-url')\">Annulla</button>\n        <button type=\"button\" class=\"btn btn-primary\" onclick=\"salvaScriptUrl()\">Salva e Riconnetti</button>\n      </div>\n    </div>\n  </div>\n\n  <!-- Toast Container -->\n  <div id=\"toast-container\"></div>\n\n  <!-- ==================== APPLICATION SCRIPT ==================== -->\n  <script>\n    // --- Configurazione Iniziale ---\n    const DEFAULT_SCRIPT_URL = \"https://script.google.com/macros/s/AKfycbwxpUbcKCoxEqKB6rBy2FvW_FQXVROdl6Gkakc9gtXhpDswRdXtml2M4eyXF5XOj340AA/exec\";\n    let SCRIPT_URL = localStorage.getItem(\"CP_SCRIPT_URL\") || DEFAULT_SCRIPT_URL;\n\n    // Sessione\n    let SESSION = {\n      token: localStorage.getItem(\"CP_TOKEN\") || null,\n      email: localStorage.getItem(\"CP_EMAIL\") || null,\n      nombre: localStorage.getItem(\"CP_NOMBRE\") || null,\n      role: localStorage.getItem(\"CP_ROLE\") || null,\n      encargadoKey: localStorage.getItem(\"CP_ENCKEY\") || null\n    };\n\n    // Cache configurazione\n    let APP_CONFIG = { apartamentos: [], encargados: [] };\n    let CURRENT_CALENDARIO_ROWS = [];\n\n    // --- PWA Service Worker & Install Prompt ---\n    let deferredInstallPrompt = null;\n    if ('serviceWorker' in navigator) {\n      window.addEventListener('load', () => {\n        navigator.serviceWorker.register('./sw.js').catch(err => console.log('SW reg error:', err));\n      });\n    }\n\n    window.addEventListener('beforeinstallprompt', (e) => {\n      e.preventDefault();\n      deferredInstallPrompt = e;\n      const banner = document.getElementById('pwa-install-banner');\n      if (banner) banner.style.display = 'flex';\n    });\n\n    document.getElementById('btn-install-pwa').addEventListener('click', async () => {\n      if (deferredInstallPrompt) {\n        deferredInstallPrompt.prompt();\n        const { outcome } = await deferredInstallPrompt.userChoice;\n        if (outcome === 'accepted') {\n          dismissPwaBanner();\n        }\n        deferredInstallPrompt = null;\n      } else {\n        // Istruzioni per iOS o browser che non supportano prompt programmatico\n        alert(\"Per installare su iOS:\\n1. Tocca il pulsante Condividi (icona quadrato con freccia in alto)\\n2. Scorri e tocca 'Aggiungi alla schermata Home' ➕\");\n      }\n    });\n\n    function dismissPwaBanner() {\n      document.getElementById('pwa-install-banner').style.display = 'none';\n    }\n\n    // --- Utility Funzioni ---\n    function toast(message, type = 'info') {\n      const container = document.getElementById('toast-container');\n      const t = document.createElement('div');\n      t.className = 'toast ' + type;\n      t.innerText = message;\n      container.appendChild(t);\n      setTimeout(() => { t.remove(); }, 4000);\n    }\n\n    function setSyncing(isSyncing) {\n      const dot = document.getElementById('sync-dot');\n      if (dot) {\n        if (isSyncing) dot.classList.add('syncing');\n        else dot.classList.remove('syncing');\n      }\n    }\n\n    function openModal(id) {\n      const el = document.getElementById(id);\n      if (el) el.classList.add('active');\n    }\n    function closeModal(id) {\n      const el = document.getElementById(id);\n      if (el) el.classList.remove('active');\n    }\n\n    function generaPasswordCasuale(targetInputId) {\n      const chars = \"ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!@#$%\";\n      let pwd = \"\";\n      for (let i = 0; i < 10; i++) {\n        pwd += chars.charAt(Math.floor(Math.random() * chars.length));\n      }\n      document.getElementById(targetInputId).value = pwd;\n    }\n\n    // --- API CLIENT (POST con CORS Friendly text/plain) ---\n    async function apiCall(action, payload = {}) {\n      setSyncing(true);\n      const bodyData = Object.assign({ action: action, token: SESSION.token }, payload);\n      try {\n        const response = await fetch(SCRIPT_URL, {\n          method: 'POST',\n          headers: { 'Content-Type': 'text/plain;charset=utf-8' },\n          body: JSON.stringify(bodyData)\n        });\n        const json = await response.json();\n        setSyncing(false);\n        return json;\n      } catch (err) {\n        setSyncing(false);\n        console.error(\"API error (\" + action + \"):\", err);\n        throw err;\n      }\n    }\n\n    // --- Avvio e Routing Viste ---\n    window.addEventListener('DOMContentLoaded', async () => {\n      // Imposta data corrente nei selettori\n      const now = new Date();\n      const currentMonth = now.getMonth() + 1;\n      const currentYear = now.getFullYear();\n\n      ['admin-filter-mes', 'admin-cont-filter-mes', 'collab-filter-mes', 'calc-mes'].forEach(id => {\n        const el = document.getElementById(id);\n        if (el && el.value === \"\") el.value = currentMonth;\n        else if (el) el.value = currentMonth;\n      });\n      ['admin-filter-anno', 'admin-cont-filter-anno', 'collab-filter-anno', 'calc-anno'].forEach(id => {\n        const el = document.getElementById(id);\n        if (el) el.value = currentYear;\n      });\n\n      document.getElementById('cfg-script-url').value = SCRIPT_URL;\n\n      // Verifica sessione salvata\n      if (SESSION.token) {\n        try {\n          const res = await apiCall(\"verificarSesion\");\n          if (res && res.valid) {\n            SESSION.role = res.role;\n            SESSION.email = res.email;\n            SESSION.nombre = res.nombre;\n            SESSION.encargadoKey = res.encargadoKey;\n            showAppView();\n            return;\n          }\n        } catch (e) {\n          console.warn(\"Verifica sessione fallita, richiesta login\");\n        }\n      }\n      showLoginView();\n    });\n\n    function showLoginView() {\n      document.querySelectorAll('.view-section').forEach(el => el.classList.remove('active'));\n      document.getElementById('view-login').classList.add('active');\n      document.getElementById('user-status-pill').style.display = 'none';\n      document.getElementById('btn-logout').style.display = 'none';\n    }\n\n    function showAppView() {\n      document.querySelectorAll('.view-section').forEach(el => el.classList.remove('active'));\n      document.getElementById('user-status-pill').style.display = 'flex';\n      document.getElementById('btn-logout').style.display = 'inline-flex';\n      document.getElementById('header-user-name').innerText = SESSION.nombre || SESSION.email;\n      document.getElementById('header-user-role').innerText = SESSION.role === 'admin' ? 'Admin' : 'Collaboratore';\n\n      if (SESSION.role === 'admin') {\n        document.getElementById('view-admin').classList.add('active');\n        initAdminView();\n      } else {\n        document.getElementById('view-collaborator').classList.add('active');\n        initCollabView();\n      }\n    }\n\n    // --- Login & Logout ---\n    async function handleLoginSubmit(e) {\n      e.preventDefault();\n      const email = document.getElementById('login-email').value.trim();\n      const password = document.getElementById('login-password').value;\n      const btn = document.getElementById('btn-login-submit');\n\n      btn.disabled = true;\n      btn.innerText = \"Accesso in corso...\";\n\n      try {\n        const res = await apiCall(\"login\", { email: email, password: password });\n        btn.disabled = false;\n        btn.innerText = \"Accedi\";\n\n        if (res && res.success) {\n          SESSION.token = res.token;\n          SESSION.email = res.user.email;\n          SESSION.nombre = res.user.nombre;\n          SESSION.role = res.user.role;\n          SESSION.encargadoKey = res.user.encargadoKey;\n\n          localStorage.setItem(\"CP_TOKEN\", SESSION.token);\n          localStorage.setItem(\"CP_EMAIL\", SESSION.email);\n          localStorage.setItem(\"CP_NOMBRE\", SESSION.nombre);\n          localStorage.setItem(\"CP_ROLE\", SESSION.role);\n          localStorage.setItem(\"CP_ENCKEY\", SESSION.encargadoKey || \"\");\n\n          toast(\"Accesso effettuato con successo!\", \"success\");\n          showAppView();\n        } else {\n          toast(res.error || \"Credenziali non valide\", \"error\");\n        }\n      } catch (err) {\n        btn.disabled = false;\n        btn.innerText = \"Accedi\";\n        toast(\"Errore di connessione: verifica l'URL dello script\", \"error\");\n      }\n    }\n\n    async function appLogout() {\n      try { await apiCall(\"logout\"); } catch(e){}\n      SESSION = { token: null, email: null, nombre: null, role: null, encargadoKey: null };\n      localStorage.removeItem(\"CP_TOKEN\");\n      localStorage.removeItem(\"CP_EMAIL\");\n      localStorage.removeItem(\"CP_NOMBRE\");\n      localStorage.removeItem(\"CP_ROLE\");\n      localStorage.removeItem(\"CP_ENCKEY\");\n      toast(\"Disconnesso\", \"info\");\n      showLoginView();\n    }\n\n    function salvaScriptUrl() {\n      const val = document.getElementById('cfg-script-url').value.trim();\n      if (!val) return;\n      SCRIPT_URL = val;\n      localStorage.setItem(\"CP_SCRIPT_URL\", val);\n      closeModal('modal-script-url');\n      toast(\"URL script aggiornato\", \"success\");\n    }\n\n    // --- Tab Navigation ---\n    function switchAdminTab(tabName) {\n      document.querySelectorAll('#view-admin .tab-btn').forEach(b => b.classList.remove('active'));\n      document.querySelectorAll('.admin-subview').forEach(s => s.style.display = 'none');\n      event.currentTarget.classList.add('active');\n      document.getElementById('admin-tab-' + tabName).style.display = 'block';\n\n      if (tabName === 'calendario') loadAdminCalendario();\n      if (tabName === 'collaboratori') loadAdminUtenti();\n      if (tabName === 'contabilita') loadAdminContabilita();\n      if (tabName === 'configurazione') loadAdminConfig();\n    }\n\n    function switchCollabTab(tabName) {\n      document.querySelectorAll('#view-collaborator .tab-btn').forEach(b => b.classList.remove('active'));\n      document.querySelectorAll('.collab-subview').forEach(s => s.style.display = 'none');\n      event.currentTarget.classList.add('active');\n      document.getElementById('collab-tab-' + tabName).style.display = 'block';\n\n      if (tabName === 'calendario') loadCollabCalendario();\n      if (tabName === 'compensi') loadCollabCompensi();\n    }\n\n    // --- Inizializzazione Viste ---\n    async function initAdminView() {\n      await loadAppConfig();\n      loadAdminCalendario();\n    }\n\n    async function initCollabView() {\n      document.getElementById('collab-info-nome').innerText = SESSION.nombre;\n      document.getElementById('collab-info-email').innerText = SESSION.email;\n      document.getElementById('collab-info-enc').innerText = SESSION.encargadoKey || \"Non assegnato\";\n      loadCollabCalendario();\n    }\n\n    async function loadAppConfig() {\n      try {\n        const res = await apiCall(\"getAppConfig\");\n        if (res && res.success) {\n          APP_CONFIG = res;\n          populateConfigDropdowns();\n        }\n      } catch (e) {\n        console.warn(\"Errore caricamento config:\", e);\n      }\n    }\n\n    function populateConfigDropdowns() {\n      const aptSelects = ['admin-filter-apt', 'riga-apt'];\n      aptSelects.forEach(id => {\n        const sel = document.getElementById(id);\n        if (!sel) return;\n        const currentVal = sel.value;\n        sel.innerHTML = id.includes('filter') ? '<option value=\"\">Tutti</option>' : '<option value=\"\">Seleziona...</option>';\n        (APP_CONFIG.apartamentos || []).forEach(apt => {\n          sel.innerHTML += '<option value=\"' + apt + '\">' + apt + '</option>';\n        });\n        if (currentVal) sel.value = currentVal;\n      });\n\n      const encSelects = ['admin-filter-enc', 'admin-cont-filter-enc', 'riga-encargado', 'user-enckey', 'calc-encargado'];\n      encSelects.forEach(id => {\n        const sel = document.getElementById(id);\n        if (!sel) return;\n        const currentVal = sel.value;\n        sel.innerHTML = id.includes('filter') ? '<option value=\"\">Tutti</option>' : '<option value=\"\">Seleziona...</option>';\n        (APP_CONFIG.encargados || []).forEach(enc => {\n          sel.innerHTML += '<option value=\"' + enc + '\">' + enc + '</option>';\n        });\n        if (currentVal) sel.value = currentVal;\n      });\n    }\n\n    // --- Calendario Admin ---\n    async function loadAdminCalendario() {\n      const mes = document.getElementById('admin-filter-mes').value;\n      const anno = document.getElementById('admin-filter-anno').value;\n      const enc = document.getElementById('admin-filter-enc').value;\n      const apt = document.getElementById('admin-filter-apt').value;\n      const tbody = document.getElementById('tbody-admin-calendario');\n\n      tbody.innerHTML = '<tr><td colspan=\"9\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento dati...</td></tr>';\n\n      try {\n        const res = await apiCall(\"getCalendario\", { mes: mes, anno: anno, userFilter: enc, aptFilter: apt });\n        if (res && res.success) {\n          CURRENT_CALENDARIO_ROWS = res.rows || [];\n          renderAdminCalendarioTable(CURRENT_CALENDARIO_ROWS);\n        } else {\n          tbody.innerHTML = '<tr><td colspan=\"9\" style=\"text-align:center; padding:30px; color:var(--danger);\">' + (res.error || \"Errore\") + '</td></tr>';\n        }\n      } catch (err) {\n        tbody.innerHTML = '<tr><td colspan=\"9\" style=\"text-align:center; padding:30px; color:var(--danger);\">Errore di connessione</td></tr>';\n      }\n    }\n\n    function renderAdminCalendarioTable(rows) {\n      const tbody = document.getElementById('tbody-admin-calendario');\n      if (!rows || rows.length === 0) {\n        tbody.innerHTML = '<tr><td colspan=\"9\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Nessuna pulizia programmata per questo periodo</td></tr>';\n        return;\n      }\n\n      let html = '';\n      rows.forEach(r => {\n        html += '<tr>';\n        html += '<td><span class=\"badge badge-apt\">' + (r.apartamento || '-') + '</span></td>';\n        html += '<td><strong>' + (r.fecha || '-') + '</strong></td>';\n        html += '<td>' + (r.horario || '-') + '</td>';\n        html += '<td>' + (r.preparar || '-') + '</td>';\n        html += '<td>' + (r.entrada || '-') + '</td>';\n        html += '<td>' + (r.notas || '-') + '</td>';\n        html += '<td><span class=\"badge badge-enc\">' + (r.encargado || 'Da assegnare') + '</span></td>';\n        html += '<td>' + (r.notaEncargado ? '<span style=\"color:var(--primary); font-size:0.8rem;\">💬 ' + r.notaEncargado + '</span>' : '<span style=\"color:var(--text-muted);\">-</span>') + '</td>';\n        html += '<td style=\"text-align:right; white-space:nowrap;\">';\n        html += '<button class=\"btn btn-sm\" onclick=\"openModalEditRiga(\\'' + r.id + '\\')\" style=\"margin-right:6px;\">✏️</button>';\n        html += '<button class=\"btn btn-sm btn-danger\" onclick=\"eliminaRiga(\\'' + r.id + '\\')\">🗑️</button>';\n        html += '</td>';\n        html += '</tr>';\n      });\n      tbody.innerHTML = html;\n    }\n\n    function openModalAddRiga() {\n      document.getElementById('modal-riga-title').innerText = \"Nuova Pulizia\";\n      document.getElementById('riga-id').value = \"\";\n      document.getElementById('riga-fecha').value = \"\";\n      document.getElementById('riga-horario').value = \"\";\n      document.getElementById('riga-preparar').value = \"\";\n      document.getElementById('riga-entrada').value = \"\";\n      document.getElementById('riga-notas').value = \"\";\n      openModal('modal-riga');\n    }\n\n    function openModalEditRiga(id) {\n      const r = CURRENT_CALENDARIO_ROWS.find(item => item.id == id);\n      if (!r) return;\n      document.getElementById('modal-riga-title').innerText = \"Modifica Pulizia\";\n      document.getElementById('riga-id').value = r.id;\n      document.getElementById('riga-apt').value = r.apartamento || \"\";\n      document.getElementById('riga-fecha').value = r.fecha || \"\";\n      document.getElementById('riga-horario').value = r.horario || \"\";\n      document.getElementById('riga-preparar').value = r.preparar || \"\";\n      document.getElementById('riga-entrada').value = r.entrada || \"\";\n      document.getElementById('riga-notas').value = r.notas || \"\";\n      document.getElementById('riga-encargado').value = r.encargado || \"\";\n      openModal('modal-riga');\n    }\n\n    async function handleSaveRiga(e) {\n      e.preventDefault();\n      const mes = document.getElementById('admin-filter-mes').value;\n      const anno = document.getElementById('admin-filter-anno').value;\n      const rowData = {\n        id: document.getElementById('riga-id').value,\n        apartamento: document.getElementById('riga-apt').value,\n        fecha: document.getElementById('riga-fecha').value,\n        horario: document.getElementById('riga-horario').value,\n        preparar: document.getElementById('riga-preparar').value,\n        entrada: document.getElementById('riga-entrada').value,\n        notas: document.getElementById('riga-notas').value,\n        encargado: document.getElementById('riga-encargado').value\n      };\n\n      try {\n        const res = await apiCall(\"saveRiga\", { rowData: rowData, mes: mes, anno: anno });\n        if (res && res.success) {\n          closeModal('modal-riga');\n          toast(\"Pulizia salvata con successo\", \"success\");\n          loadAdminCalendario();\n        } else {\n          toast(res.error || \"Errore nel salvataggio\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    async function eliminaRiga(id) {\n      if (!confirm(\"Sei sicuro di voler eliminare questa pulizia dal calendario?\")) return;\n      const mes = document.getElementById('admin-filter-mes').value;\n      const anno = document.getElementById('admin-filter-anno').value;\n      try {\n        const res = await apiCall(\"deleteRiga\", { id: id, mes: mes, anno: anno });\n        if (res && res.success) {\n          toast(\"Pulizia eliminata\", \"info\");\n          loadAdminCalendario();\n        } else {\n          toast(res.error || \"Errore eliminazione\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    // --- Gestione Collaboratori (Admin) ---\n    async function loadAdminUtenti() {\n      const tbody = document.getElementById('tbody-admin-utenti');\n      tbody.innerHTML = '<tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento collaboratori...</td></tr>';\n      try {\n        const res = await apiCall(\"getUsuarios\");\n        if (res && res.success) {\n          renderAdminUtentiTable(res.data || []);\n        } else {\n          tbody.innerHTML = '<tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--danger);\">' + (res.error || \"Errore\") + '</td></tr>';\n        }\n      } catch (e) {\n        tbody.innerHTML = '<tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--danger);\">Errore di rete</td></tr>';\n      }\n    }\n\n    function renderAdminUtentiTable(utenti) {\n      const tbody = document.getElementById('tbody-admin-utenti');\n      if (!utenti || utenti.length === 0) {\n        tbody.innerHTML = '<tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Nessun collaboratore registrato</td></tr>';\n        return;\n      }\n      let html = '';\n      utenti.forEach(u => {\n        const isBlocked = u.bloqueado === true;\n        html += '<tr>';\n        html += '<td><strong>' + (u.nombre || '-') + '</strong></td>';\n        html += '<td>' + (u.email || '-') + '</td>';\n        html += '<td><span class=\"badge badge-enc\">' + (u.encargadoKey || '-') + '</span></td>';\n        html += '<td>' + (u.role === 'admin' ? '<span class=\"badge badge-apt\">Admin</span>' : 'Collaboratore') + '</td>';\n        html += '<td>' + (isBlocked ? '<span class=\"badge badge-danger\">Bloccato</span>' : '<span class=\"badge badge-success\">Attivo</span>') + '</td>';\n        html += '<td style=\"color:var(--text-muted); font-size:0.8rem;\">' + (u.ultimoAcceso || '-') + '</td>';\n        html += '<td style=\"text-align:right; white-space:nowrap;\">';\n        if (u.role !== 'admin') {\n          html += '<button class=\"btn btn-sm\" onclick=\"openResetPasswordModal(\\'' + u.email + '\\')\" title=\"Reset Password\" style=\"margin-right:6px;\">🔑</button>';\n          html += '<button class=\"btn btn-sm ' + (isBlocked ? 'btn-success' : 'btn-danger') + '\" onclick=\"toggleBloccaUser(\\'' + u.email + '\\')\" style=\"margin-right:6px;\">' + (isBlocked ? 'Sblocca' : 'Blocca') + '</button>';\n          html += '<button class=\"btn btn-sm btn-danger\" onclick=\"eliminaUser(\\'' + u.email + '\\')\">🗑️</button>';\n        }\n        html += '</td>';\n        html += '</tr>';\n      });\n      tbody.innerHTML = html;\n    }\n\n    function openModalAddUser() {\n      document.getElementById('user-email').value = \"\";\n      document.getElementById('user-nombre').value = \"\";\n      generaPasswordCasuale('user-password');\n      openModal('modal-user');\n    }\n\n    async function handleCreateUser(e) {\n      e.preventDefault();\n      const userData = {\n        email: document.getElementById('user-email').value.trim(),\n        nombre: document.getElementById('user-nombre').value.trim(),\n        encargadoKey: document.getElementById('user-enckey').value,\n        password: document.getElementById('user-password').value,\n        role: \"colaborador\"\n      };\n\n      try {\n        const res = await apiCall(\"crearUsuario\", { userData: userData });\n        if (res && res.success) {\n          closeModal('modal-user');\n          toast(\"Account collaboratore creato con successo!\", \"success\");\n          loadAdminUtenti();\n        } else {\n          toast(res.error || \"Errore nella creazione\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    function openResetPasswordModal(email) {\n      document.getElementById('reset-user-email').value = email;\n      document.getElementById('reset-user-email-label').innerText = email;\n      generaPasswordCasuale('reset-new-password');\n      openModal('modal-reset-pwd');\n    }\n\n    async function handleResetPasswordSubmit(e) {\n      e.preventDefault();\n      const email = document.getElementById('reset-user-email').value;\n      const newPwd = document.getElementById('reset-new-password').value;\n      try {\n        const res = await apiCall(\"resetPasswordUsuario\", { id: email, newPassword: newPwd });\n        if (res && res.success) {\n          closeModal('modal-reset-pwd');\n          toast(\"Password reimpostata con successo!\", \"success\");\n        } else {\n          toast(res.error || \"Errore reset\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    async function toggleBloccaUser(email) {\n      try {\n        const res = await apiCall(\"toggleBloquearUsuario\", { id: email });\n        if (res && res.success) {\n          toast(res.bloqueado ? \"Collaboratore bloccato\" : \"Collaboratore sbloccato\", \"info\");\n          loadAdminUtenti();\n        } else {\n          toast(res.error || \"Errore\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    async function eliminaUser(email) {\n      if (!confirm(\"Sei sicuro di voler eliminare questo collaboratore?\\n\\nLo storico dei pagamenti sarà conservato nel foglio Contabilità e le pulizie future resteranno da riassegnare.\")) return;\n      try {\n        const res = await apiCall(\"eliminarUsuario\", { id: email });\n        if (res && res.success) {\n          toast(\"Collaboratore eliminato (storico contabilità preservato)\", \"info\");\n          loadAdminUtenti();\n        } else {\n          toast(res.error || \"Errore eliminazione\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    // --- Contabilità Admin ---\n    async function loadAdminContabilita() {\n      const mes = document.getElementById('admin-cont-filter-mes').value;\n      const anno = document.getElementById('admin-cont-filter-anno').value;\n      const enc = document.getElementById('admin-cont-filter-enc').value;\n      const tbody = document.getElementById('tbody-admin-contabilita');\n\n      tbody.innerHTML = '<tr><td colspan=\"10\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento contabilità...</td></tr>';\n\n      try {\n        const res = await apiCall(\"getContabilidad\", { mes: mes, anno: anno, encargado: enc });\n        if (res && res.success) {\n          renderAdminContabilita(res.data || []);\n        } else {\n          tbody.innerHTML = '<tr><td colspan=\"10\" style=\"text-align:center; padding:30px; color:var(--danger);\">' + (res.error || \"Errore\") + '</td></tr>';\n        }\n      } catch (err) {\n        tbody.innerHTML = '<tr><td colspan=\"10\" style=\"text-align:center; padding:30px; color:var(--danger);\">Errore di rete</td></tr>';\n      }\n    }\n\n    function renderAdminContabilita(records) {\n      let tot = 0, pag = 0, att = 0, pul = 0;\n      records.forEach(c => {\n        const t = parseFloat(c.total) || 0;\n        tot += t;\n        pul += parseInt(c.nLimpiezas) || 0;\n        if (c.estado === 'pagato' || c.stato === 'pagato') pag += t;\n        else att += t;\n      });\n\n      document.getElementById('admin-cont-totale').innerText = \"€ \" + tot.toFixed(2);\n      document.getElementById('admin-cont-pagato').innerText = \"€ \" + pag.toFixed(2);\n      document.getElementById('admin-cont-attesa').innerText = \"€ \" + att.toFixed(2);\n      document.getElementById('admin-cont-pulizie').innerText = pul;\n\n      const tbody = document.getElementById('tbody-admin-contabilita');\n      if (!records || records.length === 0) {\n        tbody.innerHTML = '<tr><td colspan=\"10\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Nessun record contabile trovato</td></tr>';\n        return;\n      }\n\n      let html = '';\n      records.forEach(c => {\n        const isPag = (c.estado === 'pagato' || c.stato === 'pagato');\n        html += '<tr>';\n        html += '<td><strong>' + (c.nombreMes || c.mes) + ' ' + c.anno + '</strong></td>';\n        html += '<td><span class=\"badge badge-enc\">' + (c.encargado || '-') + '</span></td>';\n        html += '<td>' + (c.nLimpiezas || 0) + '</td>';\n        html += '<td>€ ' + (c.precioPorLimpieza || 0).toFixed(2) + '</td>';\n        html += '<td>€ ' + (c.gastos || 0).toFixed(2) + '</td>';\n        html += '<td><strong style=\"color:var(--text); font-size:0.95rem;\">€ ' + (c.total || 0).toFixed(2) + '</strong></td>';\n        html += '<td>' + (isPag ? '<span class=\"badge badge-success\">Pagato</span>' : '<span class=\"badge badge-warning\">In attesa</span>') + '</td>';\n        html += '<td>' + (c.fechaPago || '-') + '</td>';\n        html += '<td style=\"color:var(--text-muted); font-size:0.8rem;\">' + (c.notas || '-') + '</td>';\n        html += '<td style=\"text-align:right;\">';\n        html += '<button class=\"btn btn-sm\" onclick=\"openModalStatoCont(\\'' + c.id + '\\', \\'' + (c.estado||c.stato) + '\\', \\'' + (c.fechaPago||'') + '\\', \\'' + (c.notas||'') + '\\')\">✏️ Stato</button>';\n        html += '</td>';\n        html += '</tr>';\n      });\n      tbody.innerHTML = html;\n    }\n\n    function openModalCalcolaComp() {\n      document.getElementById('calc-result-box').style.display = 'none';\n      document.getElementById('btn-salva-calc-comp').style.display = 'none';\n      openModal('modal-calc-comp');\n    }\n\n    let ULTIMO_CALCOLO_COMP = null;\n    async function anteprimaCalcolo() {\n      const enc = document.getElementById('calc-encargado').value;\n      const mes = document.getElementById('calc-mes').value;\n      const anno = document.getElementById('calc-anno').value;\n      const tarifa = parseFloat(document.getElementById('calc-tarifa').value) || 40;\n\n      try {\n        const res = await apiCall(\"calcolarCompensacion\", { encargado: enc, mes: mes, anno: anno, precioPorLimpieza: tarifa });\n        if (res && res.success) {\n          ULTIMO_CALCOLO_COMP = res;\n          document.getElementById('calc-count').innerText = res.nLimpiezas;\n          document.getElementById('calc-totale').innerText = \"€ \" + res.total.toFixed(2);\n          document.getElementById('calc-result-box').style.display = 'block';\n          document.getElementById('btn-salva-calc-comp').style.display = 'inline-flex';\n        } else {\n          toast(res.error || \"Errore nel calcolo\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    async function handleExecCalcolaComp(e) {\n      e.preventDefault();\n      if (!ULTIMO_CALCOLO_COMP) return;\n      const contData = {\n        encargado: ULTIMO_CALCOLO_COMP.encargado,\n        email: ULTIMO_CALCOLO_COMP.email || \"\",\n        mes: ULTIMO_CALCOLO_COMP.mes,\n        anno: ULTIMO_CALCOLO_COMP.anno,\n        nLimpiezas: ULTIMO_CALCOLO_COMP.nLimpiezas,\n        precioPorLimpieza: ULTIMO_CALCOLO_COMP.precioPorLimpieza,\n        gastos: 0,\n        estado: 'pendiente',\n        notas: \"Calcolato automaticamente da calendario\"\n      };\n\n      try {\n        const res = await apiCall(\"saveContabilidad\", { contData: contData });\n        if (res && res.success) {\n          closeModal('modal-calc-comp');\n          toast(\"Record di compenso salvato con successo!\", \"success\");\n          loadAdminContabilita();\n        } else {\n          toast(res.error || \"Errore nel salvataggio\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    function openModalStatoCont(id, estado, fecha, notas) {\n      document.getElementById('stato-cont-id').value = id;\n      document.getElementById('stato-cont-estado').value = (estado === 'pagato') ? 'pagato' : 'pendiente';\n      document.getElementById('stato-cont-fecha').value = fecha || \"\";\n      document.getElementById('stato-cont-notas').value = notas || \"\";\n      openModal('modal-stato-cont');\n    }\n\n    async function handleUpdateStatoCont(e) {\n      e.preventDefault();\n      const contId = document.getElementById('stato-cont-id').value;\n      const estado = document.getElementById('stato-cont-estado').value;\n      const fecha = document.getElementById('stato-cont-fecha').value;\n      const notas = document.getElementById('stato-cont-notas').value;\n\n      try {\n        const res = await apiCall(\"updateEstadoContabilidad\", { contId: contId, estado: estado, fechaPago: fecha, notas: notas });\n        if (res && res.success) {\n          closeModal('modal-stato-cont');\n          toast(\"Stato pagamento aggiornato\", \"success\");\n          loadAdminContabilita();\n        } else {\n          toast(res.error || \"Errore aggiornamento\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    // --- Configurazione App (Admin) ---\n    async function loadAdminConfig() {\n      await loadAppConfig();\n      document.getElementById('cfg-apartamentos').value = (APP_CONFIG.apartamentos || []).join('\\n');\n      document.getElementById('cfg-encargados').value = (APP_CONFIG.encargados || []).join('\\n');\n    }\n\n    async function salvaConfigurazioneApp() {\n      const apts = document.getElementById('cfg-apartamentos').value.split('\\n').map(s => s.trim()).filter(Boolean);\n      const encs = document.getElementById('cfg-encargados').value.split('\\n').map(s => s.trim()).filter(Boolean);\n      try {\n        const res = await apiCall(\"saveAppConfig\", { apartamentos: apts, encargados: encs });\n        if (res && res.success) {\n          toast(\"Configurazione salvata con successo\", \"success\");\n          await loadAppConfig();\n        } else {\n          toast(res.error || \"Errore salvataggio\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    // --- Collaboratore View Logic ---\n    async function loadCollabCalendario() {\n      const mes = document.getElementById('collab-filter-mes').value;\n      const anno = document.getElementById('collab-filter-anno').value;\n      const tbody = document.getElementById('tbody-collab-calendario');\n\n      tbody.innerHTML = '<tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento del tuo calendario...</td></tr>';\n\n      try {\n        const res = await apiCall(\"getCalendario\", { mes: mes, anno: anno });\n        if (res && res.success) {\n          renderCollabCalendario(res.rows || []);\n        } else {\n          tbody.innerHTML = '<tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--danger);\">' + (res.error || \"Errore\") + '</td></tr>';\n        }\n      } catch (err) {\n        tbody.innerHTML = '<tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--danger);\">Errore di rete</td></tr>';\n      }\n    }\n\n    function renderCollabCalendario(rows) {\n      const tbody = document.getElementById('tbody-collab-calendario');\n      document.getElementById('collab-count-badge').innerText = rows.length + \" Pulizie\";\n\n      if (!rows || rows.length === 0) {\n        tbody.innerHTML = '<tr><td colspan=\"7\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Nessuna pulizia assegnata a te per questo mese</td></tr>';\n        return;\n      }\n\n      let html = '';\n      rows.forEach(r => {\n        html += '<tr>';\n        html += '<td><span class=\"badge badge-apt\">' + (r.apartamento || '-') + '</span></td>';\n        html += '<td><strong>' + (r.fecha || '-') + '</strong></td>';\n        html += '<td>' + (r.horario || '-') + '</td>';\n        html += '<td>' + (r.preparar || '-') + '</td>';\n        html += '<td>' + (r.entrada || '-') + '</td>';\n        html += '<td>' + (r.notas || '-') + '</td>';\n        html += '<td>';\n        html += '<button class=\"btn btn-sm\" onclick=\"openModalNotaEnc(\\'' + r.id + '\\', \\'' + (r.notaEncargado || '').replace(/'/g, \"\\\\'\") + '\\')\">' + (r.notaEncargado ? '💬 ' + r.notaEncargado : '➕ Aggiungi nota') + '</button>';\n        html += '</td>';\n        html += '</tr>';\n      });\n      tbody.innerHTML = html;\n    }\n\n    function openModalNotaEnc(id, currentNota) {\n      document.getElementById('nota-enc-id').value = id;\n      document.getElementById('nota-enc-text').value = currentNota || \"\";\n      openModal('modal-nota-enc');\n    }\n\n    async function handleSaveNotaEnc(e) {\n      e.preventDefault();\n      const id = document.getElementById('nota-enc-id').value;\n      const nota = document.getElementById('nota-enc-text').value.trim();\n      const mes = document.getElementById('collab-filter-mes').value;\n      const anno = document.getElementById('collab-filter-anno').value;\n\n      try {\n        const res = await apiCall(\"saveNotaEncargado\", { id: id, nota: nota, mes: mes, anno: anno });\n        if (res && res.success) {\n          closeModal('modal-nota-enc');\n          toast(\"Nota salvata con successo\", \"success\");\n          loadCollabCalendario();\n        } else {\n          toast(res.error || \"Errore salvataggio nota\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n\n    async function loadCollabCompensi() {\n      const tbody = document.getElementById('tbody-collab-compensi');\n      tbody.innerHTML = '<tr><td colspan=\"8\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Caricamento compensi...</td></tr>';\n\n      try {\n        const res = await apiCall(\"getContabilidad\");\n        if (res && res.success) {\n          renderCollabCompensi(res.data || []);\n        } else {\n          tbody.innerHTML = '<tr><td colspan=\"8\" style=\"text-align:center; padding:30px; color:var(--danger);\">' + (res.error || \"Errore\") + '</td></tr>';\n        }\n      } catch (err) {\n        tbody.innerHTML = '<tr><td colspan=\"8\" style=\"text-align:center; padding:30px; color:var(--danger);\">Errore di rete</td></tr>';\n      }\n    }\n\n    function renderCollabCompensi(records) {\n      let tot = 0, pag = 0, att = 0, pul = 0;\n      records.forEach(c => {\n        const t = parseFloat(c.total) || 0;\n        tot += t;\n        pul += parseInt(c.nLimpiezas) || 0;\n        if (c.estado === 'pagato' || c.stato === 'pagato') pag += t;\n        else att += t;\n      });\n\n      document.getElementById('collab-stat-totale').innerText = \"€ \" + tot.toFixed(2);\n      document.getElementById('collab-stat-pagato').innerText = \"€ \" + pag.toFixed(2);\n      document.getElementById('collab-stat-attesa').innerText = \"€ \" + att.toFixed(2);\n      document.getElementById('collab-stat-pulizie').innerText = pul;\n\n      const tbody = document.getElementById('tbody-collab-compensi');\n      if (!records || records.length === 0) {\n        tbody.innerHTML = '<tr><td colspan=\"8\" style=\"text-align:center; padding:30px; color:var(--text-muted);\">Nessun compenso registrato per te finora</td></tr>';\n        return;\n      }\n\n      let html = '';\n      records.forEach(c => {\n        const isPag = (c.estado === 'pagato' || c.stato === 'pagato');\n        html += '<tr>';\n        html += '<td><strong>' + (c.nombreMes || c.mes) + ' ' + c.anno + '</strong></td>';\n        html += '<td>' + (c.nLimpiezas || 0) + '</td>';\n        html += '<td>€ ' + (c.precioPorLimpieza || 0).toFixed(2) + '</td>';\n        html += '<td>€ ' + (c.gastos || 0).toFixed(2) + '</td>';\n        html += '<td><strong style=\"color:var(--text); font-size:0.95rem;\">€ ' + (c.total || 0).toFixed(2) + '</strong></td>';\n        html += '<td>' + (isPag ? '<span class=\"badge badge-success\">Pagato</span>' : '<span class=\"badge badge-warning\">In attesa</span>') + '</td>';\n        html += '<td>' + (c.fechaPago || '-') + '</td>';\n        html += '<td style=\"color:var(--text-muted); font-size:0.8rem;\">' + (c.notas || '-') + '</td>';\n        html += '</tr>';\n      });\n      tbody.innerHTML = html;\n    }\n\n    // --- Cambio Password ---\n    async function handleCambiaPassword(e) {\n      e.preventDefault();\n      const isAdmin = SESSION.role === 'admin';\n      const oldPwd = document.getElementById(isAdmin ? 'admin-old-pwd' : 'collab-old-pwd').value;\n      const newPwd = document.getElementById(isAdmin ? 'admin-new-pwd' : 'collab-new-pwd').value;\n      const confPwd = document.getElementById(isAdmin ? 'admin-conf-pwd' : 'collab-conf-pwd').value;\n\n      if (newPwd !== confPwd) {\n        toast(\"Le nuove password non coincidono\", \"error\");\n        return;\n      }\n\n      try {\n        const res = await apiCall(\"cambiarPassword\", { oldPassword: oldPwd, newPassword: newPwd });\n        if (res && res.success) {\n          toast(\"Password cambiata con successo!\", \"success\");\n          e.target.reset();\n        } else {\n          toast(res.error || \"Errore cambio password\", \"error\");\n        }\n      } catch (err) {\n        toast(\"Errore di rete\", \"error\");\n      }\n    }\n  </script>\n</body>\n</html>\n";
}

function getSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId  = props.getProperty('SPREADSHEET_ID');

  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (e) {
      props.deleteProperty('SPREADSHEET_ID');
    }
  }

  // Cerca per nome
  var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    var ss = SpreadsheetApp.open(files.next());
    props.setProperty('SPREADSHEET_ID', ss.getId());
    return ss;
  }

  // Crea nuovo
  var newSs = SpreadsheetApp.create(SPREADSHEET_NAME);
  props.setProperty('SPREADSHEET_ID', newSs.getId());
  return newSs;
}

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    var headers = [
      "ID", "APARTAMENTO", "FECHA LIMPIEZA", "HORARIO",
      "PREPARAR PARA", "ENTRADA PROXIMOS", "NOTAS ADMIN",
      "DIAS LIBRES", "PLAT.", "ENCARGADO/A",
      "NOTA ENCARGADO", "TIMESTAMP"
    ];
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground("#2c3e50").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 200);
    sheet.setColumnWidth(11, 220);
  }
  return sheet;
}

function getOrCreateUsersSheet(ss) {
  var sheet = ss.getSheetByName("UTENTI");
  if (!sheet) {
    sheet = ss.insertSheet("UTENTI");
    var headers = [
      "EMAIL", "PASSWORD_HASH", "NOMBRE", "ENCARGADO_KEY",
      "ROLE", "ACTIVO", "BLOQUEADO", "CREADO", "ULTIMO_ACCESO"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#1a237e").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 150);
    // Crea admin di default
    var adminHash = _hashPassword("Admin2024!");
    sheet.appendRow([
      "admin@pulizie.app", adminHash, "Amministratore", "ADMIN",
      "admin", true, false,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      ""
    ]);
  }
  return sheet;
}

function getOrCreateContabilidadSheet(ss) {
  var sheet = ss.getSheetByName("CONTABILIDAD");
  if (!sheet) {
    sheet = ss.insertSheet("CONTABILIDAD");
    var headers = [
      "ID", "ENCARGADO", "EMAIL", "MES", "ANNO",
      "N_LIMPIEZAS", "PRECIO_LIMP", "GASTOS", "TOTAL",
      "ESTADO", "FECHA_PAGO", "NOTAS", "CREADO"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#1b5e20").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(11, 120);
    sheet.setColumnWidth(12, 200);
  }
  return sheet;
}

function getSheetName(mes, anno) {
  return MESES_ES[mes - 1] + " " + anno;
}

// ══ HASHING PASSWORD (semplice, compatibile GAS) ════════════
// NB: in produzione si raccomanda un hash più robusto tramite Utilities.computeDigest
function _hashPassword(plain) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    plain + "PULIZIE_SALT_2024",
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

function _generateRandomPassword(length) {
  length = length || 10;
  var chars = "ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789!@#";
  var result = "";
  for (var i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

// ══ AUTENTICAZIONE ═══════════════════════════════════════════

/**
 * Login: verifica email + password, restituisce token di sessione.
 * Il token è salvato in ScriptProperties con TTL di 8 ore.
 */
function login(email, password) {
  try {
    if (!email || !password) return { success: false, error: "Credenziali mancanti" };
    var emailLower = email.toString().toLowerCase().trim();
    var hash = _hashPassword(password);

    var ss = getSpreadsheet();
    var sheet = getOrCreateUsersSheet(ss);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][COL_U.EMAIL] || "").toString().toLowerCase().trim();
      if (rowEmail !== emailLower) continue;

      // Controlla se l'utente è stato eliminato (riga vuota o marcata)
      if (data[i][COL_U.ACTIVO] === false || data[i][COL_U.ACTIVO] === "FALSE" || data[i][COL_U.ACTIVO] === "") {
        return { success: false, error: "Account non attivo. Contatta l'amministratore." };
      }
      if (data[i][COL_U.BLOQUEADO] === true || data[i][COL_U.BLOQUEADO] === "TRUE") {
        return { success: false, error: "Account bloccato. Contatta l'amministratore." };
      }

      var storedHash = (data[i][COL_U.PASSWORD_HASH] || "").toString();
      if (storedHash !== hash) {
        return { success: false, error: "Password non corretta." };
      }

      // Genera token sessione
      var token = "SES_" + new Date().getTime() + "_" + Math.floor(Math.random() * 999999);
      var expiry = new Date().getTime() + (8 * 60 * 60 * 1000); // 8 ore

      var sessionData = {
        token: token,
        email: rowEmail,
        nombre: data[i][COL_U.NOMBRE] || "",
        encargadoKey: data[i][COL_U.ENCARGADO_KEY] || "",
        role: data[i][COL_U.ROLE] || "colaborador",
        expiry: expiry
      };

      PropertiesService.getScriptProperties().setProperty(token, JSON.stringify(sessionData));

      // Aggiorna ultimo accesso
      sheet.getRange(i + 1, COL_U.ULTIMO_ACCESO + 1).setValue(
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      );

      return {
        success: true,
        token: token,
        role: sessionData.role,
        nombre: sessionData.nombre,
        email: rowEmail,
        encargadoKey: sessionData.encargadoKey
      };
    }

    return { success: false, error: "Email non trovata." };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Verifica token di sessione. Restituisce i dati utente o errore.
 */
function verificarSesion(token) {
  try {
    if (!token) return { valid: false, error: "Token mancante" };
    var raw = PropertiesService.getScriptProperties().getProperty(token);
    if (!raw) return { valid: false, error: "Sessione non trovata" };
    var sess = JSON.parse(raw);
    if (new Date().getTime() > sess.expiry) {
      PropertiesService.getScriptProperties().deleteProperty(token);
      return { valid: false, error: "Sessione scaduta" };
    }
    return {
      valid: true,
      email: sess.email,
      nombre: sess.nombre,
      role: sess.role,
      encargadoKey: sess.encargadoKey,
      token: token
    };
  } catch (e) {
    return { valid: false, error: e.message };
  }
}

/**
 * Logout: invalida il token di sessione.
 */
function logout(token) {
  try {
    if (token) PropertiesService.getScriptProperties().deleteProperty(token);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Cambio password collaboratore (richiede vecchia password).
 */
function cambiarPassword(token, oldPassword, newPassword) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid) return { success: false, error: "Sessione non valida" };
    if (!newPassword || newPassword.length < 6) {
      return { success: false, error: "La nuova password deve avere almeno 6 caratteri" };
    }

    var ss = getSpreadsheet();
    var sheet = getOrCreateUsersSheet(ss);
    var data = sheet.getDataRange().getValues();
    var emailLower = sess.email.toLowerCase();
    var oldHash = _hashPassword(oldPassword);

    for (var i = 1; i < data.length; i++) {
      var rowEmail = (data[i][COL_U.EMAIL] || "").toString().toLowerCase();
      if (rowEmail !== emailLower) continue;

      // Admin può cambiare senza vecchia password; collaboratore deve fornirla
      if (sess.role !== "admin") {
        var storedHash = (data[i][COL_U.PASSWORD_HASH] || "").toString();
        if (storedHash !== oldHash) {
          return { success: false, error: "La vecchia password non è corretta" };
        }
      }

      sheet.getRange(i + 1, COL_U.PASSWORD_HASH + 1).setValue(_hashPassword(newPassword));
      return { success: true };
    }
    return { success: false, error: "Utente non trovato" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ══ GESTIONE UTENTI (ADMIN) ══════════════════════════════════

/**
 * Restituisce la lista di tutti i collaboratori (solo admin).
 */
function getUsuarios(token) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var ss = getSpreadsheet();
    var sheet = getOrCreateUsersSheet(ss);
    var data = sheet.getDataRange().getValues();
    var users = [];

    for (var i = 1; i < data.length; i++) {
      var email = (data[i][COL_U.EMAIL] || "").toString();
      if (!email) continue;
      users.push({
        email:        email,
        nombre:       (data[i][COL_U.NOMBRE] || "").toString(),
        encargadoKey: (data[i][COL_U.ENCARGADO_KEY] || "").toString(),
        role:         (data[i][COL_U.ROLE] || "colaborador").toString(),
        activo:       data[i][COL_U.ACTIVO] === true || data[i][COL_U.ACTIVO] === "TRUE",
        bloqueado:    data[i][COL_U.BLOQUEADO] === true || data[i][COL_U.BLOQUEADO] === "TRUE",
        creado:       (data[i][COL_U.CREADO] || "").toString(),
        ultimoAcceso: (data[i][COL_U.ULTIMO_ACCESO] || "").toString()
      });
    }

    return { success: true, users: users };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Crea un nuovo collaboratore (solo admin).
 * Se password è vuota, ne genera una random.
 */
function crearUsuario(token, userData) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var email = (userData.email || "").toString().toLowerCase().trim();
    if (!email || email.indexOf("@") === -1) return { success: false, error: "Email non valida" };

    var password = userData.password || _generateRandomPassword(10);
    var nombre   = (userData.nombre || "").toString().trim();
    var encKey   = (userData.encargadoKey || "").toString().toUpperCase().trim();
    var role     = (userData.role || "colaborador").toString();

    var ss = getSpreadsheet();
    var sheet = getOrCreateUsersSheet(ss);
    var data = sheet.getDataRange().getValues();

    // Controlla duplicati
    for (var i = 1; i < data.length; i++) {
      if ((data[i][COL_U.EMAIL] || "").toString().toLowerCase() === email) {
        return { success: false, error: "Email già esistente" };
      }
    }

    sheet.appendRow([
      email,
      _hashPassword(password),
      nombre,
      encKey,
      role,
      true,
      false,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
      ""
    ]);

    return { success: true, password: password, email: email };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Aggiorna i dati di un collaboratore (solo admin).
 * Non aggiorna la password (usa resetPasswordUsuario per quello).
 */
function actualizarUsuario(token, email, updates) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var emailLower = email.toString().toLowerCase().trim();
    var ss = getSpreadsheet();
    var sheet = getOrCreateUsersSheet(ss);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if ((data[i][COL_U.EMAIL] || "").toString().toLowerCase() !== emailLower) continue;

      if (updates.nombre !== undefined)       sheet.getRange(i+1, COL_U.NOMBRE+1).setValue(updates.nombre);
      if (updates.encargadoKey !== undefined) sheet.getRange(i+1, COL_U.ENCARGADO_KEY+1).setValue(updates.encargadoKey.toUpperCase());
      if (updates.role !== undefined)         sheet.getRange(i+1, COL_U.ROLE+1).setValue(updates.role);

      return { success: true };
    }
    return { success: false, error: "Utente non trovato" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Resetta la password di un collaboratore (solo admin).
 * Se newPassword è vuota, ne genera una random.
 */
function resetPasswordUsuario(token, email, newPassword) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var emailLower = email.toString().toLowerCase().trim();
    var pwd = newPassword || _generateRandomPassword(10);

    var ss = getSpreadsheet();
    var sheet = getOrCreateUsersSheet(ss);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if ((data[i][COL_U.EMAIL] || "").toString().toLowerCase() !== emailLower) continue;
      sheet.getRange(i+1, COL_U.PASSWORD_HASH+1).setValue(_hashPassword(pwd));
      return { success: true, newPassword: pwd };
    }
    return { success: false, error: "Utente non trovato" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Blocca o sblocca un collaboratore (solo admin).
 * Un collaboratore bloccato non può accedere ma i suoi dati sono conservati.
 */
function toggleBloquearUsuario(token, email, bloqueado) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var emailLower = email.toString().toLowerCase().trim();
    var ss = getSpreadsheet();
    var sheet = getOrCreateUsersSheet(ss);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if ((data[i][COL_U.EMAIL] || "").toString().toLowerCase() !== emailLower) continue;
      sheet.getRange(i+1, COL_U.BLOQUEADO+1).setValue(bloqueado === true);
      return { success: true, bloqueado: bloqueado === true };
    }
    return { success: false, error: "Utente non trovato" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Cancella un collaboratore (solo admin).
 * Le pulizie assegnate rimangono nel calendario (da riassegnare).
 * La contabilità storica è conservata.
 * NOTA: l'utente viene marcato come eliminato (ACTIVO=false) per conservare lo storico.
 */
function eliminarUsuario(token, email) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var emailLower = email.toString().toLowerCase().trim();
    // Non eliminare l'admin stesso
    if (emailLower === sess.email.toLowerCase()) {
      return { success: false, error: "Non puoi eliminare il tuo stesso account" };
    }

    var ss = getSpreadsheet();
    var sheet = getOrCreateUsersSheet(ss);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if ((data[i][COL_U.EMAIL] || "").toString().toLowerCase() !== emailLower) continue;
      // Marca come eliminato (non cancella la riga per preservare lo storico)
      sheet.getRange(i+1, COL_U.ACTIVO+1).setValue(false);
      sheet.getRange(i+1, COL_U.BLOQUEADO+1).setValue(true);
      // Invalida eventuali sessioni attive
      _invalidarSesionesUsuario(emailLower);
      return { success: true };
    }
    return { success: false, error: "Utente non trovato" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function _invalidarSesionesUsuario(emailLower) {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    Object.keys(all).forEach(function(k) {
      if (k.indexOf("SES_") === 0) {
        try {
          var s = JSON.parse(all[k]);
          if (s.email && s.email.toLowerCase() === emailLower) {
            props.deleteProperty(k);
          }
        } catch(e2) {}
      }
    });
  } catch(e) {}
}

// ══ CALENDARIO ═══════════════════════════════════════════════

/**
 * Legge il calendario per admin (tutti) o collaboratore (solo proprio).
 */
function getCalendario(token, mes, anno, userFilter, aptFilter) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid) return { success: false, error: "Sessione non valida" };

    // Collaboratori vedono solo le proprie limpiezas
    if (sess.role !== "admin") {
      userFilter = sess.encargadoKey;
      aptFilter  = "";
    }

    var ss        = getSpreadsheet();
    var sheetName = getSheetName(mes, anno);
    var sheet     = ss.getSheetByName(sheetName);
    if (!sheet) return { success: true, data: [], sheetName: sheetName };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [], sheetName: sheetName };

    var headers = data[0];
    var rows    = [];
    var ufLow   = (userFilter || "").toString().toLowerCase().trim();
    var afLow   = (aptFilter  || "").toString().toLowerCase().trim();

    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        row[headers[j]] = val;
      }

      var rowEnc = (row["ENCARGADO/A"] || "").toString().toLowerCase();
      var rowApt = (row["APARTAMENTO"] || "").toString().toLowerCase();

      if (ufLow) {
        if (ufLow === "daniela") {
          if (rowEnc !== "daniela" && rowEnc !== "levys") continue;
        } else {
          if (rowEnc.indexOf(ufLow) === -1) continue;
        }
      }

      if (afLow) {
        if (rowApt.indexOf(afLow) === -1) continue;
      }

      rows.push(row);
    }

    return { success: true, data: rows, sheetName: sheetName };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Salva/aggiorna una riga del calendario (solo admin).
 */
function saveRiga(token, rowData, mes, anno) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var ss = getSpreadsheet();
    var target = _parseMesAnno(rowData["FECHA LIMPIEZA"], mes, anno);
    var targetSheetName = getSheetName(target.mes, target.anno);
    var targetSheet     = getOrCreateSheet(ss, targetSheetName);
    var notaEncargado   = "";

    if (rowData.ID) {
      var oldSheetName = getSheetName(mes, anno);
      var oldSheet     = ss.getSheetByName(oldSheetName);
      if (oldSheet && oldSheetName !== targetSheetName) {
        var oldData = oldSheet.getDataRange().getValues();
        for (var k = 1; k < oldData.length; k++) {
          if (String(oldData[k][0]) === String(rowData.ID)) {
            notaEncargado = oldData[k][COL.NOTA_ENCARGADO] || "";
            oldSheet.deleteRow(k + 1);
            break;
          }
        }
      }

      var targetData = targetSheet.getDataRange().getValues();
      for (var i = 1; i < targetData.length; i++) {
        if (String(targetData[i][0]) === String(rowData.ID)) {
          if (!notaEncargado) notaEncargado = targetData[i][COL.NOTA_ENCARGADO] || "";
          targetSheet.getRange(i + 1, 1, 1, NUM_COLS).setValues([[
            rowData.ID,
            rowData["APARTAMENTO"]      || "",
            rowData["FECHA LIMPIEZA"]   || "",
            rowData["HORARIO"]          || "",
            rowData["PREPARAR PARA"]    || "",
            rowData["ENTRADA PROXIMOS"] || "",
            rowData["NOTAS ADMIN"]      || "",
            rowData["DIAS LIBRES"]      || 0,
            rowData["PLAT."]            || "",
            rowData["ENCARGADO/A"]      || "",
            notaEncargado,
            new Date()
          ]]);
          _aplicarColorFila(targetSheet, i + 1, rowData["ENCARGADO/A"]);
          _ordenarPorFecha(targetSheet);
          return { success: true, action: "updated", id: rowData.ID, mes: target.mes, anno: target.anno };
        }
      }
    }

    var newId = new Date().getTime().toString();
    targetSheet.appendRow([
      rowData.ID || newId,
      rowData["APARTAMENTO"]      || "",
      rowData["FECHA LIMPIEZA"]   || "",
      rowData["HORARIO"]          || "",
      rowData["PREPARAR PARA"]    || "",
      rowData["ENTRADA PROXIMOS"] || "",
      rowData["NOTAS ADMIN"]      || "",
      rowData["DIAS LIBRES"]      || 0,
      rowData["PLAT."]            || "",
      rowData["ENCARGADO/A"]      || "",
      notaEncargado,
      new Date()
    ]);
    var newRowIndex = targetSheet.getLastRow();
    _aplicarColorFila(targetSheet, newRowIndex, rowData["ENCARGADO/A"]);
    _ordenarPorFecha(targetSheet);
    return { success: true, action: "created", id: rowData.ID || newId, mes: target.mes, anno: target.anno };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Elimina una riga del calendario (solo admin).
 */
function deleteRiga(token, id, mes, anno) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var ss        = getSpreadsheet();
    var sheetName = getSheetName(mes, anno);
    var sheet     = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: "Foglio non trovato" };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: "Riga non trovata" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Salva nota del collaboratore (accessibile ai collaboratori).
 */
function saveNotaEncargado(token, id, mes, anno, nota) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid) return { success: false, error: "Sessione non valida" };

    var ss        = getSpreadsheet();
    var sheetName = getSheetName(mes, anno);
    var sheet     = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: "Foglio non trovato" };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        // Collaboratori possono solo modificare le note delle proprie righe
        if (sess.role !== "admin") {
          var rowEnc = (data[i][COL.ENCARGADO] || "").toString().toUpperCase();
          var myKey  = (sess.encargadoKey || "").toString().toUpperCase();
          if (rowEnc !== myKey && !(myKey === "DANIELA" && rowEnc === "LEVYS")) {
            return { success: false, error: "Non autorizzato su questa riga" };
          }
        }
        sheet.getRange(i + 1, COL.NOTA_ENCARGADO + 1).setValue(nota || "");
        return { success: true };
      }
    }
    return { success: false, error: "Riga non trovata" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ══ CONTABILITÀ ══════════════════════════════════════════════

/**
 * Salva un record di contabilità (solo admin).
 */
function saveContabilidad(token, contData) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var ss    = getSpreadsheet();
    var sheet = getOrCreateContabilidadSheet(ss);

    var id = "CONT_" + new Date().getTime() + "_" + Math.floor(Math.random() * 9999);
    var nLimp   = parseInt(contData.nLimpiezas)    || 0;
    var precio  = parseFloat(contData.precioPorLimpieza) || 40;
    var gastos  = parseFloat(contData.gastos)       || 0;
    var total   = (nLimp * precio) + gastos;

    sheet.appendRow([
      id,
      (contData.encargado || "").toString(),
      (contData.email     || "").toString(),
      parseInt(contData.mes)  || 0,
      parseInt(contData.anno) || 0,
      nLimp,
      precio,
      gastos,
      total,
      (contData.estado    || "pendiente").toString(),
      (contData.fechaPago || "").toString(),
      (contData.notas     || "").toString(),
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
    ]);

    return { success: true, id: id, total: total };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Aggiorna lo stato di un pagamento (solo admin).
 */
function updateEstadoContabilidad(token, contId, estado, fechaPago, notas) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var ss    = getSpreadsheet();
    var sheet = getOrCreateContabilidadSheet(ss);
    var data  = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][COL_C.ID]) === String(contId)) {
        sheet.getRange(i+1, COL_C.ESTADO+1).setValue(estado || "pendiente");
        sheet.getRange(i+1, COL_C.FECHA_PAGO+1).setValue(fechaPago || "");
        sheet.getRange(i+1, COL_C.NOTAS+1).setValue(notas || "");
        return { success: true };
      }
    }
    return { success: false, error: "Record non trovato" };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Legge la contabilità:
 * - Admin: tutti i record (con filtri opzionali)
 * - Collaboratore: solo i propri record
 */
function getContabilidad(token, filtros) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid) return { success: false, error: "Sessione non valida" };

    var ss    = getSpreadsheet();
    var sheet = getOrCreateContabilidadSheet(ss);
    var data  = sheet.getDataRange().getValues();
    var rows  = [];

    var fEnc   = filtros && filtros.encargado ? filtros.encargado.toString().toUpperCase() : "";
    var fEmail = filtros && filtros.email     ? filtros.email.toString().toLowerCase()     : "";
    var fMes   = filtros && filtros.mes       ? parseInt(filtros.mes)                       : 0;
    var fAnno  = filtros && filtros.anno      ? parseInt(filtros.anno)                      : 0;

    for (var i = 1; i < data.length; i++) {
      var rowEnc   = (data[i][COL_C.ENCARGADO] || "").toString().toUpperCase();
      var rowEmail = (data[i][COL_C.EMAIL]      || "").toString().toLowerCase();
      var rowMes   = parseInt(data[i][COL_C.MES])  || 0;
      var rowAnno  = parseInt(data[i][COL_C.ANNO]) || 0;

      // Collaboratori vedono solo i propri
      if (sess.role !== "admin") {
        var myKey  = (sess.encargadoKey || "").toUpperCase();
        var myEmail = (sess.email || "").toLowerCase();
        if (rowEmail !== myEmail && rowEnc !== myKey) continue;
      } else {
        // Admin con filtri
        if (fEnc   && rowEnc   !== fEnc)   continue;
        if (fEmail && rowEmail !== fEmail)  continue;
      }

      if (fMes  && rowMes  !== fMes)  continue;
      if (fAnno && rowAnno !== fAnno) continue;

      rows.push({
        id:           (data[i][COL_C.ID]          || "").toString(),
        encargado:    (data[i][COL_C.ENCARGADO]   || "").toString(),
        email:        (data[i][COL_C.EMAIL]        || "").toString(),
        mes:          rowMes,
        anno:         rowAnno,
        nombreMes:    rowMes > 0 ? MESES_ES[rowMes - 1] : "",
        nLimpiezas:   parseInt(data[i][COL_C.N_LIMPIEZAS]) || 0,
        precioPorLimpieza: parseFloat(data[i][COL_C.PRECIO_LIMP]) || 40,
        gastos:       parseFloat(data[i][COL_C.GASTOS]) || 0,
        total:        parseFloat(data[i][COL_C.TOTAL])  || 0,
        estado:       (data[i][COL_C.ESTADO]      || "pendiente").toString(),
        stato:        (data[i][COL_C.ESTADO]      || "pendiente").toString(), // alias italiano
        fechaPago:    (data[i][COL_C.FECHA_PAGO]  || "").toString(),
        notas:        (data[i][COL_C.NOTAS]        || "").toString(),
        creado:       (data[i][COL_C.CREADO]       || "").toString()
      });
    }

    return { success: true, data: rows };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Calcola il compenso di un collaboratore per un dato mese
 * leggendo direttamente il calendario (usato dall'admin per creare record contabilità).
 */
function calcolarCompensacion(token, encargado, mes, anno, precioPorLimpieza) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var result = getCalendario(token, mes, anno, encargado, "");
    if (!result.success) return result;

    var precio  = parseFloat(precioPorLimpieza) || 40;
    var limp    = result.data;
    var subtotal = limp.length * precio;

    // Desglose per appartamento
    var mapaApt = {};
    limp.forEach(function(r) {
      var apt = (r["APARTAMENTO"] || "—").toString();
      if (!mapaApt[apt]) mapaApt[apt] = { apartamento: apt, nLimpiezas: 0, subtotal: 0 };
      mapaApt[apt].nLimpiezas++;
      mapaApt[apt].subtotal += precio;
    });

    return {
      success: true,
      encargado: encargado,
      mes: mes,
      anno: anno,
      nombreMes: MESES_ES[mes - 1],
      nLimpiezas: limp.length,
      precioPorLimpieza: precio,
      subtotal: subtotal,
      limpiezas: limp,
      desglosePorApartamento: Object.values(mapaApt)
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ══ CONFIGURAZIONE APP ═══════════════════════════════════════
function getAppConfig(token) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid) return { success: false, error: "Sessione non valida" };

    var props  = PropertiesService.getScriptProperties();
    var encStr = props.getProperty('CONFIG_ENCARGADOS');
    var aptStr = props.getProperty('CONFIG_APARTAMENTOS');

    var encargados  = encStr ? JSON.parse(encStr) : DEFAULT_ENCARGADOS;
    var apartamentos = aptStr ? JSON.parse(aptStr) : DEFAULT_APARTAMENTOS;

    return { success: true, encargados: encargados, apartamentos: apartamentos };
  } catch (e) {
    return { success: true, encargados: DEFAULT_ENCARGADOS, apartamentos: DEFAULT_APARTAMENTOS };
  }
}

function saveAppConfig(token, newEncargados, newApartamentos) {
  try {
    var sess = verificarSesion(token);
    if (!sess.valid || sess.role !== "admin") return { success: false, error: "Non autorizzato" };

    var props = PropertiesService.getScriptProperties();
    if (newEncargados)  props.setProperty('CONFIG_ENCARGADOS',   JSON.stringify(newEncargados));
    if (newApartamentos) props.setProperty('CONFIG_APARTAMENTOS', JSON.stringify(newApartamentos));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ══ UTILITÀ ══════════════════════════════════════════════════
function getNombreMes(mes) { return MESES_ES[mes - 1]; }

function _parseMesAnno(fechaStr, fallbackMes, fallbackAnno) {
  if (!fechaStr) return { mes: fallbackMes, anno: fallbackAnno };
  var d = new Date(fechaStr);
  if (!isNaN(d.getTime())) {
    return { mes: d.getMonth() + 1, anno: d.getFullYear() };
  }
  return { mes: fallbackMes, anno: fallbackAnno };
}

function _aplicarColorFila(sheet, rowIndex, encargado) {
  var color = "#ffffff";
  var enc   = (encargado || "").toString().toUpperCase();
  if      (enc === "DANIELA" || enc === "LEVYS") color = "#FFF9C4";
  else if (enc === "LINA")                        color = "#E3F2FD";
  else if (enc === "ALEJANDRA")                   color = "#F3E5F5";
  else if (enc === "MARCELLO")                    color = "#E8F5E9";
  sheet.getRange(rowIndex, 1, 1, NUM_COLS).setBackground(color);
}

function _ordenarPorFecha(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, NUM_COLS)
           .sort({ column: COL.FECHA + 1, ascending: true });
    }
  } catch (e) {}
}

function getDatosPDF(token, mes, anno) {
  var result = getCalendario(token, mes, anno, "", "");
  if (!result.success) return result;
  return {
    success:  true,
    titulo:   "PROGRAMMAZIONE " + MESES_ES[mes - 1] + " " + anno,
    mes:      MESES_ES[mes - 1],
    anno:     anno,
    data:     result.data,
    generato: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
  };
}

// ══ INIZIALIZZAZIONE ════════════════════════════════════════
/**
 * Eseguire una volta sola per inizializzare il sistema.
 * Crea i fogli UTENTI e CONTABILIDAD se non esistono.
 */
function initializeSystem() {
  try {
    var ss = getSpreadsheet();
    getOrCreateUsersSheet(ss);
    getOrCreateContabilidadSheet(ss);
    Logger.log("Sistema inizializzato correttamente. Spreadsheet ID: " + ss.getId());
    return { success: true, spreadsheetId: ss.getId() };
  } catch(e) {
    Logger.log("Errore: " + e.message);
    return { success: false, error: e.message };
  }
}
