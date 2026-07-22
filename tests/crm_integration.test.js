// Tests F1 — integración CRM Campaign Manager en RappiMind.
// Valida invariantes de la fusión SIN runtime de Apps Script:
//  sintaxis de los .js backend, contrato frontend↔backend de las páginas Crm*,
//  router en Código.js, allowlist de clasp, y que no haya secretos ni doGet duplicado.
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');

let passed = 0, failed = 0;
function t(name, fn) {
  try { fn(); console.log('  ✓ ' + name); passed++; }
  catch (e) { console.log('  ✗ ' + name + ' — ' + e.message); failed++; }
}
function assert(cond, msg) { if (!cond) throw new Error(msg || 'assert'); }

// 1) Sintaxis de todos los backend .js (V8 los parsea igual que GAS)
['Crm.gs.js', 'Código.js', 'Admin.gs.js', 'Config.gs.js', 'Helpers.gs.js', 'Setup.gs.js', 'P0.gs.js'].forEach(f => {
  t('sintaxis OK: ' + f, () => {
    const r = spawnSync('node', ['--check', path.join(ROOT, f)], { encoding: 'utf8' });
    assert(r.status === 0, (r.stderr || '').split('\n')[0]);
  });
});

const crm = read('Crm.gs.js');
const codigo = read('Código.js');
const pages = { 'CrmForm.html': read('CrmForm.html'), 'CrmStatus.html': read('CrmStatus.html'), 'CrmAdmin.html': read('CrmAdmin.html') };

// 2) Un solo doGet en todo el proyecto (el de Código.js)
t('doGet único (Crm.gs.js no define doGet)', () => {
  assert(!/function\s+doGet\s*\(/.test(crm), 'Crm.gs.js define doGet');
  assert((codigo.match(/function\s+doGet\s*\(/g) || []).length === 1, 'Código.js debe tener exactamente 1 doGet');
});

// 3) Router: las 3 páginas servidas desde Código.js
t('router sirve CrmForm/CrmStatus/CrmAdmin', () => {
  ['CrmForm', 'CrmStatus', 'CrmAdmin'].forEach(p =>
    assert(codigo.indexOf("_crmServePage('" + p + "'") >= 0, 'falta ' + p + ' en el router'));
  assert(/function\s+_crmServePage\s*\(/.test(crm), 'falta _crmServePage en Crm.gs.js');
});

// 4) Sin secretos: token de Slack fuera del código
t('sin token Slack hardcodeado (xoxb-)', () => {
  ['Crm.gs.js', 'Código.js', 'Admin.gs.js', 'Config.gs.js', 'Helpers.gs.js', 'Setup.gs.js',
   'CrmForm.html', 'CrmStatus.html', 'CrmAdmin.html', 'WebApp.Html'].forEach(f =>
    assert(read(f).indexOf('xoxb-') === -1, 'token en ' + f));
});

// 5) Config por Script Properties (nada de IDs de prod del CRM hardcodeados)
t('config CRM por Script Properties', () => {
  ['CRM_SPREADSHEET_ID', 'CRM_SLACK_BOT_TOKEN', 'CRM_SLACK_CHANNEL', 'CRM_CSV_FOLDER_ID'].forEach(k =>
    assert(crm.indexOf("_crmProp('" + k + "'") >= 0, 'falta property ' + k));
  assert(crm.indexOf('1yzcRTWhdVlm9G') === -1, 'spreadsheet del CRM hardcodeado');
  assert(crm.indexOf('1pZ03_RgDTVtud') === -1, 'carpeta CSV hardcodeada');
});

// 6) Contrato frontend↔backend: toda función google.script.run de las páginas existe
const API = ['getScriptUrl', 'saveRequest', 'saveRappiCreditosRequest', 'getTicketStatus',
  'getTicketsByEmail', 'getAllTickets', 'getAllRappiCreditosTickets', 'updateStatus', 'updateGlobalOffer'];
t('contrato: API usada por las páginas existe en Crm.gs.js', () => {
  API.forEach(fn => assert(new RegExp('function\\s+' + fn + '\\s*\\(').test(crm), 'falta function ' + fn));
  Object.entries(pages).forEach(([name, html]) => {
    const used = new Set();
    (html.match(/\.\s*([A-Za-z_]\w*)\s*\(/g) || []).forEach(m => {
      const fn = m.replace(/[.\s(]/g, '');
      if (API.includes(fn)) used.add(fn);
    });
    used.forEach(fn => assert(new RegExp('function\\s+' + fn + '\\s*\\(').test(crm), name + ' usa ' + fn + ' inexistente'));
  });
});

// 7) Slugs migrados: ninguna página apunta a las rutas viejas del proyecto CRM
t('slugs viejos (page=form|home|status|admin) eliminados', () => {
  Object.entries(pages).forEach(([name, html]) => {
    assert(!/navTo\('(home|form|status|admin)'\)/.test(html), name + ': navTo viejo');
    assert(!/page=(home|form|status|admin)\b/.test(html), name + ': ?page= viejo');
  });
});

// 8) TC_DOC_URL cableado end-to-end en el backend
t('TC_DOC_URL en HEADERS/IDX/save/status/all', () => {
  assert(crm.indexOf("'TC_DOC_URL'") >= 0, 'falta en CRM_HEADERS');
  assert(/TC_DOC_URL:\s*55/.test(crm), 'falta CRM_IDX.TC_DOC_URL=55');
  assert(crm.indexOf('data.tcDocUrl') >= 0, 'saveRequest no acepta tcDocUrl');
  assert((crm.match(/tcDocUrl:/g) || []).length >= 2, 'getTicketStatus/getAllTickets no exponen tcDocUrl');
});

// 9) Lock anti-duplicados alrededor de generateTicketId
t('LockService en saveRequest y saveRappiCreditosRequest', () => {
  assert((crm.match(/LockService\.getScriptLock\(\)/g) || []).length === 2, 'faltan locks');
  assert((crm.match(/releaseLock\(\)/g) || []).length >= 2, 'faltan releases');
});

// 10) claspignore: los 4 archivos nuevos permitidos (si no, el push a /dev no los sube)
t('.claspignore permite los archivos CRM', () => {
  const ci = read('.claspignore');
  ['!Crm.gs.js', '!CrmForm.html', '!CrmStatus.html', '!CrmAdmin.html'].forEach(l =>
    assert(ci.indexOf(l) >= 0, 'falta ' + l));
});

console.log('\nF1 (integración CRM) — ' + passed + ' passed, ' + failed + ' failed');
if (failed) process.exit(1);
