// Tests F2a — nombres automáticos + condiciones especiales (funciones puras con marcadores).
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(root, f), 'utf8');
function extractBetween(file, s0, e0) {
  const src = read(file);
  let s = src.indexOf(s0); if (s < 0) throw new Error('marcador inicio ' + s0 + ' en ' + file);
  s = src.indexOf('\n', s) + 1;
  const e = src.indexOf(e0, s); if (e < 0) throw new Error('marcador fin ' + e0);
  return src.slice(s, e);
}

const ctx = {};
vm.createContext(ctx);
vm.runInContext(extractBetween('Código.js', '//==NAMING_START==', '//==NAMING_END=='), ctx);
vm.runInContext(extractBetween('Config.gs.js', '//==SPECIAL_COND_START==', '//==SPECIAL_COND_END=='), ctx);
vm.runInContext(extractBetween('Código.js', '//==SPECIAL_COND_FN_START==', '//==SPECIAL_COND_FN_END=='), ctx);
const { _buildCampaignNames, _slugify, _buildSpecialConditionsText, _validateSpecialConditions, SPECIAL_CONDITIONS_CATALOG } = ctx;

let passed = 0, failed = 0;
function t(n, fn) { try { fn(); console.log('  ✓ ' + n); passed++; } catch (e) { console.log('  ✗ ' + n + ' — ' + e.message); failed++; } }
function eq(a, b, m) { if (a !== b) throw new Error((m || '') + ' esperaba "' + b + '" fue "' + a + '"'); }

// ---- Nombres ----
t('doc title: marca + mes/año, sin duplicar marca', () => {
  const n = _buildCampaignNames({ countryCode: 'co', campaignType: 'Cashback', brand: "McDonald's", startDate: '2026-11-15' });
  eq(n.publicTitle, "Términos y Condiciones — McDonald's (Noviembre 2026)");
});
t('doc title: nombre comercial distinto de la marca', () => {
  const n = _buildCampaignNames({ countryCode: 'CO', campaignType: 'Cashback', brand: 'KFC', campaignName: 'Mega Cashback', startDate: '2026-03-01' });
  eq(n.publicTitle, 'Términos y Condiciones — Mega Cashback — KFC (Marzo 2026)');
});
t('interno: CC · Tipo · Marca · AAAA-MM (+ ticket)', () => {
  const n = _buildCampaignNames({ countryCode: 'co', campaignType: 'Cashback', brand: 'KFC', startDate: '2026-03-09', ticketId: 'CAM-0042' });
  eq(n.internalName, 'CO · Cashback · KFC · 2026-03 · CAM-0042');
});
t('interno sin ticket: no agrega separador colgante', () => {
  const n = _buildCampaignNames({ countryCode: 'CO', campaignType: 'Cashback', brand: 'KFC', startDate: '2026-03-09' });
  eq(n.internalName, 'CO · Cashback · KFC · 2026-03');
});
t('slug: kebab-case ascii sin acentos ni apóstrofes', () => {
  const n = _buildCampaignNames({ countryCode: 'CO', campaignType: 'Cashback', brand: "McDonald's Ñandú", startDate: '2026-11-15' });
  eq(n.slug, 'mcdonalds-nandu-cashback-2026-11');
});
t('sin fecha: título sin paréntesis, interno sin AAAA-MM', () => {
  const n = _buildCampaignNames({ countryCode: 'CO', campaignType: 'Cashback', brand: 'KFC' });
  eq(n.publicTitle, 'Términos y Condiciones — KFC');
  eq(n.internalName, 'CO · Cashback · KFC');
});
t('_slugify determinista y colapsa separadores', () => {
  eq(_slugify('  Hola   Mundo!! '), 'hola-mundo');
  eq(_slugify('Á é í ó ú'), 'a-e-i-o-u');
});

// ---- Condiciones especiales ----
t('catálogo: toda llave tiene label y texto', () => {
  const keys = Object.keys(SPECIAL_CONDITIONS_CATALOG);
  if (keys.length < 5) throw new Error('catálogo incompleto');
  keys.forEach(k => { const it = SPECIAL_CONDITIONS_CATALOG[k]; if (!it.label || !it.texto) throw new Error('falta label/texto en ' + k); });
});
t('ninguna seleccionada → ""', () => eq(_buildSpecialConditionsText([]), ''));
t('una → una viñeta', () => {
  const s = _buildSpecialConditionsText(['no_combos']);
  if (s.indexOf('• ') !== 0 || s.indexOf('combo') < 0) throw new Error('viñeta mal: ' + s);
  if (s.indexOf('\n') !== -1) throw new Error('no debería tener salto con una sola');
});
t('varias → viñetas separadas por salto de línea, en orden', () => {
  const s = _buildSpecialConditionsText(['limite_100_dia', 'tope_50pct']);
  const lineas = s.split('\n');
  eq(lineas.length, 2);
  if (lineas[0].indexOf('cien (100)') < 0 || lineas[1].indexOf('50%') < 0) throw new Error('orden/contenido: ' + s);
});
t('dedupe: la misma llave dos veces no repite viñeta', () => {
  eq(_buildSpecialConditionsText(['no_combos', 'no_combos']).split('\n').length, 1);
});
t('validación: local + domicilio son excluyentes', () => {
  const r = _validateSpecialConditions(['consumo_local', 'envio_domicilio']);
  if (r.ok) throw new Error('debería rechazar local+domicilio');
});
t('validación: combinación válida pasa', () => {
  const r = _validateSpecialConditions(['consumo_local', 'no_combos', 'tope_50pct']);
  if (!r.ok) throw new Error('debería aceptar: ' + JSON.stringify(r));
});

// ---- Wiring de las APIs GAS ----
t('Código.js expone previewCampaignNames y getSpecialConditionsCatalog', () => {
  const cod = read('Código.js');
  if (!/function\s+previewCampaignNames\s*\(/.test(cod)) throw new Error('falta previewCampaignNames');
  if (!/function\s+getSpecialConditionsCatalog\s*\(/.test(cod)) throw new Error('falta getSpecialConditionsCatalog');
});

console.log('\nF2a (nombres + condiciones) — ' + passed + ' passed, ' + failed + ' failed');
if (failed) process.exit(1);
