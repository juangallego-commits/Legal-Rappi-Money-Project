// Tests — catálogo de T&C GENERALES (regla pura extraída de Código.js con marcadores).
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(root, f), 'utf8');

function extractBetween(file, startMarker, endMarker) {
  const src = read(file);
  let s = src.indexOf(startMarker);
  if (s < 0) throw new Error('Falta marcador inicio en ' + file);
  s = src.indexOf('\n', s) + 1;
  const e = src.indexOf(endMarker, s);
  if (e < 0) throw new Error('Falta marcador fin en ' + file);
  return src.slice(s, e);
}

const ctx = {};
vm.createContext(ctx);
vm.runInContext(extractBetween('Código.js', '//==TCGEN_START==', '//==TCGEN_END=='), ctx);
const pick = ctx._pickGeneralTcKey;

let passed = 0, failed = 0;
function t(name, fn) {
  try { fn(); console.log('  ✓ ' + name); passed++; }
  catch (e) { console.log('  ✗ ' + name + ' — ' + e.message); failed++; }
}
function eq(a, b, msg) { if (a !== b) throw new Error((msg || '') + ' esperaba "' + b + '", fue "' + a + '"'); }

t('percentage → dto_tienda', () => eq(pick('percentage', {}), 'dto_tienda'));
t('percentage + PRIME SI → dto_tienda_pro', () => eq(pick('percentage', { prime: 'SI' }), 'dto_tienda_pro'));
t('percentage + productos → dto_prod', () => eq(pick('percentage', { hasProducts: true }), 'dto_prod'));
t('percentage + productos + PRO → dto_prod_pro', () => eq(pick('percentage', { hasProducts: true, prime: 'SI' }), 'dto_prod_pro'));
t('value → dto_valor', () => eq(pick('value', {}), 'dto_valor'));
t('value + PRO → dto_valor_pro', () => eq(pick('value', { prime: 'SI' }), 'dto_valor_pro'));
t('free_shipping → envio_gratis', () => eq(pick('free_shipping', {}), 'envio_gratis'));
t('service_fee → tarifa_servicio', () => eq(pick('service_fee', {}), 'tarifa_servicio'));
t('offer_by_product → dto_prod', () => eq(pick('offer_by_product', {}), 'dto_prod'));
t('offer_by_product + PRO → dto_prod_pro', () => eq(pick('offer_by_product', { prime: 'SI' }), 'dto_prod_pro'));
t('PRIME TODOS/NO ≠ versión Pro', () => { eq(pick('value', { prime: 'TODOS' }), 'dto_valor'); eq(pick('value', { prime: 'NO' }), 'dto_valor'); });
t('cashback → "" (T&C específico)', () => eq(pick('cashback', {}), ''));
t('rappicreditos → "" (T&C específico)', () => eq(pick('rappicreditos', {}), ''));
t('desconocido/vacío → ""', () => { eq(pick('otra_cosa', {}), ''); eq(pick('', {}), ''); eq(pick(null, {}), ''); });

// Seed: cada key que la regla puede devolver existe en el seed de Setup.gs.js con URL
t('seed CO cubre todas las keys de la regla', () => {
  const setup = read('Setup.gs.js');
  const m = setup.match(/var TC_GENERALES_SEED_CO = (\[[\s\S]*?\]);/);
  if (!m) throw new Error('falta TC_GENERALES_SEED_CO');
  const seed = JSON.parse(m[1]);
  const keys = new Set(seed.map(s => s.key));
  ['dto_tienda', 'dto_prod', 'dto_valor', 'envio_gratis', 'tarifa_servicio',
   'dto_tienda_pro', 'dto_prod_pro', 'dto_valor_pro'].forEach(k => {
    if (!keys.has(k)) throw new Error('falta key en seed: ' + k);
  });
  seed.forEach(s => { if (!/^https:\/\/promos\.rappi\.com\//.test(s.url)) throw new Error('URL rara en ' + s.key); });
});

t('Config define TC_GENERALES_SHEET y override RAPPIMIND_DB_ID', () => {
  const cfg = read('Config.gs.js');
  if (cfg.indexOf("TC_GENERALES_SHEET = 'TC_Generales'") < 0) throw new Error('falta TC_GENERALES_SHEET');
  if (cfg.indexOf("RAPPIMIND_DB_ID") < 0) throw new Error('falta override RAPPIMIND_DB_ID');
});

t('Código.js expone getGeneralTcCatalog y Setup expone setupTcGenerales', () => {
  if (!/function\s+getGeneralTcCatalog\s*\(/.test(read('Código.js'))) throw new Error('falta getGeneralTcCatalog');
  if (!/function\s+setupTcGenerales\s*\(/.test(read('Setup.gs.js'))) throw new Error('falta setupTcGenerales');
});

// ---- F2b: sincronía backend↔frontend de la regla + wiring de la página/autofill ----
t('espejo JS en CrmForm produce EXACTAMENTE lo mismo que el backend', () => {
  const jsCtx = {};
  vm.createContext(jsCtx);
  vm.runInContext(extractBetween('CrmForm.html', '//==TCGEN_JS_START==', '//==TCGEN_JS_END=='), jsCtx);
  const pickJS = jsCtx._pickGeneralTcKeyJS;
  const casos = [];
  ['percentage', 'value', 'free_shipping', 'service_fee', 'offer_by_product', 'cashback', 'rappicreditos', ''].forEach(tipo => {
    ['SI', 'NO', 'TODOS'].forEach(prime => {
      [true, false].forEach(hasProducts => casos.push({ tipo, prime, hasProducts }));
    });
  });
  casos.forEach(c => {
    const a = pick(c.tipo, { prime: c.prime, hasProducts: c.hasProducts });
    const b = pickJS(c.tipo, { prime: c.prime, hasProducts: c.hasProducts });
    if (a !== b) throw new Error(JSON.stringify(c) + ': backend="' + a + '" vs front="' + b + '"');
  });
});

t('F2b wiring: página TcGenerales + router + autofill + aliados', () => {
  const tcPage = read('TcGenerales.html');
  if (tcPage.indexOf('getGeneralTcCatalog') < 0) throw new Error('TcGenerales no llama getGeneralTcCatalog');
  if (tcPage.indexOf('getScriptUrl') < 0) throw new Error('TcGenerales no resuelve la URL');
  const codigo = read('Código.js');
  if (codigo.indexOf("_crmServePage('TcGenerales'") < 0) throw new Error('router sin tc-generales');
  if (!/function\s+getAliadosCatalog\s*\(/.test(codigo)) throw new Error('falta getAliadosCatalog');
  if (!/function\s+_learnAliadoFromPayload\s*\(/.test(codigo)) throw new Error('falta _learnAliadoFromPayload');
  if (codigo.indexOf('_learnAliadoFromPayload(payload)') < 0) throw new Error('processWebPayload no aprende aliados');
  if (read('Config.gs.js').indexOf("ALIADOS_SHEET = 'Aliados'") < 0) throw new Error('falta ALIADOS_SHEET');
  const form = read('CrmForm.html');
  if (form.indexOf('_autofillGeneralTC') < 0 || form.indexOf('getGeneralTcCatalog') < 0) throw new Error('CrmForm sin autollenado');
  if (read('.claspignore').indexOf('!TcGenerales.html') < 0) throw new Error('claspignore sin TcGenerales');
});

console.log('\nT&C Generales — ' + passed + ' passed, ' + failed + ' failed');
if (failed) process.exit(1);
