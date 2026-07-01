// Tests FASE C (node) — NO se suben a Apps Script (.claspignore solo permite los 7 archivos).
// Extraen el código REAL entre marcadores de Config.gs.js y Admin.gs.js y lo evalúan para
// probar la clasificación determinista de placeholders (sin Gemini, sin tocar la hoja).
//
// Correr:  node tests/field_catalog.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function extractBetween(file, startMarker, endMarker) {
  const src = fs.readFileSync(file, 'utf8');
  let s = src.indexOf(startMarker);
  if (s < 0) throw new Error('Falta el marcador de inicio en ' + file);
  s = src.indexOf('\n', s) + 1;            // empezar DESPUÉS de la línea del marcador
  const e = src.indexOf(endMarker, s);
  if (e < 0) throw new Error('Falta el marcador de fin en ' + file);
  return src.slice(s, e);
}

const root = path.join(__dirname, '..');
const catalog = extractBetween(path.join(root, 'Config.gs.js'), '//==FASE_C_CATALOG_START==', '//==FASE_C_CATALOG_END==');
const pure = extractBetween(path.join(root, 'Admin.gs.js'), '//==FASE_C_PURE_START==', '//==FASE_C_PURE_END==');
const optBlocks = extractBetween(path.join(root, 'Código.js'), '//==FASE_B_OPT_START==', '//==FASE_B_OPT_END==');
const sentence = extractBetween(path.join(root, 'Config.gs.js'), '//==FASE_D_SENTENCE_START==', '//==FASE_D_SENTENCE_END==');

// Fragmento representativo del template real de Cashback (organizador + base + derived + legal + desconocido).
const CASHBACK_SNIPPET = [
  'La presente campaña es organizada por {{ORGANIZADOR}}, identificada con {{ID_ORGANIZADOR}}.',
  'Válida en {{TEXTO_TERRITORIO}} desde el {{FECHA_INICIO}}. Podrán participar {{TEXTO_SEGMENTO}}.',
  'Se rige por {{LEY_APLICABLE}} y se somete a {{JURISDICCION}} ante {{ENTIDAD_VIGILANCIA}}.',
  'Campo nuevo aún no catalogado: {{FOO_DESCONOCIDO}}.'
].join('\n');

const results = [];
const testCode = `
  var out = deriveFieldRowsForTemplate(CASHBACK_SNIPPET, 'Cashback', 'CO');
  function findRow(id){ return out.rows.filter(function(r){return r.field_id===id;})[0]; }
  function skippedAs(cat, tok){ return out.skipped.filter(function(s){return s.token===tok && s.category===cat;}).length>0; }

  results.push(['ORGANIZADOR -> input OBLIGATORIO', !!findRow('organizerLegalName') && findRow('organizerLegalName').required==='TRUE']);
  results.push(['ID_ORGANIZADOR -> input OBLIGATORIO', !!findRow('organizerTaxId') && findRow('organizerTaxId').required==='TRUE']);
  results.push(['JURISDICCION NO genera campo (legal)', skippedAs('legal','JURISDICCION')]);
  results.push(['LEY_APLICABLE NO genera campo (legal)', skippedAs('legal','LEY_APLICABLE')]);
  results.push(['ENTIDAD_VIGILANCIA NO genera campo (legal)', skippedAs('legal','ENTIDAD_VIGILANCIA')]);
  results.push(['TEXTO_SEGMENTO NO genera campo (derived)', skippedAs('derived','TEXTO_SEGMENTO')]);
  results.push(['TEXTO_TERRITORIO NO genera campo (derived)', skippedAs('derived','TEXTO_TERRITORIO')]);
  results.push(['FECHA_INICIO NO genera campo (base)', skippedAs('base','FECHA_INICIO')]);
  results.push(['token DESCONOCIDO -> input, needs_review=TRUE, NO obligatorio', !!findRow('foo_desconocido') && findRow('foo_desconocido').needs_review==='TRUE' && findRow('foo_desconocido').required==='FALSE']);
  results.push(['unknown reportado en out.unknown', out.unknown.indexOf('FOO_DESCONOCIDO')>=0]);
  results.push(['_extractPlaceholders dedup + tolera espacios', (function(){ var t=_extractPlaceholders('{{ A }} {{A}} {{B}}'); return t.length===2 && t[0]==='A' && t[1]==='B'; })()]);
  results.push(['_classifyPlaceholder(ORGANIZADOR).required===true', _classifyPlaceholder('ORGANIZADOR').required===true]);
  results.push(['required NUNCA lo pone la IA: derived no aparece como campo', !findRow('texto_segmento')]);

  // ---- FASE B: bloques opcionales (sentinelas) ----
  var vacio = function(t){ return t === 'ORGANIZADOR'; };
  results.push(['B: bloque con token VACÍO se borra completo', _stripOptionalBlocks('X [[?ORGANIZADOR]]por Y, id Z[[/?]] W', vacio) === 'X  W']);
  results.push(['B: bloque con token CON VALOR conserva contenido y quita markers', _stripOptionalBlocks('X [[?ORGANIZADOR]]por Y[[/?]] W', function(){return false;}) === 'X por Y W']);
  results.push(['B: sin markers => no-op (template actual)', _stripOptionalBlocks('organizada por {{ORGANIZADOR}}.', vacio) === 'organizada por {{ORGANIZADOR}}.']);
  results.push(['B: opcional vacío NO deja restos de coma/preposición', _stripOptionalBlocks('organizada[[?ORGANIZADOR]] por {{ORGANIZADOR}}, identificada con {{ID_ORGANIZADOR}}[[/?]].', vacio) === 'organizada.']);

  // ---- FASE D: contrato de oración ----
  results.push(['D: capitaliza y agrega punto', _asSentence('pueden participar todos') === 'Pueden participar todos.']);
  results.push(['D: idempotente (ya cumple)', _asSentence('Pueden participar todos.') === 'Pueden participar todos.']);
  results.push(['D: recorta espacios', _asSentence('   se aclara que...  ') === 'Se aclara que...']);
  results.push(['D: respeta signos ? !', _asSentence('¿aplica?') === '¿aplica?']);
  results.push(['D: vacío => vacío', _asSentence('') === '']);
`;

const sandbox = { CASHBACK_SNIPPET, results, console };
vm.runInNewContext(catalog + '\n' + pure + '\n' + optBlocks + '\n' + sentence + '\n' + testCode, sandbox, { timeout: 5000 });

let passed = 0, failed = 0;
for (const [name, ok] of results) {
  console.log((ok ? '  ✓ ' : '  ✗ ') + name);
  ok ? passed++ : failed++;
}
console.log('\nFASE C (field catalog) — ' + passed + ' passed, ' + failed + ' failed');
process.exit(failed ? 1 : 0);
