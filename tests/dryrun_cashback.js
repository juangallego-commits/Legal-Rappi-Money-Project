// SIMULACIÓN del dry-run FASE C para Cashback — evalúa el código REAL (catálogo + motor puro)
// contra los datos REALES de la hoja en vivo (Template_Fields, 2026-07-01) y los placeholders
// REALES del template de Cashback (doc 1bRBQ2b_u6NrhjxoNE5-le5-GEXzvdCcmAqtD57sAwxM).
// NO escribe nada: solo imprime el diff que produciría previewFieldDerivation('Cashback','ALL').
//
// Correr:  node tests/dryrun_cashback.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function extractBetween(file, startMarker, endMarker) {
  const src = fs.readFileSync(file, 'utf8');
  let s = src.indexOf(startMarker);
  if (s < 0) throw new Error('Falta marcador inicio en ' + file);
  s = src.indexOf('\n', s) + 1;
  const e = src.indexOf(endMarker, s);
  if (e < 0) throw new Error('Falta marcador fin en ' + file);
  return src.slice(s, e);
}

const root = path.join(__dirname, '..');
const catalog = extractBetween(path.join(root, 'Config.gs.js'), '//==FASE_C_CATALOG_START==', '//==FASE_C_CATALOG_END==');
const pure = extractBetween(path.join(root, 'Admin.gs.js'), '//==FASE_C_PURE_START==', '//==FASE_C_PURE_END==');

// Placeholders REALES del template Cashback (extraídos del Doc en vivo).
const CASHBACK_TOKENS = ['NOMBRE_CAMPANA_UPPER','ORGANIZADOR','ID_ORGANIZADOR','TEXTO_TERRITORIO','REF_TIENDA','HORA_INICIO','FECHA_INICIO','HORA_FIN','FECHA_FIN','PAIS_LEGAL','PRESUPUESTO_LETRAS','PRESUPUESTO_NUM','TEXTO_SEGMENTO','TEXTO_METODO_PAGO','TEXTO_PORCENTAJE','TOPE_LETRAS','TOPE_NUM','LIMITE_ORDENES','TEXTO_ORDENES','TEXTO_CARGA','URL_TC_CREDITOS','TEXTO_VIGENCIA_CREDITOS','TEXTO_LUGAR_REDENCION','URL_TC_PLATAFORMA','ENTIDAD_LEGAL','NOMBRE_POLITICA_PRIVACIDAD','URL_PRIVACIDAD','LEY_APLICABLE','ENTIDAD_VIGILANCIA','JURISDICCION'];
const CASHBACK_TEXT = CASHBACK_TOKENS.map(t => '{{' + t + '}}').join(' ');

// Template_Fields REALES en vivo (subagente). f(field_id, campaign_type, placeholder, required)
function f(id, ct, ph, req) { return { field_id: id, campaign_type: ct, placeholder: ph, required: req || '', field_type: 'text' }; }
const EXISTING = [
  // campaign_type = 'ALL' (14) — NO deben tocarse
  f('userEmail','ALL',''), f('dynamicType','ALL',''), f('campaignName','ALL','{{NOMBRE_CAMPANA}}','FALSE'),
  f('shopName','ALL','{{TIENDA_BASE}}','TRUE'), f('territory','ALL','{{TEXTO_TERRITORIO}}','TRUE'),
  f('startDate','ALL','{{FECHA_INICIO}}','TRUE'), f('endDate','ALL','{{FECHA_FIN}}','TRUE'),
  f('paymentMethods','ALL','{{TEXTO_METODO_PAGO}}','FALSE'), f('userSegment','ALL','{{TEXTO_SEGMENTO}}','FALSE'),
  f('maxOrders','ALL','{{LIMITE_ORDENES}}','FALSE'), f('specialConditions','ALL','{{CONDICIONES_ESPECIALES}}','FALSE'),
  f('startTime','ALL',''), f('endTime','ALL',''), f('extraEmails','ALL',''),
  // campaign_type = 'Cashback' (11)
  f('cashbackPct','Cashback','{{TEXTO_PORCENTAJE}}','TRUE'), f('cap','Cashback','{{TOPE_NUM}}','TRUE'),
  f('budget','Cashback','{{PRESUPUESTO_NUM}}','TRUE'), f('redemptionPlace','Cashback','{{TEXTO_LUGAR_REDENCION}}','FALSE'),
  f('loadType','Cashback',''), f('loadDate','Cashback',''), f('validityType','Cashback',''),
  f('validityDays','Cashback',''), f('redemptionStart','Cashback',''), f('redemptionEnd','Cashback',''), f('minPurchase','Cashback',''),
  // campaign_type = 'Concurso Mayor Comprador' (16) — NO deben tocarse
  f('organizerLegalName','Concurso Mayor Comprador','{{ORGANIZADOR}}','TRUE'),
  f('organizerPhone','Concurso Mayor Comprador','{{TELEFONO_CONTACTO}}','TRUE'),
  f('organizerEmail','Concurso Mayor Comprador','{{EMAIL_CONTACTO}}','TRUE'),
  f('numberOfWinners','Concurso Mayor Comprador','{{NUM_GANADORES}}','TRUE'),
  f('winnerCriteria','Concurso Mayor Comprador','{{CRITERIO_GANADOR}}','TRUE'),
  f('announcementDate','Concurso Mayor Comprador','{{FECHA_ANUNCIO}}','TRUE'),
  f('verticals','Concurso Mayor Comprador','{{VERTICALES}}','FALSE'),
  f('participatingProducts','Concurso Mayor Comprador','{{PRODUCTOS_PARTICIPANTES}}','FALSE'),
  f('prizeType','Concurso Mayor Comprador','','TRUE'), f('creditsAmount','Concurso Mayor Comprador',''),
  f('creditLoadDays','Concurso Mayor Comprador',''), f('creditsValidityDays','Concurso Mayor Comprador',''),
  f('creditsRedemptionPlace','Concurso Mayor Comprador',''), f('physicalPrizeDescription','Concurso Mayor Comprador',''),
  f('prizeDeliveryBy','Concurso Mayor Comprador',''), f('minParticipation','Concurso Mayor Comprador','')
];

const sandbox = { CASHBACK_TEXT, EXISTING, console };
const harness = catalog + '\n' + pure + '\n' +
  'var derived = deriveFieldRowsForTemplate(CASHBACK_TEXT, "Cashback", "ALL");' +
  'var tokens = _extractPlaceholders(CASHBACK_TEXT);' +
  'var diff = _computeFieldDiff(derived.rows, EXISTING, "Cashback", tokens);' +
  'globalThis.__derived = derived; globalThis.__diff = diff;';
vm.runInNewContext(harness, sandbox);
const derived = sandbox.__derived, diff = sandbox.__diff;

console.log('================ DRY-RUN FASE C — Cashback / ALL (doc 1bRBQ2b…) ================\n');
console.log('AGREGA (' + diff.adds.length + ' filas nuevas, campaign_type=Cashback):');
diff.adds.forEach(function(a) {
  console.log('  + field_id=' + a.field_id + '  placeholder=' + a.placeholder + '  required=' + a.required +
              '  type=' + a.field_type + '  label="' + a.label_es + '"  section=' + a.section + '  group=' + a.group);
});
console.log('\nMODIFICA: ' + diff.wouldModify.length + '   (reportado, NO aplicado)');
diff.wouldModify.forEach(function(m){ console.log('  ~ ' + JSON.stringify(m)); });
console.log('\nBORRA: ' + diff.deletes.length + '   |   ORPHANS (reportado, NO borrado): ' + diff.orphans.length);
console.log('\nFilas de OTROS campaign_type que NO se tocan: ' + JSON.stringify(diff.untouchedOtherTypes) +
            '  (total ' + diff.untouchedOtherCount + ')');
console.log('Placeholders no catalogados: ' + (derived.unknown.length ? derived.unknown.join(', ') : 'ninguno'));

// Confirmaciones automáticas exigidas
console.log('\n---------------- CONFIRMACIONES ----------------');
var touchesALL = diff.adds.concat(diff.orphans).some(function(r){ return String(r.campaign_type) === 'ALL'; });
var touchesConcurso = diff.adds.concat(diff.orphans).some(function(r){ return String(r.campaign_type) === 'Concurso Mayor Comprador'; });
var org = diff.adds.filter(function(a){ return a.field_id === 'organizerLegalName'; })[0];
var orgId = diff.adds.filter(function(a){ return a.field_id === 'organizerTaxId'; })[0];
var concursoOrg = EXISTING.filter(function(r){ return r.field_id === 'organizerLegalName' && r.campaign_type === 'Concurso Mayor Comprador'; })[0];
console.log((!touchesALL && !touchesConcurso ? '✓' : '✗') + ' El diff NO agrega/borra ninguna fila ALL ni Concurso');
console.log((org && org.required === 'TRUE' && org.field_type === 'text' ? '✓' : '✗') + ' {{ORGANIZADOR}} → input required (organizerLegalName, text)');
console.log((orgId && orgId.required === 'TRUE' && orgId.field_type === 'text' ? '✓' : '✗') + ' {{ID_ORGANIZADOR}} → input required (organizerTaxId, text)');
console.log((org && concursoOrg && org.field_id === concursoOrg.field_id ? '✓' : '✗') + ' organizerLegalName reusa el MISMO field_id canónico que Concurso (no duplicado nuevo)');
console.log((diff.scopeOk ? '✓' : '✗') + ' scopeOk (write acotado a campaign_type=Cashback)');
