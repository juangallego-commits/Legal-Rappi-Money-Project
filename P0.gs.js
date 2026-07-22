// =================================================================
// P0 — FUNCIONES DE UN CLIC para el runbook (docs/RUNBOOK_P0_PASO_A_PASO.md).
// En el editor de Apps Script: abrir este archivo (P0.gs), elegir la función
// en el selector de arriba y presionar "Ejecutar". Ninguna pide argumentos.
// Correr en orden: P0_0 → P0_1 → P0_2 → P0_3 (dos veces) → P0_4 → probar.
// =================================================================

// Muestra a qué DB apunta este proyecto y qué configuración falta. NO escribe nada.
function P0_0_estado() {
  var p = PropertiesService.getScriptProperties();
  var db = p.getProperty('RAPPIMIND_DB_ID');
  Logger.log('=== ESTADO RAPPIMIND ===');
  Logger.log('DB del generador: ' + (db ? db + '  → COPIA (/dev) ✓' : AUDIT_SHEET_ID + '  → ⚠️ PRODUCCIÓN'));
  Logger.log('A2 (validación pre-entrega): ' + String(p.getProperty('RAPPIMIND_A2') || 'off'));
  ['CRM_SPREADSHEET_ID', 'CRM_SLACK_BOT_TOKEN', 'CRM_SLACK_CHANNEL', 'CRM_CSV_FOLDER_ID'].forEach(function (k) {
    Logger.log(k + ': ' + (p.getProperty(k) ? 'configurada ✓' : 'FALTA'));
  });
  try {
    var tf = _getSheet('Template_Fields');
    var n = tf ? _sheetToObjects(tf).filter(function (r) { return String(r.campaign_type) === 'Cashback'; }).length : 0;
    Logger.log('Template_Fields de Cashback: ' + n + ' filas (11 = falta Organizador → correr P0_3; 13 = ya aplicado)');
  } catch (e) { Logger.log('Template_Fields: ' + e.message); }
  try {
    var tg = _getSheet(TC_GENERALES_SHEET);
    Logger.log('TC_Generales: ' + (tg ? (_sheetToObjects(tg).length + ' filas') : 'hoja no existe → correr P0_1'));
  } catch (e) { Logger.log('TC_Generales: ' + e.message); }
}

// Siembra el catálogo de T&C generales (idempotente: re-correr no duplica).
function P0_1_sembrarTcGenerales() {
  Logger.log(JSON.stringify(setupTcGenerales()));
}

// FASE C — SOLO MIRA (no escribe): qué cambiaría en Template_Fields de Cashback.
// Esperado: adds = 2 (organizador y su ID), wouldModify = 0, deletes = 0.
function P0_2_previewOrganizador() {
  Logger.log(JSON.stringify(previewFieldDerivation('Cashback', 'ALL'), null, 2));
}

// FASE C — ESCRIBE los 2 campos de Organizador (ADD-only, idempotente).
// Correrla DOS veces: la primera → added: 2; la segunda → added: 0 (prueba de idempotencia).
function P0_3_aplicarOrganizador() {
  Logger.log(JSON.stringify(applyFieldDerivation('Cashback', 'ALL'), null, 2));
}

// Enciende A2: si un documento queda con {{}} o [ ] sin resolver, NO se entrega.
function P0_4_encenderA2() {
  PropertiesService.getScriptProperties().setProperty('RAPPIMIND_A2', 'on');
  Logger.log('A2 = on ✓ (apagar con P0_5 si hiciera falta)');
}

function P0_5_apagarA2() {
  PropertiesService.getScriptProperties().setProperty('RAPPIMIND_A2', 'off');
  Logger.log('A2 = off');
}
