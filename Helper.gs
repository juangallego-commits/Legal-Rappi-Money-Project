// =================================================================
// RAPPIMIND · COMMON HELPERS
// -----------------------------------------------------------------
// Pure-utility functions shared across the Apps Script project:
// date/number formatting, string manipulation, sheet I/O helpers.
// Anything with security implications lives in Security.gs.
// =================================================================

/**
 * Parse a yyyy-MM-dd string into a local Date. Returns null on bad input.
 * @param {string} dateString
 * @return {Date|null}
 */
function parseFormDate(dateString) {
  if (!dateString || typeof dateString !== 'string') return null;
  var parts = dateString.split('-');
  if (parts.length < 3) return null;
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  var d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return null;
  var dt = new Date(y, m - 1, d);
  return isNaN(dt.getTime()) ? null : dt;
}

/**
 * Title-case a string, preserving non-word characters.
 * @param {string} str
 * @return {string}
 */
function capitalize(str) {
  if (!str) return '';
  if (typeof str !== 'string') return str;
  return str.replace(/\w\S*/g, function(txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}

/**
 * Format a Date as "15 de marzo de 2026" (Spanish legal format).
 * @param {Date} date
 * @return {string|null}
 */
function formatDateInSpanish(date) {
  if (!date || isNaN(new Date(date).getTime())) return null;
  var d = new Date(date);
  return d.getDate() + ' de ' + MESES_ES[d.getMonth()] + ' de ' + d.getFullYear();
}

/**
 * Format a Date as "12:00 p.m." (Spanish legal format).
 * @param {Date} date
 * @return {string}
 */
function formatTimeInSpanish(date) {
  var d = new Date(date);
  if (isNaN(d.getTime())) return '';
  var hours = d.getHours();
  var minutes = d.getMinutes();
  var ampm = hours >= 12 ? 'p.m.' : 'a.m.';
  hours = hours % 12;
  hours = hours ? hours : 12;
  var minutesStr = minutes < 10 ? '0' + minutes : minutes;
  return hours + ':' + minutesStr + ' ' + ampm;
}

/**
 * Render an array as "a, b y c" (Spanish enumeration).
 * @param {Array<string>} list
 * @return {string}
 */
function formatListToText(list) {
  if (!Array.isArray(list)) return list;
  if (list.length === 0) return '';
  if (list.length === 1) return list[0];
  var copy = list.slice();
  var last = copy.pop();
  return copy.join(', ') + ' y ' + last;
}

/**
 * Validate campaign start/end dates from the legacy variable bag.
 * Throws with a user-facing message on failure.
 *
 * @param {Object<string, string>} vars
 */
function validateDates(vars) {
  var startCamp = parseFormDate(vars['Fecha de INICIO de Campaña']);
  var endCamp = parseFormDate(vars['Fecha de FIN de Campaña']);
  if (!startCamp || !endCamp) throw new Error('Fechas obligatorias.');
  if (endCamp.getTime() < startCamp.getTime()) throw new Error('Fecha Fin anterior a Inicio.');
}

/**
 * Convert a positive integer to its Spanish literal form ("dos mil cuarenta").
 * Returns "ERROR_NUMERO" for non-numeric input, "cero" for 0.
 *
 * @param {number} num
 * @return {string}
 */
function numeroALetras(num) {
  if (isNaN(num)) return 'ERROR_NUMERO';
  num = Math.floor(Number(num));
  if (num === 0) return 'cero';

  var unidades   = ['', 'un', 'dos', 'tres', 'cuatro', 'cinco', 'seis', 'siete', 'ocho', 'nueve'];
  var especiales = ['diez', 'once', 'doce', 'trece', 'catorce', 'quince', 'dieciséis', 'diecisiete', 'dieciocho', 'diecinueve'];
  var decenas    = ['', 'diez', 'veinte', 'treinta', 'cuarenta', 'cincuenta', 'sesenta', 'setenta', 'ochenta', 'noventa'];
  var centenas   = ['', 'ciento', 'doscientos', 'trescientos', 'cuatrocientos', 'quinientos', 'seiscientos', 'setecientos', 'ochocientos', 'novecientos'];

  function convertirGrupo(n) {
    var output = '';
    if (n === 100) return 'cien ';
    if (n > 100) {
      output += centenas[Math.floor(n / 100)] + ' ';
      n %= 100;
    }
    if (n >= 10 && n <= 19) {
      output += especiales[n - 10] + ' ';
      return output;
    }
    if (n >= 20) {
      output += decenas[Math.floor(n / 10)];
      if (n % 10 !== 0) {
        if (Math.floor(n / 10) === 2) output = 'veinti';
        else output += ' y ';
      } else {
        output += ' ';
      }
      n %= 10;
    }
    if (n > 0) output += unidades[n] + ' ';
    return output;
  }

  var texto = '';
  if (num >= 1000000) {
    var millones = Math.floor(num / 1000000);
    if (millones === 1) texto += 'un millón ';
    else texto += convertirGrupo(millones) + 'millones ';
    num %= 1000000;
  }
  if (num >= 1000) {
    var miles = Math.floor(num / 1000);
    if (miles === 1) texto += 'mil ';
    else texto += convertirGrupo(miles) + 'mil ';
    num %= 1000;
  }
  if (num > 0) texto += convertirGrupo(num);
  return texto.trim();
}

/**
 * Normalise common Apple brand names so they aren't lower-cased by
 * the generic capitalize() pass.
 *
 * @param {string} str
 * @return {string}
 */
function cleanTechNames(str) {
  if (!str) return str;
  var result = str;
  result = result.replace(/\biphone\b/gi, 'iPhone');
  result = result.replace(/\bipad\b/gi, 'iPad');
  result = result.replace(/\bios\b/gi, 'iOS');
  result = result.replace(/\bmacbook\b/gi, 'MacBook');
  result = result.replace(/\bmac\b(?!\s*book)/gi, 'Mac');
  result = result.replace(/\bairpods\b/gi, 'AirPods');
  result = result.replace(/\bapple\s+watch\b/gi, 'Apple Watch');
  return result;
}

/**
 * Share a freshly created Document publicly with view-only access and
 * return the canonical URL. Falls back to the editor URL if sharing
 * fails (e.g. when the script lacks Drive scopes).
 *
 * @param {GoogleAppsScript.Document.Document} doc
 * @return {string}
 */
function setPublicViewPermissions(doc) {
  try {
    var file = DriveApp.getFileById(doc.getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log('⚠️ setPublicViewPermissions failed: ' + e.message);
    return doc.getUrl();
  }
}

/**
 * Build a uniform response envelope returned over `google.script.run`.
 * Always returns a plain object (ContentService cannot cross that bridge).
 *
 * @param {boolean} success
 * @param {string} message
 * @param {*} [data]
 * @return {{success: boolean, message: string, data: *}}
 */
function buildResponse(success, message, data) {
  return { success: !!success, message: message || '', data: data === undefined ? null : data };
}

/**
 * Get the Sheet ID used by the audit/registry spreadsheet. Looks up
 * Script Properties first (so ops can rotate without code change),
 * falls back to the legacy hard-coded constant.
 *
 * @return {string}
 */
function _resolveSheetId() {
  if (typeof resolveAuditSheetId === 'function') {
    try { return resolveAuditSheetId(); } catch (e) { /* fall through */ }
  }
  // Legacy constant from Config.gs
  if (typeof AUDIT_SHEET_ID !== 'undefined' && AUDIT_SHEET_ID) return AUDIT_SHEET_ID;
  throw new Error('No se pudo resolver AUDIT_SHEET_ID. Configúralo en Script Properties.');
}

/**
 * Open (or lazily create) a tab in the audit spreadsheet. If the tab
 * is missing, headers are written and styled in one shot.
 *
 * @param {string} name
 * @param {Array<string>} headers
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function _getOrCreateSheet(name, headers) {
  var ss;
  try {
    ss = SpreadsheetApp.openById(_resolveSheetId());
  } catch (e) {
    throw new Error('No se pudo abrir el spreadsheet de auditoría: ' + e.message);
  }
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Open an existing sheet by name. Returns null when the sheet (or the
 * spreadsheet itself) cannot be opened. Use this when the caller wants
 * to handle "missing" gracefully.
 *
 * @param {string} name
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function _getSheet(name) {
  try {
    var ss = SpreadsheetApp.openById(_resolveSheetId());
    return ss.getSheetByName(name);
  } catch (e) {
    Logger.log('⚠️ _getSheet("' + name + '") failed: ' + e.message);
    return null;
  }
}

/**
 * Defensive payload check that flags fields whose stringified value
 * literally contains the substring "undefined" — which usually means
 * a backend template referenced a missing key.
 *
 * NOTE: This intentionally does not throw on null/empty strings; only
 * on the "undefined" sentinel.
 *
 * @param {Object} data
 */
function auditData(data) {
  if (!data || typeof data !== 'object') throw new Error('Payload vacío.');
  var keys = Object.keys(data);
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var v = data[k];
    if (v === undefined) throw new Error('Dato faltante (undefined): ' + k);
    if (typeof v === 'string' && /\bundefined\b/.test(v)) {
      throw new Error('Dato corrupto, contiene la cadena "undefined": ' + k);
    }
  }
}

/**
 * Convert a Sheet of rows into an array of objects keyed by the header row.
 * Trims headers; coerces numbers/strings naturally.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Array<Object>}
 */
function _sheetToObjects(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return String(h).trim(); });
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}
