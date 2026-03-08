function parseFormDate(dateString) {
  if (!dateString || typeof dateString !== 'string') return null;
  const parts = dateString.split('-');
  if (parts.length < 3) return null;
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}
function capitalize(str) {
  if (!str) return '';
  if (typeof str !== 'string') return str;
  return str.replace(/\w\S*/g, txt => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
}
function formatDateInSpanish(date) {
  if (!date) return null;
  return `${date.getDate()} de ${MESES_ES[date.getMonth()]} de ${date.getFullYear()}`;
}
function formatTimeInSpanish(date) {
  const d = new Date(date);
  let hours = d.getHours();
  const minutes = d.getMinutes();
  const ampm = hours >= 12 ? 'p.m.' : 'a.m.';
  hours = hours % 12;
  hours = hours ? hours : 12;
  const minutesStr = minutes < 10 ? '0' + minutes : minutes;
  return `${hours}:${minutesStr} ${ampm}`;
}
function formatListToText(list) {
  if (!Array.isArray(list)) return list;
  if (list.length === 0) return '';
  if (list.length === 1) return list[0];
  const last = list.pop();
  return list.join(', ') + ' y ' + last;
}

function validateDates(vars) {
  const startCamp = parseFormDate(vars['Fecha de INICIO de Campaña']);
  const endCamp = parseFormDate(vars['Fecha de FIN de Campaña']);
  if (!startCamp || !endCamp) throw new Error("Fechas obligatorias.");
  if (endCamp.getTime() < startCamp.getTime()) throw new Error("Fecha Fin anterior a Inicio.");
}

function numeroALetras(num) {
  if (isNaN(num)) return "ERROR_NUMERO";
  if (num === 0) return "cero";
  
  const unidades = ["", "un", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"];
  const especiales = ["diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve"];
  const decenas = ["", "diez", "veinte", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"];
  const centenas = ["", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"];
  
  function convertirGrupo(n) {
    let output = "";
    if (n === 100) return "cien ";
    if (n > 100) {
      output += centenas[Math.floor(n / 100)] + " ";
      n %= 100;
    }
    if (n >= 10 && n <= 19) {
      output += especiales[n - 10] + " ";
      return output;
    }
    if (n >= 20) {
      output += decenas[Math.floor(n / 10)];
      if (n % 10 !== 0) {
        if (Math.floor(n / 10) === 2) output = "veinti";
        else output += " y ";
      } else {
        output += " ";
      }
      n %= 10;
    }
    if (n > 0) output += unidades[n] + " ";
    return output;
  }
  
  let texto = "";
  if (num >= 1000000) {
    let millones = Math.floor(num / 1000000);
    if (millones === 1) texto += "un millón ";
    else texto += convertirGrupo(millones) + "millones ";
    num %= 1000000;
  }
  if (num >= 1000) {
    let miles = Math.floor(num / 1000);
    if (miles === 1) texto += "mil ";
    else texto += convertirGrupo(miles) + "mil ";
    num %= 1000;
  }
  if (num > 0) texto += convertirGrupo(num);
  return texto.trim();
}

function cleanTechNames(str) {
  if (!str) return str;
  let result = str;
  result = result.replace(/\biphone\b/gi, 'iPhone');
  result = result.replace(/\bipad\b/gi, 'iPad');
  result = result.replace(/\bios\b/gi, 'iOS');
  result = result.replace(/\bmacbook\b/gi, 'MacBook');
  result = result.replace(/\bmac\b(?!\s*book)/gi, 'Mac');
  result = result.replace(/\bairpods\b/gi, 'AirPods');
  result = result.replace(/\bapple\s+watch\b/gi, 'Apple Watch');
  return result;
}
function setPublicViewPermissions(doc) {
  try {
    const file = DriveApp.getFileById(doc.getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return doc.getUrl();
  }
}
function buildResponse(success, message, data) {
  // Retorna objeto plano compatible con google.script.run
  // (ContentService solo funciona con doGet/doPost, no con google.script.run)
  return { success, message, data: data || null };
}
function _getOrCreateSheet(name, headers) {
  const SHEET_ID = '1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function _getSheet(name) {
  const SHEET_ID = '1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name);
}
function auditData(data) {
  const keys = Object.keys(data);
  for (const key of keys) {
    if (String(data[key]).includes('undefined')) {
      throw new Error(`Dato faltante: ${key}`);
    }
  }
}
function _sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}  
