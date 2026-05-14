// =================================================================
// RAPPIMIND · SECURITY & VALIDATION UTILITIES
// -----------------------------------------------------------------
// Centralised helpers for input validation, sanitisation, escaping
// and access control. All public Apps Script entry points exposed
// to `google.script.run` should funnel user-supplied data through
// these helpers before touching Sheets / Drive / Docs APIs.
// =================================================================

/**
 * Regex used to detect characters that must be encoded before being
 * placed inside HTML attributes or text nodes. Mirrors the OWASP
 * recommended escaping set for HTML body context.
 */
var _HTML_ESCAPE_PATTERN = /[&<>"'`=\/]/g;
var _HTML_ESCAPE_MAP = {
  '&':  '&amp;',
  '<':  '&lt;',
  '>':  '&gt;',
  '"':  '&quot;',
  "'":  '&#39;',
  '`':  '&#96;',
  '=':  '&#61;',
  '/':  '&#47;'
};

/**
 * Escape a string for safe insertion into HTML.
 * Use this on backend strings that will be rendered as HTML on the
 * client (chat messages, log details, notifications).
 *
 * @param {*} value — Anything; non-strings are coerced via String().
 * @return {string} Escaped representation, empty string if null/undef.
 */
function htmlEscape(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(_HTML_ESCAPE_PATTERN, function(ch) {
    return _HTML_ESCAPE_MAP[ch];
  });
}

/**
 * Strip every HTML tag from a string and return the resulting text.
 * Useful when a free-text field (notes, descriptions) needs to be
 * stored as plain text in Sheets but the input may contain markup.
 */
function stripHtml(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/<[^>]*>/g, '').trim();
}

/**
 * Truncate a string to `max` characters. Defensive cap to keep cells
 * inside Sheets' 50k-character limit and to thwart oversized payloads.
 */
function safeTruncate(value, max) {
  var s = (value === null || value === undefined) ? '' : String(value);
  var limit = (typeof max === 'number' && max > 0) ? max : 5000;
  return s.length > limit ? s.substring(0, limit) : s;
}

/**
 * Trim whitespace and collapse runs of whitespace to single spaces.
 */
function normaliseWhitespace(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/\s+/g, ' ').trim();
}

// -----------------------------------------------------------------
// EMAIL & DOMAIN VALIDATION
// -----------------------------------------------------------------

var _EMAIL_PATTERN = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;

/**
 * Validate an email address. By default returns the lowercase email
 * if valid, throws otherwise. Pass `{ silent: true }` for boolean.
 *
 * @param {string} email
 * @param {{ silent?: boolean, requireDomain?: string }} [opts]
 * @return {string|boolean}
 */
function validateEmail(email, opts) {
  opts = opts || {};
  var clean = (email || '').toString().trim().toLowerCase();
  var ok = _EMAIL_PATTERN.test(clean) && clean.length <= 254;
  if (ok && opts.requireDomain) {
    ok = clean.endsWith('@' + opts.requireDomain.toLowerCase());
  }
  if (opts.silent) return ok;
  if (!ok) throw new Error('Email inválido: ' + email);
  return clean;
}

/**
 * Validate a list of email addresses separated by commas/semicolons.
 * Returns the array of normalised emails. Throws on the first invalid
 * one unless `silent` is true.
 */
function validateEmails(emailsStr, opts) {
  opts = opts || {};
  var parts = (emailsStr || '').toString()
    .split(/[,;]/)
    .map(function(s) { return s.trim(); })
    .filter(Boolean);
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var ok = validateEmail(parts[i], { silent: true, requireDomain: opts.requireDomain });
    if (!ok) {
      if (opts.silent) return null;
      throw new Error('Email inválido en la lista: ' + parts[i]);
    }
    out.push(parts[i].toLowerCase());
  }
  return out;
}

// -----------------------------------------------------------------
// COUNTRY / CAMPAIGN CODES
// -----------------------------------------------------------------

/**
 * Validate a country code against the supported list (TW_CONFIG.COUNTRY_FOLDERS).
 * Returns the uppercase code or throws.
 */
function validateCountryCode(code) {
  var clean = (code || '').toString().trim().toUpperCase();
  if (!clean) throw new Error('Country code requerido');
  if (typeof TW_CONFIG !== 'undefined' && TW_CONFIG.COUNTRY_FOLDERS
      && !TW_CONFIG.COUNTRY_FOLDERS[clean]) {
    throw new Error('Country code no soportado: ' + clean);
  }
  return clean;
}

/**
 * Validate that a string only contains "safe" characters allowed in
 * folder names / sheet names. Prevents path traversal & Sheets quirks.
 */
function safeName(value, max) {
  var s = (value || '').toString().trim();
  if (!s) return '';
  s = s.replace(/[\/\\:\*\?"<>\|\x00-\x1F]/g, ''); // chars forbidden by Drive/Sheets
  return safeTruncate(s, max || 200);
}

// -----------------------------------------------------------------
// URLS
// -----------------------------------------------------------------

/**
 * Validate a URL, optionally constraining the host. Returns the URL
 * or throws.
 */
function validateUrl(url, opts) {
  opts = opts || {};
  var s = (url || '').toString().trim();
  if (!s) {
    if (opts.silent) return '';
    throw new Error('URL requerida');
  }
  if (!/^https?:\/\//i.test(s)) {
    if (opts.silent) return '';
    throw new Error('URL debe iniciar con http(s)://');
  }
  if (opts.allowedHosts && opts.allowedHosts.length) {
    var ok = opts.allowedHosts.some(function(h) { return s.indexOf(h) >= 0; });
    if (!ok) {
      if (opts.silent) return '';
      throw new Error('URL no permitida: ' + s);
    }
  }
  return s;
}

/**
 * Extract a Google Doc ID from a URL or raw ID.
 * Throws if no plausible ID can be parsed (prevents IDOR via crafted URLs).
 */
function extractDocId(input) {
  var s = (input || '').toString().trim();
  if (!s) throw new Error('Doc URL/ID requerido');
  var m = s.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{25,}$/.test(s)) return s;
  throw new Error('Doc ID inválido en: ' + s);
}

// -----------------------------------------------------------------
// NUMBERS & DATES
// -----------------------------------------------------------------

/**
 * Parse a number from arbitrary input, optionally validating range.
 * Returns the parsed number or throws.
 */
function validateNumber(value, opts) {
  opts = opts || {};
  if (value === null || value === undefined || value === '') {
    if (opts.optional) return null;
    throw new Error('Número requerido' + (opts.label ? ': ' + opts.label : ''));
  }
  var n = Number(String(value).replace(/[,\s]/g, ''));
  if (isNaN(n)) throw new Error('Valor numérico inválido: ' + value);
  if (typeof opts.min === 'number' && n < opts.min) {
    throw new Error('Valor menor al mínimo (' + opts.min + '): ' + n);
  }
  if (typeof opts.max === 'number' && n > opts.max) {
    throw new Error('Valor mayor al máximo (' + opts.max + '): ' + n);
  }
  return n;
}

/**
 * Parse an ISO date string (yyyy-MM-dd) into a Date object.
 * Returns null on empty input if `optional`, throws otherwise.
 */
function validateIsoDate(value, opts) {
  opts = opts || {};
  if (!value) {
    if (opts.optional) return null;
    throw new Error('Fecha requerida');
  }
  var s = String(value).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Fecha inválida (formato yyyy-MM-dd): ' + s);
  var y = +m[1], mo = +m[2], d = +m[3];
  var dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) {
    throw new Error('Fecha calendaríca inválida: ' + s);
  }
  return dt;
}

// -----------------------------------------------------------------
// ACCESS CONTROL HELPERS
// -----------------------------------------------------------------

/**
 * Get the active user's email. Returns '' if Apps Script can't resolve
 * the session (e.g. unauthenticated access). Never throws.
 */
function getActiveUserEmailSafe() {
  try {
    return (Session.getActiveUser().getEmail() || '').toLowerCase();
  } catch (e) {
    return '';
  }
}

/**
 * Compare two emails case-insensitively, treating null/undefined as ''.
 */
function emailEquals(a, b) {
  return (a || '').toString().toLowerCase() === (b || '').toString().toLowerCase();
}

// -----------------------------------------------------------------
// PAYLOAD HELPERS
// -----------------------------------------------------------------

/**
 * Parse a JSON string and return the resulting object. Throws a
 * user-friendly error on malformed input — never leaks raw stack to
 * the front end.
 *
 * @param {string} jsonStr
 * @param {string} [label] — used in error messages.
 */
function parseJsonSafely(jsonStr, label) {
  if (jsonStr === null || jsonStr === undefined || jsonStr === '') return {};
  if (typeof jsonStr === 'object') return jsonStr; // already an object
  try {
    return JSON.parse(jsonStr);
  } catch (e) {
    throw new Error('Payload JSON inválido' + (label ? ' (' + label + ')' : '') + ': ' + e.message);
  }
}

/**
 * Coerce a value to a sanitised string for a Sheets cell:
 *   - strips HTML
 *   - normalises whitespace
 *   - truncates to a hard limit
 * Pass this around fields that come from the front-end and end up in
 * a Sheets cell to prevent CSV/formula injection.
 */
function sheetCellSafe(value, max) {
  var s = normaliseWhitespace(stripHtml(value));
  // Neutralise Sheets formula injection (a leading =, +, -, @ becomes literal text).
  if (s && /^[=+\-@]/.test(s)) s = "'" + s;
  return safeTruncate(s, max || 5000);
}

// -----------------------------------------------------------------
// PROPERTY LOOKUPS (lazy, cacheable)
// -----------------------------------------------------------------

var _PROP_CACHE = {};

/**
 * Lookup a script property with an optional fallback. Cached for the
 * duration of an execution to avoid repeated PropertiesService calls.
 *
 * @param {string} key
 * @param {string} [fallback]
 * @return {string}
 */
function getScriptProperty(key, fallback) {
  if (_PROP_CACHE.hasOwnProperty(key)) return _PROP_CACHE[key] || fallback || '';
  var v = '';
  try {
    v = PropertiesService.getScriptProperties().getProperty(key) || '';
  } catch (e) {
    v = '';
  }
  _PROP_CACHE[key] = v;
  return v || fallback || '';
}

/**
 * Convenience: resolve the audit Sheet ID from PropertiesService with
 * fallback to the legacy hard-coded constant. Lets ops rotate the ID
 * without redeploying.
 */
function resolveAuditSheetId() {
  var id = getScriptProperty('AUDIT_SHEET_ID', typeof AUDIT_SHEET_ID !== 'undefined' ? AUDIT_SHEET_ID : '');
  if (!id) throw new Error('AUDIT_SHEET_ID no está configurado en Script Properties.');
  return id;
}
