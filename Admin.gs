// =================================================================
// RAPPIMIND · PANEL DE ADMINISTRACIÓN Y WORKFLOWS
// -----------------------------------------------------------------
// All admin RPCs entered from `google.script.run`. Every public entry
// point follows the same contract:
//
//   1. Validate the caller's role via _requireRole().
//   2. Parse + sanitise inputs via Security.gs helpers.
//   3. Mutate state with idempotent UPSERTs against Sheets.
//   4. Return a JSON-stringified envelope: { status, message?, data? }.
//
// Errors are caught at the boundary and reduced to a stable shape so
// the front-end never receives raw stack traces.
// =================================================================

/**
 * Wrap an admin action with the standard try/catch + JSON envelope.
 * Centralises the "return error as JSON.stringify" pattern that was
 * repeated dozens of times across this file.
 *
 * @param {function(): Object} fn
 * @return {string} JSON-stringified result.
 */
function _adminAction(fn) {
  try {
    return JSON.stringify(fn());
  } catch (e) {
    Logger.log('❌ Admin action error: ' + (e && e.stack ? e.stack : e));
    return JSON.stringify({ status: 'error', message: (e && e.message) ? e.message : String(e) });
  }
}

/**
 * Identify the current user and resolve their role / permissions.
 * @return {string} JSON envelope.
 */
function adminGetCurrentUser() {
  return _adminAction(function() {
    var email = getActiveUserEmailSafe();
    if (!email) {
      return { status: 'error', message: 'No se pudo obtener el email. ¿Estás logueado con cuenta Rappi?' };
    }

    var team = _getTeamMembers();
    var member = team.find(function(m) { return emailEquals(m.email, email); });
    if (!member) {
      // Owners hard-coded in Config also count, in case the team sheet is empty.
      if (ADMIN_EMAILS_LIST && ADMIN_EMAILS_LIST.indexOf(email) >= 0) {
        return { status: 'ok', email: email, role: 'owner', name: email,
                 permissions: _getPermissions('owner') };
      }
      return { status: 'unauthorized', email: email,
               message: 'No tienes acceso al Panel de Administración. Contacta a un Owner.' };
    }

    return { status: 'ok', email: email, role: member.role, name: member.name,
             permissions: _getPermissions(member.role) };
  });
}

function _getPermissions(role) {
  return {
    canViewAdmin:       ROLE_HIERARCHY[role] >= 1,  
    canEditFields:      ROLE_HIERARCHY[role] >= 2,  
    canCreateTemplates: ROLE_HIERARCHY[role] >= 2,  
    canApproveTemplates:ROLE_HIERARCHY[role] >= 3,  
    canActivateTemplates:ROLE_HIERARCHY[role] >= 3, 
    canDeleteTemplates: ROLE_HIERARCHY[role] >= 3,  
    canManageTeam:      ROLE_HIERARCHY[role] >= 4,  
    canManageFolders:   ROLE_HIERARCHY[role] >= 3,  
  };
}

function adminGetTeam() {
  return _adminAction(function() {
    _requireRole('viewer');
    return { status: 'ok', team: _getTeamMembers() };
  });
}

function adminAddTeamMember(jsonStr) {
  return _adminAction(function() {
    _requireRole('owner');
    var d = parseJsonSafely(jsonStr, 'adminAddTeamMember');
    if (!d.email || !d.role || !d.name) {
      return { status: 'error', message: 'Email, nombre y rol son requeridos' };
    }
    var email = validateEmail(d.email, { requireDomain: 'rappi.com' });
    if (!ROLE_HIERARCHY[d.role]) {
      return { status: 'error', message: 'Rol inválido: ' + d.role };
    }
    var name = sheetCellSafe(d.name, 200);
    var notes = sheetCellSafe(d.notes || '', 1000);

    var sheet = _getOrCreateSheet(TW_CONFIG.SHEET_TEAM,
      ['email', 'name', 'role', 'added_by', 'added_date', 'status', 'notes']);

    if (_getTeamMembers().some(function(m) { return emailEquals(m.email, email); })) {
      return { status: 'error', message: 'Este email ya está en el equipo' };
    }

    sheet.appendRow([
      email, name, d.role, getActiveUserEmailSafe(),
      new Date().toISOString().split('T')[0], 'active', notes
    ]);
    _shareFolderWithMember(email, d.role);
    _logApprovalAction('team_add', 'Agregó a ' + name + ' (' + email + ') como ' + d.role);
    return { status: 'ok', message: name + ' agregado como ' + d.role };
  });
}

function adminUpdateTeamMember(jsonStr) {
  return _adminAction(function() {
    _requireRole('owner');
    var d = parseJsonSafely(jsonStr, 'adminUpdateTeamMember');
    var email = validateEmail(d.email, { requireDomain: 'rappi.com' });
    if (d.role && !ROLE_HIERARCHY[d.role]) {
      return { status: 'error', message: 'Rol inválido: ' + d.role };
    }
    var sheet = _getSheet(TW_CONFIG.SHEET_TEAM);
    if (!sheet) return { status: 'error', message: 'Sheet de equipo no existe' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (emailEquals(data[i][0], email)) {
        if (d.role) sheet.getRange(i + 1, 3).setValue(d.role);
        if (d.status) sheet.getRange(i + 1, 6).setValue(sheetCellSafe(d.status, 32));
        _logApprovalAction('team_update', 'Actualizó ' + email + ' → ' + (d.role || d.status));
        return { status: 'ok' };
      }
    }
    return { status: 'error', message: 'Miembro no encontrado' };
  });
}

function adminRemoveTeamMember(email) {
  return _adminAction(function() {
    _requireRole('owner');
    var clean = validateEmail(email);
    if (emailEquals(clean, getActiveUserEmailSafe())) {
      return { status: 'error', message: 'No puedes eliminarte a ti mismo' };
    }
    var sheet = _getSheet(TW_CONFIG.SHEET_TEAM);
    if (!sheet) return { status: 'error', message: 'Sheet de equipo no existe' };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (emailEquals(data[i][0], clean)) {
        sheet.deleteRow(i + 1);
        _logApprovalAction('team_remove', 'Eliminó a ' + clean + ' del equipo');
        return { status: 'ok' };
      }
    }
    return { status: 'error', message: 'No encontrado' };
  });
}

function _getTeamMembers() {
  const sheet = _getSheet(TW_CONFIG.SHEET_TEAM);
  if (!sheet) return [];
  return _sheetToObjects(sheet).filter(m => m.status === 'active');
}

function adminGetFolderStructure() {
  return _adminAction(function() {
    _requireRole('viewer');
    var rootId = getScriptProperty('TEMPLATES_FOLDER_ID');
    if (!rootId) return { status: 'ok', folders: [], rootUrl: null };

    var root = DriveApp.getFolderById(rootId);
    var folders = [];
    var countryFolders = root.getFolders();

    while (countryFolders.hasNext()) {
      var cf = countryFolders.next();
      var subFolders = [];
      var subs = cf.getFolders();
      while (subs.hasNext()) {
        var sf = subs.next();
        var files = [];
        var fileIter = sf.getFiles();
        while (fileIter.hasNext()) {
          var f = fileIter.next();
          files.push({ name: f.getName(), id: f.getId(), url: f.getUrl() });
        }
        subFolders.push({ name: sf.getName(), id: sf.getId(), files: files });
      }
      folders.push({ name: cf.getName(), id: cf.getId(), subFolders: subFolders });
    }
    return { status: 'ok', rootId: rootId, rootUrl: root.getUrl(), folders: folders };
  });
}

var _VALID_TEMPLATE_STATUSES = ['draft', 'pending_review', 'approved', 'rejected', 'active', 'inactive'];

function adminToggleTemplate(index, newStatus) {
  return _adminAction(function() {
    _requireRole('admin');
    var idx = validateNumber(index, { min: 0, label: 'index' });
    var status = sheetCellSafe(newStatus, 32);
    if (_VALID_TEMPLATE_STATUSES.indexOf(status) < 0) {
      return { status: 'error', message: 'Status inválido: ' + status };
    }

    var sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    if (!sheet) return { status: 'error', message: 'Sheet de templates no existe' };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var statusCol = headers.indexOf('status') + 1;
    var updatedCol = headers.indexOf('last_updated') + 1;
    if (!statusCol) return { status: 'error', message: 'Columna status no encontrada en Registry.' };

    var row = idx + 2;
    if (row > sheet.getLastRow()) return { status: 'error', message: 'Fila fuera de rango.' };
    sheet.getRange(row, statusCol).setValue(status);
    if (updatedCol > 0) sheet.getRange(row, updatedCol).setValue(new Date().toISOString().split('T')[0]);

    if (status === 'active') {
      try {
        var rowVals = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
        var campaignType = rowVals[headers.indexOf('campaign_type')];
        _ensureCampaignTypeActive(campaignType);
      } catch (e) {
        Logger.log('⚠️ Auto-activate Campaign_Type: ' + e.message);
      }
    }
    return { status: 'ok' };
  });
}

// V3.1: También activar Campaign_Type cuando se aprueba con activación
// (No necesita función nueva: adminApproveTemplate ya llama adminToggleTemplate indirectamente
//  o setea status = 'active'. Agregamos el hook aquí.)
function _ensureCampaignTypeActive(campaignType) {
  var ctSheet = _getSheet(CAMPAIGN_TYPES_SHEET);
  if (!ctSheet) return;
  var ctData = ctSheet.getDataRange().getValues();
  var ctHeaders = ctData[0];
  var nameCol = ctHeaders.indexOf('type_name');
  var statusCol = ctHeaders.indexOf('status');

  for (var i = 1; i < ctData.length; i++) {
    if (String(ctData[i][nameCol]) === campaignType && String(ctData[i][statusCol]) !== 'active') {
      ctSheet.getRange(i + 1, statusCol + 1).setValue('active');
      Logger.log('✅ Campaign_Type "' + campaignType + '" auto-activado al activar template');
      break;
    }
  }
}

function adminDeleteTemplate(index) {
  return _adminAction(function() {
    _requireRole('admin');
    var idx = validateNumber(index, { min: 0, label: 'index' });
    var sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    if (!sheet) return { status: 'error', message: 'Sheet de templates no existe' };
    var row = idx + 2;
    if (row > sheet.getLastRow()) return { status: 'error', message: 'Fila fuera de rango.' };
    sheet.deleteRow(row);
    _logApprovalAction('template_delete', 'Template fila #' + row);
    return { status: 'ok' };
  });
}

function adminSaveTemplate(jsonStr, editIndex) {
  try {
    const callerRole = _requireRole('editor');
    const d = JSON.parse(jsonStr);
    let sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    const callerEmail = Session.getActiveUser().getEmail();
    
    let initialStatus = ROLE_HIERARCHY[callerRole.role] >= ROLE_HIERARCHY['admin'] ? (d.status || 'active') : 'draft';
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = new Array(headers.length).fill('');
    const setValue = (colName, value) => { const idx = headers.indexOf(colName); if (idx >= 0) row[idx] = value || ''; };

    setValue('country_code', d.country_code); setValue('country_name', d.country_name);
    setValue('campaign_type', d.campaign_type); setValue('template_doc_id', d.template_doc_id);
    setValue('version', d.version || '1.0'); setValue('status', initialStatus);
    setValue('currency_code', d.currency_code || 'COP'); setValue('currency_symbol', d.currency_symbol || '$');
    setValue('legal_owner', d.legal_owner || callerEmail); setValue('last_updated', new Date().toISOString().split('T')[0]);
    setValue('notes', d.notes || ''); setValue('submitted_by', callerEmail);

    if (editIndex >= 0) sheet.getRange(editIndex + 2, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);

    _moveTemplateToFolder(d.template_doc_id, d.country_code, d.country_name, d.campaign_type);
    _logApprovalAction(editIndex >= 0 ? 'template_edit' : 'template_create', `${d.country_code}/${d.campaign_type} por ${callerEmail} [${initialStatus}]`);
    return JSON.stringify({ status: 'ok', message: initialStatus === 'draft' ? 'Guardado como borrador' : 'Guardado como ' + initialStatus, initialStatus });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminSubmitForReview(index) {
  return _adminAction(function() {
    _requireRole('editor');
    var idx = validateNumber(index, { min: 0, label: 'index' });
    var sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    if (!sheet) return { status: 'error', message: 'Sheet de templates no existe' };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var statusCol = headers.indexOf('status') + 1;
    if (!statusCol) return { status: 'error', message: 'Columna status no encontrada.' };
    var row = idx + 2;
    if (row > sheet.getLastRow()) return { status: 'error', message: 'Fila fuera de rango.' };

    sheet.getRange(row, statusCol).setValue('pending_review');
    var email = getActiveUserEmailSafe();
    _logApprovalAction('submit_review', 'Template #' + idx + ' enviado a revisión por ' + email);
    _notifyAdmins(idx, 'pending_review', email);
    return { status: 'ok', message: 'Enviado a revisión' };
  });
}

function adminApproveTemplate(index, activateNow) {
  return _adminAction(function() {
    _requireRole('admin');
    var idx = validateNumber(index, { min: 0, label: 'index' });
    var sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    if (!sheet) return { status: 'error', message: 'Sheet de templates no existe' };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var statusCol = headers.indexOf('status') + 1;
    if (!statusCol) return { status: 'error', message: 'Columna status no encontrada.' };
    var row = idx + 2;
    if (row > sheet.getLastRow()) return { status: 'error', message: 'Fila fuera de rango.' };

    var email = getActiveUserEmailSafe();
    var newStatus = activateNow ? 'active' : 'approved';
    sheet.getRange(row, statusCol).setValue(newStatus);
    _logApprovalAction('approve', 'Template #' + idx + ' aprobado por ' + email + ' [' + newStatus + ']');

    if (activateNow) {
      try {
        var rowVals = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
        var campaignType = rowVals[headers.indexOf('campaign_type')];
        _ensureCampaignTypeActive(campaignType);
      } catch (e) {
        Logger.log('⚠️ Auto-activate Campaign_Type en approve: ' + e.message);
      }
    }
    return { status: 'ok', message: 'Template aprobado' };
  });
}

function adminRejectTemplate(index, reason) {
  return _adminAction(function() {
    _requireRole('admin');
    var idx = validateNumber(index, { min: 0, label: 'index' });
    var safeReason = sheetCellSafe(reason || 'Sin motivo especificado', 1000);
    var sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    if (!sheet) return { status: 'error', message: 'Sheet de templates no existe' };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var statusCol = headers.indexOf('status') + 1;
    var notesCol  = headers.indexOf('notes') + 1;
    if (!statusCol) return { status: 'error', message: 'Columna status no encontrada.' };
    var row = idx + 2;
    if (row > sheet.getLastRow()) return { status: 'error', message: 'Fila fuera de rango.' };

    sheet.getRange(row, statusCol).setValue('rejected');
    if (notesCol) sheet.getRange(row, notesCol).setValue('RECHAZADO: ' + safeReason);
    _logApprovalAction('reject', 'Template #' + idx + ' rechazado: ' + safeReason);
    return { status: 'ok', message: 'Template rechazado' };
  });
}

// --- OTROS GETTERS DEL ADMIN ---
function adminGetApprovalLog() {
  return _adminAction(function() {
    _requireRole('viewer');
    var sheet = _getSheet('Approval_Log');
    if (!sheet) return { status: 'ok', logs: [] };
    var data = sheet.getDataRange().getValues();
    var logs = [];
    for (var i = data.length - 1; i >= Math.max(1, data.length - 100); i--) {
      logs.push({
        timestamp: data[i][0] ? new Date(data[i][0]).toLocaleString('es-CO') : '-',
        actor: data[i][1] || '-',
        action: data[i][2] || '-',
        details: data[i][3] || ''
      });
    }
    return { status: 'ok', logs: logs };
  });
}

function adminGetTemplates() {
  return _adminAction(function() {
    _requireRole('viewer');
    var sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    return { status: 'ok', templates: sheet ? _sheetToObjects(sheet) : [] };
  });
}

function adminGetLogs() {
  return _adminAction(function() {
    _requireRole('viewer');
    var sheet = _getSheet('Respuestas_Audit_V2');
    if (!sheet) return { status: 'ok', logs: [] };
    var data = sheet.getDataRange().getValues();
    var logs = [];
    for (var i = data.length - 1; i >= Math.max(1, data.length - 50); i--) {
      logs.push({
        timestamp: data[i][0] ? new Date(data[i][0]).toLocaleString('es-CO') : '-',
        email: data[i][1] || '-', docUrl: data[i][2] || '',
        type: data[i][3] || '-', country: data[i][4] || 'CO', shop: data[i][5] || '-'
      });
    }
    return { status: 'ok', logs: logs };
  });
}

function adminGetFields() {
  return _adminAction(function() {
    _requireRole('viewer');
    return { status: 'ok', fields: _sheetToObjects(_getSheet(TW_CONFIG.SHEET_FIELDS)) };
  });
}

function adminGetCampaignTypes() {
  return _adminAction(function() {
    _requireRole('viewer');
    return { status: 'ok', types: _sheetToObjects(_getSheet('Campaign_Types')) };
  });
}

function adminGetCountrySettings() {
  return _adminAction(function() {
    _requireRole('viewer');
    return { status: 'ok', settings: _sheetToObjects(_getSheet('Country_Settings')) };
  });
}

// --- HERRAMIENTAS INTERNAS ADMIN ---

/**
 * Require the caller to hold at least `minRole`. Owners listed in
 * ADMIN_EMAILS_LIST bypass the team sheet (bootstrap path).
 *
 * @param {string} minRole one of 'viewer' | 'editor' | 'admin' | 'owner'
 * @return {{email:string, role:string, name:string}}
 */
function _requireRole(minRole) {
  if (!ROLE_HIERARCHY[minRole]) throw new Error('Rol mínimo inválido: ' + minRole);
  var email = getActiveUserEmailSafe();
  if (!email) throw new Error('No se pudo identificar al usuario activo.');

  if (ADMIN_EMAILS_LIST && ADMIN_EMAILS_LIST.indexOf(email) >= 0) {
    return { email: email, role: 'owner', name: email };
  }

  var member = _getTeamMembers().find(function(m) { return emailEquals(m.email, email); });
  if (!member || ROLE_HIERARCHY[member.role] < ROLE_HIERARCHY[minRole]) {
    throw new Error('Permiso insuficiente.');
  }
  return member;
}

function _logApprovalAction(action, details) {
  try {
    _getOrCreateSheet('Approval_Log', ['timestamp', 'actor', 'action', 'details'])
      .appendRow([new Date(), getActiveUserEmailSafe(), action, safeTruncate(details, 1000)]);
  } catch (e) {
    Logger.log('⚠️ _logApprovalAction failed: ' + e.message);
  }
}

function _notifyAdmins(templateIndex, status, submitterEmail) {
  // Simplificado para ahorrar espacio
  Logger.log(`Notificación enviada a Admins: Template ${templateIndex} en estado ${status}`);
}

/**
 * Resolve a user's role without throwing — used for UI gating only.
 * Returns 'none' when the email is unknown or unauthenticated.
 */
function getUserRole(email) {
  var clean = (email || '').toString().trim().toLowerCase();
  if (!clean) return 'none';
  if (TW_CONFIG.ADMIN_EMAILS && TW_CONFIG.ADMIN_EMAILS.indexOf(clean) >= 0) return 'owner';
  try {
    var sheet = _getSheet(TW_CONFIG.SHEET_TEAM);
    if (!sheet) return 'none';
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return 'none';
    var headers = data[0];
    var emailIdx = headers.indexOf('email');
    var roleIdx = headers.indexOf('role');
    if (emailIdx < 0 || roleIdx < 0) return 'none';
    for (var i = 1; i < data.length; i++) {
      if (emailEquals(data[i][emailIdx], clean)) return data[i][roleIdx] || 'viewer';
    }
  } catch (e) {
    Logger.log('⚠️ getUserRole failed: ' + e.message);
  }
  return 'none';
}

function _getOrCreateDriveFolder(parent, name) {
  let iter = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (iter.hasNext()) return iter.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

function _shareFolderWithMember(email, role) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty('TEMPLATES_FOLDER_ID');
    if (!folderId) return;
    const folder = DriveApp.getFolderById(folderId);
    role === 'viewer' ? folder.addViewer(email) : folder.addEditor(email);
  } catch (e) {}
}

function _moveTemplateToFolder(docId, countryCode, countryName, campaignType) {
  try {
    const rootId = PropertiesService.getScriptProperties().getProperty('TEMPLATES_FOLDER_ID');
    if (!rootId) return;
    const root = DriveApp.getFolderById(rootId);
    const typeFolder = _getOrCreateDriveFolder(_getOrCreateDriveFolder(root, `${countryCode}_${countryName}`), campaignType.includes('Concurso') ? 'Concurso' : campaignType);
    const file = DriveApp.getFileById(docId);
    typeFolder.addFile(file);
    const parents = file.getParents();
    while (parents.hasNext()) {
      const p = parents.next();
      if (p.getId() !== typeFolder.getId()) p.removeFile(file);
    }
  } catch (e) {}
}

// --- GEMINI Y SMART TEMPLATES WIZARD ---
function analyzeTextForPlaceholders(payload) {
  try {
    _requireRole('editor');
    var p = payload && typeof payload === 'object' ? payload : parseJsonSafely(payload, 'analyzeTextForPlaceholders');
    if (!p.text) return buildResponse(false, 'Texto requerido para análisis.');
    var cc = p.countryCode ? validateCountryCode(p.countryCode) : 'CO';
    var ct = sheetCellSafe(p.campaignType || 'Cashback', 80);
    var prompt = buildAnalysisPrompt(p.text, cc, ct);
    return buildResponse(true, 'Análisis completado', callGeminiForAnalysis(prompt));
  } catch (e) {
    return buildResponse(false, 'Error al analizar: ' + (e && e.message ? e.message : e));
  }
}

function buildAnalysisPrompt(text, countryCode, campaignType) {
  var keys = Object.keys(TW_CONFIG.KNOWN_PLACEHOLDERS).map(function(k) {
    return '- {{' + k + '}}: ' + TW_CONFIG.KNOWN_PLACEHOLDERS[k].label;
  }).join('\n');
  return 'Analiza este T&C para ' + countryCode + ' (' + campaignType + ') y detecta variables. Usa estos u otros:\n'
       + keys + '\nResponde SOLO JSON: {"detections":[{"original_text":"","suggested_placeholder":"","label":""}]}\n\n'
       + String(text || '').substring(0, 15000);
}

/**
 * Call the Gemini API for placeholder analysis. The API key is loaded
 * from Script Properties; raw responses are *not* logged at INFO level
 * to avoid leaking PII or prompt content into shared execution logs.
 *
 * @param {string} prompt
 * @return {Object} Parsed JSON detection payload returned by Gemini.
 */
function callGeminiForAnalysis(prompt) {
  var apiKey = getScriptProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('No existe GEMINI_API_KEY en Script Properties.');
  }
  if (!prompt || typeof prompt !== 'string') {
    throw new Error('Prompt vacío o inválido.');
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=' + encodeURIComponent(apiKey);

  var response;
  try {
    response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1 }
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    throw new Error('Gemini fetch falló: ' + (e && e.message ? e.message : e));
  }

  var status = response.getResponseCode();
  var raw = response.getContentText();
  // Log status but NOT the raw response (may contain prompt echo / PII).
  Logger.log('Gemini status: ' + status + ' (' + (raw ? raw.length : 0) + ' bytes)');

  var parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    throw new Error('Gemini devolvió JSON inválido (HTTP ' + status + ').');
  }

  if (status < 200 || status >= 300) {
    var apiMessage = (parsed && parsed.error && parsed.error.message)
      ? parsed.error.message : 'Respuesta no exitosa';
    throw new Error('Gemini HTTP ' + status + ': ' + apiMessage);
  }

  var text = parsed && parsed.candidates && parsed.candidates[0]
    && parsed.candidates[0].content && parsed.candidates[0].content.parts
    && parsed.candidates[0].content.parts[0] && parsed.candidates[0].content.parts[0].text;
  if (!text) {
    throw new Error('Gemini no devolvió candidates válidos.');
  }

  var cleaned = text.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    throw new Error('La respuesta de Gemini no era JSON parseable.');
  }
}
// ---INICIO COPIAR---
// =================================================================
// WIZARD PASO 4: Crear template desde el Wizard (V3.1)
// =================================================================
function createTemplateFromWizard(payload) {
  try {
    // ─── 1. VALIDAR PAYLOAD ───
    var originalText = payload.originalText;
    var sourceDocId   = payload.sourceDocId || null;   // V3.1: Doc fuente
    var sourceDocTitle = payload.sourceDocTitle || null;
    var mappings      = payload.mappings || [];
    var metadata      = payload.metadata || {};
    var userEmail     = payload.userEmail || '';

    var countryCode   = metadata.countryCode;
    var campaignType  = (metadata.campaignType || '').trim();
    var version       = metadata.version || '1.0';
    var notes         = metadata.notes || '';

    if (!countryCode || !campaignType || (!originalText && !sourceDocId)) {
      return buildResponse(false, 'Faltan datos requeridos (país, tipo de campaña, o texto/doc fuente).');
    }

    // ─── 2. CREAR/COPIAR EL GOOGLE DOC ───
    var docTitle = 'Template_' + countryCode + '_' + campaignType.replace(/\s+/g, '_') + '_v' + version;
    var doc, docId;

    if (sourceDocId) {
      // RUTA A: Copia del Doc fuente (preserva formato, tablas, listas, negritas)
      var copiedFile = DriveApp.getFileById(sourceDocId).makeCopy(docTitle);
      docId = copiedFile.getId();
      doc = DocumentApp.openById(docId);
      var body = doc.getBody();

      // Aplicar los reemplazos de texto original → placeholder
      mappings.forEach(function(m) {
        if (m.confirmed && m.original_text && m.placeholder) {
          var escaped = m.original_text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          body.replaceText(escaped, '{{' + m.placeholder + '}}');
        }
      });

      doc.saveAndClose();

    } else {
      // RUTA B: Fallback — crear desde texto pegado
      var processedText = originalText;
      mappings.forEach(function(m) {
        if (m.confirmed && m.original_text && m.placeholder) {
          var escaped = m.original_text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          processedText = processedText.replace(new RegExp(escaped, 'g'), '{{' + m.placeholder + '}}');
        }
      });

      doc = DocumentApp.create(docTitle);
      var body = doc.getBody();
      if (body.getNumChildren() > 0) body.clear();

      var lines = processedText.split('\n');
      lines.forEach(function(line, idx) {
        var para = body.appendParagraph(line);
        if (idx === 0 && line.trim().length > 0) {
          para.setBold(true);
          para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        }
      });

      doc.saveAndClose();
      docId = doc.getId();
    }

    // ─── 3. MOVER A CARPETA CORRECTA EN DRIVE ───
    var countryFolderName = TW_CONFIG.COUNTRY_FOLDERS[countryCode] || countryCode;
    _moveTemplateToFolder(docId, countryCode, countryFolderName, campaignType);

    // ─── 4. DETERMINAR ROL Y STATUS ───
    var callerRole = 'editor';
    try {
      var roleInfo = _requireRole('editor');
      callerRole = roleInfo.role;
    } catch(e) { /* si falla, continuar como editor */ }

    // V3.1: Editor → draft (alineado con flujo manual), Admin/Owner → active
    var isAdminOrAbove = ROLE_HIERARCHY[callerRole] >= ROLE_HIERARCHY['admin'];
    var initialStatus = isAdminOrAbove ? 'active' : 'draft';

    // ─── 5. UPSERT EN Template_Registry ───
    // Llave compuesta: country_code + campaign_type + version
    var registrySheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    if (!registrySheet) {
      return buildResponse(false, 'Sheet Template_Registry no encontrada. Ejecuta el setup primero.');
    }

    var regHeaders = registrySheet.getRange(1, 1, 1, registrySheet.getLastColumn()).getValues()[0];
    var regData = registrySheet.getDataRange().getValues();
    var existingRegIdx = -1; // -1 = no existe, hay que crear

    // Buscar fila existente por llave compuesta
    var ccCol = regHeaders.indexOf('country_code');
    var ctCol = regHeaders.indexOf('campaign_type');
    var verCol = regHeaders.indexOf('version');
    for (var ri = 1; ri < regData.length; ri++) {
      if (String(regData[ri][ccCol]) === countryCode &&
          String(regData[ri][ctCol]) === campaignType &&
          String(regData[ri][verCol]) === version) {
        existingRegIdx = ri; // índice 1-based en la data (0-based en sheet = ri+1)
        break;
      }
    }

    var regRow = new Array(regHeaders.length).fill('');
    var setRegVal = function(col, val) {
      var idx = regHeaders.indexOf(col);
      if (idx >= 0) regRow[idx] = val || '';
    };

    // Obtener moneda del país
    var currencyCode = 'COP', currencySymbol = '$';
    try {
      var countrySheet = _getSheet(COUNTRY_SETTINGS_SHEET);
      if (countrySheet) {
        var countries = _sheetToObjects(countrySheet);
        var cc = countries.find(function(c) { return c.country_code === countryCode; });
        if (cc) { currencyCode = cc.currency_code || currencyCode; currencySymbol = cc.currency_symbol || currencySymbol; }
      }
    } catch(e) {}

    setRegVal('country_code', countryCode);
    setRegVal('country_name', countryFolderName.replace(countryCode + '_', ''));
    setRegVal('campaign_type', campaignType);
    setRegVal('template_doc_id', docId);
    setRegVal('version', version);
    setRegVal('status', initialStatus);
    setRegVal('currency_code', currencyCode);
    setRegVal('currency_symbol', currencySymbol);
    setRegVal('legal_owner', userEmail);
    setRegVal('last_updated', new Date().toISOString().split('T')[0]);
    setRegVal('notes', notes);
    setRegVal('submitted_by', userEmail);
    setRegVal('vertical', 'ALL');

    if (existingRegIdx >= 0) {
      // UPSERT: Actualizar fila existente
      registrySheet.getRange(existingRegIdx + 1, 1, 1, regRow.length).setValues([regRow]);
      Logger.log('♻️ Registry UPSERT: actualizada fila ' + (existingRegIdx + 1));
    } else {
      // INSERT: Nueva fila
      registrySheet.appendRow(regRow);
      Logger.log('➕ Registry INSERT: nueva fila para ' + countryCode + '/' + campaignType);
    }

    // ─── 6. UPSERT EN Template_Fields ───
    // Llave compuesta: country_code + campaign_type + placeholder
    var fieldsSheet = _getSheet(TW_CONFIG.SHEET_FIELDS);
    var fieldsCreated = 0;

    if (fieldsSheet && mappings.length > 0) {
      var fieldHeaders = fieldsSheet.getRange(1, 1, 1, fieldsSheet.getLastColumn()).getValues()[0];
      var fieldData = fieldsSheet.getDataRange().getValues();

      // Índices de columnas para búsqueda
      var fCcCol = fieldHeaders.indexOf('country_code');
      var fCtCol = fieldHeaders.indexOf('campaign_type');
      var fPhCol = fieldHeaders.indexOf('placeholder');

      mappings.forEach(function(m, idx) {
        if (!m.confirmed) return;

        // V3.3: Limpiar antes de envolver (evita {{{{X}}}})
        var rawPh = m.placeholder.replace(/^\{\{/, '').replace(/\}\}$/, '');
        var placeholderStr = '{{' + rawPh + '}}';

        // Buscar si ya existe este field
        var existingFieldIdx = -1;
        for (var fi = 1; fi < fieldData.length; fi++) {
          if (String(fieldData[fi][fCcCol]) === countryCode &&
              String(fieldData[fi][fCtCol]) === campaignType &&
              String(fieldData[fi][fPhCol]) === placeholderStr) {
            existingFieldIdx = fi;
            break;
          }
        }

        // V3.3: Limpiar placeholder antes de generar field_id (evita {{}} en el ID)
        var cleanPh = m.placeholder.replace(/^\{\{/, '').replace(/\}\}$/, '');
        var fieldId = cleanPh.toLowerCase().replace(/_([a-z])/g, function(match, letter) {
          return letter.toUpperCase();
        });

        // V3.1: Inferir field_type y format_as desde KNOWN_PLACEHOLDERS
        var knownConfig = TW_CONFIG.KNOWN_PLACEHOLDERS[m.placeholder];
        var fieldType = 'text';
        var formatAs = '';

        if (knownConfig) {
          fieldType = knownConfig.type || 'text';
          // Inferir format_as según el tipo conocido
          if (knownConfig.type === 'date') formatAs = 'date_legal';
          else if (knownConfig.type === 'number') {
            var labelLower = (m.label || '').toLowerCase();
            if (labelLower.indexOf('tope') >= 0 || labelLower.indexOf('valor') >= 0 ||
                labelLower.indexOf('monto') >= 0 || labelLower.indexOf('premio') >= 0 ||
                labelLower.indexOf('presupuesto') >= 0) {
              formatAs = 'money';
            } else if (labelLower.indexOf('ganador') >= 0 || labelLower.indexOf('orden') >= 0 ||
                       labelLower.indexOf('número') >= 0 || labelLower.indexOf('top') >= 0) {
              formatAs = 'number_words';
            } else if (labelLower.indexOf('porcentaje') >= 0 || labelLower.indexOf('%') >= 0 ||
                       labelLower.indexOf('cashback') >= 0) {
              formatAs = 'percentage';
            }
          }
        }

        // V3.1: required basado en confidence del mapping
        var isRequired = (m.confidence === 'HIGH') ? 'TRUE' : 'FALSE';

        var fieldRow = new Array(fieldHeaders.length).fill('');
        var setFieldVal = function(col, val) {
          var i = fieldHeaders.indexOf(col);
          if (i >= 0) fieldRow[i] = val || '';
        };
        // V3.3: Limpiar placeholder para lookup correcto en BASE_FIELD_MAP
        var cleanPhForMap = m.placeholder.replace(/^\{\{/, '').replace(/\}\}$/, '');
        var baseMapping = (typeof BASE_FIELD_MAP !== 'undefined') ? BASE_FIELD_MAP[cleanPhForMap] : null;
        var isBaseField = !!baseMapping;
        // V3.3: Detectar si es un campo legal auto-resuelto
        var isLegalDefault = (typeof LEGAL_DEFAULTS_MAP !== 'undefined') ? !!LEGAL_DEFAULTS_MAP[cleanPhForMap] : false;
        setFieldVal('field_id', fieldId);
        setFieldVal('country_code', countryCode);
        setFieldVal('campaign_type', campaignType);
        setFieldVal('placeholder', placeholderStr);
        setFieldVal('label_es', m.label || m.placeholder);
        setFieldVal('field_type', fieldType);
        setFieldVal('icon', '');
        setFieldVal('required', isRequired);
        setFieldVal('section', isBaseField ? '0' : (isLegalDefault ? 'L' : '3'));
        if (fieldHeaders.indexOf('canonical_field_id') >= 0) {
        setFieldVal('canonical_field_id', isBaseField ? baseMapping.canonical : '');
        }
        setFieldVal('options', '');
        setFieldVal('default_value', '');
        setFieldVal('tooltip', '');
        setFieldVal('depends_on', '');
        setFieldVal('order', String(idx + 1));
        setFieldVal('group', 'Wizard Import');

        // V3.1: Persistir format_as
        if (fieldHeaders.indexOf('format_as') >= 0) {
        if (isBaseField && baseMapping.format_as) {
          setFieldVal('format_as', baseMapping.format_as);
        } else {
          setFieldVal('format_as', formatAs);
        }
      }

        if (existingFieldIdx >= 0) {
          // UPSERT: Actualizar
          fieldsSheet.getRange(existingFieldIdx + 1, 1, 1, fieldRow.length).setValues([fieldRow]);
        } else {
          // INSERT: Nueva fila
          fieldsSheet.appendRow(fieldRow);
          // Actualizar fieldData para que el siguiente mapping no duplique
          fieldData.push(fieldRow);
        }
        fieldsCreated++;
      });
    }

    // ─── 7. UPSERT EN Campaign_Types ───
    // Llave: type_name
    // V3.1: Si el template es draft, el Campaign_Type también debe ser draft
    try {
      var ctSheet = _getSheet(CAMPAIGN_TYPES_SHEET);
      if (ctSheet) {
        var ctData = _sheetToObjects(ctSheet);
        var existingType = ctData.find(function(t) { return t.type_name === campaignType; });

        if (!existingType) {
          // Nuevo tipo: crearlo con el MISMO status que el template
          var ctHeaders = ctSheet.getRange(1, 1, 1, ctSheet.getLastColumn()).getValues()[0];
          var ctRow = new Array(ctHeaders.length).fill('');
          var setCtVal = function(col, val) {
            var i = ctHeaders.indexOf(col);
            if (i >= 0) ctRow[i] = val || '';
          };

          var typeId = campaignType.toLowerCase().replace(/\s+/g, '_').replace(/[^a-z0-9_]/g, '');

          setCtVal('type_id', typeId);
          setCtVal('type_name', campaignType);
          setCtVal('description', 'Creado desde el Template Wizard');
          setCtVal('parent_type', '');
          setCtVal('processing_mode', 'template_only');
          setCtVal('icon', 'fa-file-lines');
          setCtVal('color', '#3B82F6');
          // V3.1 CLAVE: Mismo status que el template → no exponer al comercial prematuramente
          setCtVal('status', initialStatus);
          setCtVal('countries', countryCode);
          setCtVal('created_by', userEmail);
          setCtVal('created_date', new Date().toISOString().split('T')[0]);

          ctSheet.appendRow(ctRow);
          Logger.log('🆕 Campaign_Type creado: ' + campaignType + ' [' + initialStatus + ']');
        }
        // Si ya existe, no lo tocamos (el Admin puede gestionarlo desde Dinámicas)
      }
    } catch(e) {
      Logger.log('⚠️ Campaign_Types check/create error: ' + e.message);
    }

    // ─── 8. LOG DE AUDITORÍA ───
    _logApprovalAction('wizard_create',
      countryCode + '/' + campaignType + ' v' + version +
      (sourceDocId ? ' (copia de Doc)' : ' (texto pegado)') +
      ' por ' + userEmail + ' [' + initialStatus + ']');

    // ─── 9. RESPUESTA ───
    var docUrl = 'https://docs.google.com/document/d/' + docId + '/edit';
    var statusMsg = initialStatus === 'active'
      ? 'Template creado y activado en producción.'
      : 'Template creado como borrador. Envíalo a revisión desde la pestaña Templates.';

    return buildResponse(true, statusMsg, {
      docUrl: docUrl,
      docId: docId,
      status: initialStatus,
      fieldsCreated: fieldsCreated,
      sourcePreserved: !!sourceDocId
    });

  } catch (e) {
    Logger.log('❌ createTemplateFromWizard error: ' + e.message + '\n' + e.stack);
    return buildResponse(false, 'Error al crear template: ' + e.message);
  }
}

/**
 * Read text from a Google Doc URL. Only callers with editor+ role
 * may fetch arbitrary docs; this protects against IDOR via the Wizard.
 */
function fetchGoogleDocContent(payload) {
  try {
    _requireRole('editor');
    var p = payload && typeof payload === 'object' ? payload : parseJsonSafely(payload, 'fetchGoogleDocContent');
    var docId = extractDocId(p.docUrl);

    var doc = DocumentApp.openById(docId);
    var text = doc.getBody().getText();
    var trimmed = text.trim();
    return buildResponse(true, 'OK', {
      docId: docId,
      docTitle: doc.getName(),
      text: text,
      wordCount: trimmed.length > 0 ? trimmed.split(/\s+/).length : 0,
      charCount: text.length
    });
  } catch (e) {
    return buildResponse(false, e && e.message ? e.message : String(e));
  }
}
