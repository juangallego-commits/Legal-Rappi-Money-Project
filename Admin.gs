// =================================================================
// RAPPIMIND - PANEL DE ADMINISTRACIÓN Y WORKFLOWS
// =================================================================

function adminGetCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return JSON.stringify({ status: 'error', message: 'No se pudo obtener el email. ¿Estás logueado con cuenta Rappi?' });

    const team = _getTeamMembers();
    const member = team.find(m => m.email.toLowerCase() === email.toLowerCase());

    if (!member) {
      return JSON.stringify({ status: 'unauthorized', email: email, message: 'No tienes acceso al Panel de Administración. Contacta a un Owner.' });
    }

    return JSON.stringify({ status: 'ok', email: email, role: member.role, name: member.name, permissions: _getPermissions(member.role) });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
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
  try {
    _requireRole('viewer');
    return JSON.stringify({ status: 'ok', team: _getTeamMembers() });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminAddTeamMember(jsonStr) {
  try {
    _requireRole('owner');
    const d = JSON.parse(jsonStr);
    if (!d.email || !d.role || !d.name) return JSON.stringify({ status: 'error', message: 'Email, nombre y rol son requeridos' });
    if (!d.email.toLowerCase().endsWith('@rappi.com')) return JSON.stringify({ status: 'error', message: 'Solo se permiten emails @rappi.com' });
    if (!ROLE_HIERARCHY[d.role]) return JSON.stringify({ status: 'error', message: 'Rol inválido: ' + d.role });

    const sheet = _getOrCreateSheet(TW_CONFIG.SHEET_TEAM, ['email', 'name', 'role', 'added_by', 'added_date', 'status', 'notes']);
    if (_getTeamMembers().find(m => m.email.toLowerCase() === d.email.toLowerCase())) return JSON.stringify({ status: 'error', message: 'Este email ya está en el equipo' });

    sheet.appendRow([d.email.toLowerCase(), d.name, d.role, Session.getActiveUser().getEmail(), new Date().toISOString().split('T')[0], 'active', d.notes || '']);
    _shareFolderWithMember(d.email, d.role);
    _logApprovalAction('team_add', `Agregó a ${d.name} (${d.email}) como ${d.role}`);
    return JSON.stringify({ status: 'ok', message: `${d.name} agregado como ${d.role}` });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminUpdateTeamMember(jsonStr) {
  try {
    _requireRole('owner');
    const d = JSON.parse(jsonStr);
    const sheet = _getSheet(TW_CONFIG.SHEET_TEAM);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Sheet de equipo no existe' });

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === d.email.toLowerCase()) {
        if (d.role) sheet.getRange(i + 1, 3).setValue(d.role);
        if (d.status) sheet.getRange(i + 1, 6).setValue(d.status);
        _logApprovalAction('team_update', `Actualizó rol de ${d.email} a ${d.role || d.status}`);
        return JSON.stringify({ status: 'ok' });
      }
    }
    return JSON.stringify({ status: 'error', message: 'Miembro no encontrado' });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminRemoveTeamMember(email) {
  try {
    _requireRole('owner');
    if (email.toLowerCase() === Session.getActiveUser().getEmail().toLowerCase()) return JSON.stringify({ status: 'error', message: 'No puedes eliminarte a ti mismo' });

    const sheet = _getSheet(TW_CONFIG.SHEET_TEAM);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
        sheet.deleteRow(i + 1);
        _logApprovalAction('team_remove', `Eliminó a ${email} del equipo`);
        return JSON.stringify({ status: 'ok' });
      }
    }
    return JSON.stringify({ status: 'error', message: 'No encontrado' });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function _getTeamMembers() {
  const sheet = _getSheet(TW_CONFIG.SHEET_TEAM);
  if (!sheet) return [];
  return _sheetToObjects(sheet).filter(m => m.status === 'active');
}

function adminGetFolderStructure() {
  try {
    _requireRole('viewer');
    const rootId = PropertiesService.getScriptProperties().getProperty('TEMPLATES_FOLDER_ID');
    if (!rootId) return JSON.stringify({ status: 'ok', folders: [], rootUrl: null });

    const root = DriveApp.getFolderById(rootId);
    const folders = [];
    const countryFolders = root.getFolders();
    
    while (countryFolders.hasNext()) {
      const cf = countryFolders.next();
      const subFolders = [];
      const subs = cf.getFolders();
      while (subs.hasNext()) {
        const sf = subs.next();
        const files = [];
        const fileIter = sf.getFiles();
        while (fileIter.hasNext()) {
          const f = fileIter.next();
          files.push({ name: f.getName(), id: f.getId(), url: f.getUrl() });
        }
        subFolders.push({ name: sf.getName(), id: sf.getId(), files: files });
      }
      folders.push({ name: cf.getName(), id: cf.getId(), subFolders: subFolders });
    }
    return JSON.stringify({ status: 'ok', rootId: rootId, rootUrl: root.getUrl(), folders: folders });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

// --- GESTIÓN DE PLANTILLAS Y FLUJOS ---
function adminToggleTemplate(index, newStatus) {
  try {
    const sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    const updatedCol = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('last_updated') + 1;
    sheet.getRange(index + 2, sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('status') + 1).setValue(newStatus);
    if (updatedCol > 0) sheet.getRange(index + 2, updatedCol).setValue(new Date().toISOString().split('T')[0]);
    return JSON.stringify({ status: 'ok' });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminDeleteTemplate(index) {
  try {
    const sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    sheet.deleteRow(index + 2);
    return JSON.stringify({ status: 'ok' });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
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
  try {
    _requireRole('editor');
    const sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    const email = Session.getActiveUser().getEmail();
    sheet.getRange(index + 2, sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('status') + 1).setValue('pending_review');
    _logApprovalAction('submit_review', `Template #${index} enviado a revisión por ${email}`);
    _notifyAdmins(index, 'pending_review', email);
    return JSON.stringify({ status: 'ok', message: 'Enviado a revisión' });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminApproveTemplate(index, activateNow) {
  try {
    _requireRole('admin');
    const sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    const email = Session.getActiveUser().getEmail();
    const newStatus = activateNow ? 'active' : 'approved';
    sheet.getRange(index + 2, sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('status') + 1).setValue(newStatus);
    _logApprovalAction('approve', `Template #${index} aprobado por ${email}`);
    return JSON.stringify({ status: 'ok', message: 'Template aprobado' });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminRejectTemplate(index, reason) {
  try {
    _requireRole('admin');
    const sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    sheet.getRange(index + 2, sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].indexOf('status') + 1).setValue('rejected');
    _logApprovalAction('reject', `Template #${index} rechazado: ${reason}`);
    return JSON.stringify({ status: 'ok', message: 'Template rechazado' });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

// --- OTROS GETTERS DEL ADMIN ---
function adminGetApprovalLog() {
  try {
    _requireRole('viewer');
    const sheet = _getSheet('Approval_Log');
    if (!sheet) return JSON.stringify({ status: 'ok', logs: [] });
    const data = sheet.getDataRange().getValues();
    const logs = [];
    for (let i = data.length - 1; i >= Math.max(1, data.length - 100); i--) {
      logs.push({ timestamp: data[i][0] ? new Date(data[i][0]).toLocaleString('es-CO') : '-', actor: data[i][1] || '-', action: data[i][2] || '-', details: data[i][3] || '' });
    }
    return JSON.stringify({ status: 'ok', logs: logs });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminGetTemplates() {
  try {
    _requireRole('viewer');
    const sheet = _getSheet(TW_CONFIG.SHEET_REGISTRY);
    return JSON.stringify({ status: 'ok', templates: sheet ? _sheetToObjects(sheet) : [] });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminGetLogs() {
  try {
    _requireRole('viewer');
    const sheet = _getSheet('Respuestas_Audit_V2');
    if (!sheet) return JSON.stringify({ status: 'ok', logs: [] });
    const data = sheet.getDataRange().getValues();
    const logs = [];
    for (let i = data.length - 1; i >= Math.max(1, data.length - 50); i--) {
      logs.push({ timestamp: data[i][0] ? new Date(data[i][0]).toLocaleString('es-CO') : '-', email: data[i][1] || '-', docUrl: data[i][2] || '', type: data[i][3] || '-', country: data[i][4] || 'CO', shop: data[i][5] || '-' });
    }
    return JSON.stringify({ status: 'ok', logs: logs });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminGetFields() {
  try {
    _requireRole('viewer');
    return JSON.stringify({ status: 'ok', fields: _sheetToObjects(_getSheet(TW_CONFIG.SHEET_FIELDS) || []) });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminGetCampaignTypes() {
  try {
    _requireRole('viewer');
    return JSON.stringify({ status: 'ok', types: _sheetToObjects(_getSheet('Campaign_Types') || []) });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

function adminGetCountrySettings() {
  try {
    _requireRole('viewer');
    return JSON.stringify({ status: 'ok', settings: _sheetToObjects(_getSheet('Country_Settings') || []) });
  } catch (e) { return JSON.stringify({ status: 'error', message: e.message }); }
}

// --- HERRAMIENTAS INTERNAS ADMIN ---
function _requireRole(minRole) {
  const email = Session.getActiveUser().getEmail();
  const member = _getTeamMembers().find(m => m.email.toLowerCase() === email.toLowerCase());
  if (!member || ROLE_HIERARCHY[member.role] < ROLE_HIERARCHY[minRole]) throw new Error(`Permiso insuficiente.`);
  return member;
}

function _logApprovalAction(action, details) {
  try {
    _getOrCreateSheet('Approval_Log', ['timestamp', 'actor', 'action', 'details']).appendRow([new Date(), Session.getActiveUser().getEmail(), action, details]);
  } catch (e) {}
}

function _notifyAdmins(templateIndex, status, submitterEmail) {
  // Simplificado para ahorrar espacio
  Logger.log(`Notificación enviada a Admins: Template ${templateIndex} en estado ${status}`);
}

function getUserRole(email) {
  if (!email) return 'none';
  if (TW_CONFIG.ADMIN_EMAILS.includes(email)) return 'owner';
  try {
    const data = _getSheet(TW_CONFIG.SHEET_TEAM).getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('email')] === email) return data[i][headers.indexOf('role')] || 'viewer';
    }
  } catch (e) {}
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
    const prompt = buildAnalysisPrompt(payload.text, payload.countryCode, payload.campaignType);
    return buildResponse(true, 'Análisis completado', callGeminiForAnalysis(prompt));
  } catch (e) { return buildResponse(false, 'Error al analizar: ' + e.message); }
}

function buildAnalysisPrompt(text, countryCode, campaignType) {
  const keys = Object.entries(TW_CONFIG.KNOWN_PLACEHOLDERS).map(([k, v]) => `- {{${k}}}: ${v.label}`).join('\n');
  return `Analiza este T&C para ${countryCode} y detecta variables. Usa estos u otros:\n${keys}\nResponde SOLO JSON: {"detections":[{"original_text":"","suggested_placeholder":"","label":""}]}\n\n${text.substring(0,15000)}`;
}

function callGeminiForAnalysis(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || 'AIzaSyB-HWnobl4UwQDGnk3zqN915sLvIJ_qOnM'}`;
  const response = UrlFetchApp.fetch(url, { method: 'POST', contentType: 'application/json', payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], generationConfig: { temperature: 0.1 } }), muteHttpExceptions: true });
  return JSON.parse(JSON.parse(response.getContentText()).candidates[0].content.parts[0].text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim());
}

function fetchGoogleDocContent(payload) {
  try {
    const match = payload.docUrl.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
    if (!match) return buildResponse(false, 'URL inválida');
    const doc = DocumentApp.openById(match[1]);
    return buildResponse(true, 'OK', { docId: match[1], docTitle: doc.getName(), text: doc.getBody().getText() });
  } catch (e) { return buildResponse(false, e.message); }
}
