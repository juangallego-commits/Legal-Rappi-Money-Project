// ════════════════════════════════════════════════════════════════
// Legal Team Tracker · Google Apps Script · Web App
// RappiPlus · Global Legal · v3.3 (Projects + Tasks)
// ----------------------------------------------------------------
// NOTE: This file backs a *separate* Apps Script deployment from the
// RappiMind T&C generator (Admin.gs / Setupg.gs / WebApp.Html). It
// reads its own spreadsheet (TRACKER_SHEET_ID) and renders the
// "Dashboard" HTML template. The two projects share this repo for
// convenience but do not share runtime state.
//
// Sheet layout (in the tracker spreadsheet):
//   • Tracking Activo  — active tasks (16 cols, see TASK_COLS)
//   • Historial        — completed tasks
//   • Config           — KV pairs starting at row 3
//   • Equipos          — teams: code, country, leader, leader_email,
//                        members CSV, emails CSV, slack, notes
//   • Proyectos        — projects (15 cols, see PROJ_COLS)
//
// All public entry points (doGet, getTrackerData, addTask, …) are
// callable via google.script.run from Dashboard.html.
// ════════════════════════════════════════════════════════════════

/** Default sheet ID — override via Script Properties → TRACKER_SHEET_ID. */
const TRACKER_DEFAULT_SHEET_ID = '19eR-pXzVLTSEdCADeBZ8fsd5x4f2t0GowUJiJm2X6ms';

const SHEET_ACTIVO    = 'Tracking Activo';
const SHEET_HISTORIAL = 'Historial';
const SHEET_CONFIG    = 'Config';
const SHEET_EQUIPOS   = 'Equipos';
const SHEET_PROYECTOS = 'Proyectos';

// Tasks: 16 cols — ID, Nombre, Resp, Acc, Deadline, Prioridad, Estado,
// Semana, Creado, Cerrado, Notas, ProyectoId, País, Líder, TipoTrabajo, Riesgo
const TASK_COLS = 16;
// Projects: 15 cols — ID, Nombre, País, Líder, Responsable, Deadline,
// Prioridad, Estado, Descripción, Notas, Creado, Semana, Participantes,
// TipoTrabajo, Riesgo
const PROJ_COLS = 15;

const STATUS_ORDER  = { 'Bloqueado': 0, 'En curso': 1, 'Pendiente': 2, 'En revisión': 3, 'Listo': 4 };
const PRIO_ORDER    = { 'Alta': 0, 'Media': 1, 'Baja': 2 };
const PROJ_STATUSES = ['Activo', 'En pausa', 'Completado', 'Cancelado'];

/** Resolve tracker sheet ID — script properties first, fallback to const. */
function _trackerSheetId() {
  try {
    var v = PropertiesService.getScriptProperties().getProperty('TRACKER_SHEET_ID');
    if (v) return v;
  } catch (e) { /* fall through */ }
  return TRACKER_DEFAULT_SHEET_ID;
}

/** Open the tracker spreadsheet, throwing a friendly error on failure. */
function _trackerSpreadsheet() {
  try {
    return SpreadsheetApp.openById(_trackerSheetId());
  } catch (e) {
    throw new Error('No se pudo abrir el tracker spreadsheet: ' + e.message);
  }
}

// ── WEB APP ─────────────────────────────────────────────────────
function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  if (page === 'api') {
    return ContentService
      .createTextOutput(JSON.stringify(getTrackerData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var html = HtmlService.createTemplateFromFile('Dashboard');
  html.data = JSON.stringify(getTrackerData());
  return html.evaluate()
    .setTitle('Legal Tracker · Rappi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(f) {
  return HtmlService.createHtmlOutputFromFile(f).getContent();
}

// ════════════════════════════════════════════════════════════════
// GET ALL DATA
// ════════════════════════════════════════════════════════════════

/** Aggregate the entire tracker payload sent to the front-end. */
function getTrackerData() {
  var ss = _trackerSpreadsheet();
  var activeTasks = readTasks(ss.getSheetByName(SHEET_ACTIVO));
  var histTasks   = readTasks(ss.getSheetByName(SHEET_HISTORIAL));
  var config      = readConfig(ss);
  var equipos     = readEquipos(ss);
  var projects    = readProjects(ss);

  // ── Enrich projects with task stats ─────────────────────────
  projects.forEach(function(p) {
    p.tasks = [];
    p.taskStats = { total: 0, pendiente: 0, enCurso: 0, enRevision: 0,
                    bloqueado: 0, listo: 0, alta: 0, media: 0, baja: 0 };
  });
  var projMap = {};
  projects.forEach(function(p) { projMap[p.id] = p; });

  function _bumpStats(s, t, isHistorial) {
    s.total++;
    if (isHistorial) { s.listo++; return; }
    if (t.status === 'Pendiente')   s.pendiente++;
    if (t.status === 'En curso')    s.enCurso++;
    if (t.status === 'En revisión') s.enRevision++;
    if (t.status === 'Bloqueado')   s.bloqueado++;
    if (t.status === 'Listo')       s.listo++;
    if (t.priority === 'Alta')  s.alta++;
    if (t.priority === 'Media') s.media++;
    if (t.priority === 'Baja')  s.baja++;
  }

  activeTasks.forEach(function(t) {
    var pid = t.proyectoId;
    if (pid && projMap[pid]) {
      projMap[pid].tasks.push(t);
      _bumpStats(projMap[pid].taskStats, t, false);
    }
  });
  // Count completed tasks from historial toward project stats
  histTasks.forEach(function(t) {
    var pid = t.proyectoId;
    if (pid && projMap[pid]) _bumpStats(projMap[pid].taskStats, t, true);
  });

  // Auto-calculate project status (unless manually forced)
  projects.forEach(function(p) {
    if (p.statusForced) return;
    var s = p.taskStats;
    if (s.total > 0) {
      if (s.listo === s.total) p.status = 'Completado';
      else if (s.bloqueado > 0 && s.enCurso === 0 && s.pendiente === 0 && s.enRevision === 0) p.status = 'En pausa';
      else p.status = 'Activo';
    }
    p.pctDone = s.total > 0 ? Math.round(s.listo / s.total * 100) : 0;
  });

  // ── KPIs ────────────────────────────────────────────────────
  var kpi = { total: activeTasks.length, alta: 0, media: 0, baja: 0,
              pendiente: 0, enCurso: 0, bloqueado: 0, enRevision: 0, listo: 0 };
  activeTasks.forEach(function(t) {
    if (t.priority === 'Alta')  kpi.alta++;
    if (t.priority === 'Media') kpi.media++;
    if (t.priority === 'Baja')  kpi.baja++;
    if (t.status === 'Pendiente')   kpi.pendiente++;
    if (t.status === 'En curso')    kpi.enCurso++;
    if (t.status === 'Bloqueado')   kpi.bloqueado++;
    if (t.status === 'En revisión') kpi.enRevision++;
    if (t.status === 'Listo')       kpi.listo++;
  });

  // ── Per-person stats ────────────────────────────────────────
  var allMembers = getAllMembers(equipos);
  var teamMap = {};
  function _newPersonRow() {
    return { total: 0, alta: 0, media: 0, baja: 0, pendiente: 0,
             enCurso: 0, bloqueado: 0, enRevision: 0, listo: 0 };
  }
  allMembers.forEach(function(n) { teamMap[n] = _newPersonRow(); });
  activeTasks.forEach(function(t) {
    if (!teamMap[t.resp]) teamMap[t.resp] = _newPersonRow();
    _bumpStats(teamMap[t.resp], t, false);
  });
  var team = Object.keys(teamMap).sort().map(function(name) {
    var p = teamMap[name];
    return {
      name: name,
      initials: name.split(' ').slice(0, 2)
        .map(function(w) { return (w[0] || '').toUpperCase(); }).join(''),
      country: getCountryForMember(name, equipos),
      total: p.total, alta: p.alta, media: p.media, baja: p.baja,
      pendiente: p.pendiente, enCurso: p.enCurso, bloqueado: p.bloqueado,
      enRevision: p.enRevision, listo: p.listo,
      pctDone: p.total > 0 ? Math.round(p.listo / p.total * 100) : 0
    };
  });

  // ── Per-country stats ───────────────────────────────────────
  var countryStats = {};
  equipos.forEach(function(eq) {
    countryStats[eq.code] = {
      code: eq.code, name: eq.country, leader: eq.leader,
      total: 0, alta: 0, media: 0, baja: 0
    };
  });
  activeTasks.forEach(function(t) {
    var cc = t.pais || getCountryForMember(t.resp, equipos);
    if (cc && countryStats[cc]) {
      var c = countryStats[cc];
      c.total++;
      if (t.priority === 'Alta')  c.alta++;
      if (t.priority === 'Media') c.media++;
      if (t.priority === 'Baja')  c.baja++;
    }
  });

  // ── SLA ─────────────────────────────────────────────────────
  var now = new Date();
  var slaData = { onTime: 0, atRisk: 0, overdue: 0 };
  var slaLimits = { Alta: 2, Media: 5, Baja: 7 };
  activeTasks.forEach(function(t) {
    if (t.status === 'Listo') return;
    if (!t.creadoRaw) { slaData.onTime++; return; }
    var bizDays = countBizDays(new Date(t.creadoRaw), now);
    var limit = slaLimits[t.priority] || 5;
    if (bizDays > limit) slaData.overdue++;
    else if (bizDays >= limit - 1) slaData.atRisk++;
    else slaData.onTime++;
  });

  // ── Project list for dropdowns (id + name) ──────────────────
  var projectList = projects
    .filter(function(p) { return p.status !== 'Completado' && p.status !== 'Cancelado'; })
    .map(function(p) { return { id: p.id, nombre: p.nombre }; });

  return {
    tasks: activeTasks, historial: histTasks,
    kpi: kpi, sla: slaData, team: team,
    countries: Object.keys(countryStats).map(function(k) { return countryStats[k]; }),
    equipos: equipos, projects: projects, projectList: projectList,
    semana: activeTasks.length > 0 ? activeTasks[0].semana : getCurrentWeekLabel(),
    generated: Utilities.formatDate(new Date(), 'America/Bogota', 'dd/MM/yyyy HH:mm'),
    config: config
  };
}

// ════════════════════════════════════════════════════════════════
// PROJECTS CRUD
// ════════════════════════════════════════════════════════════════
function readProjects(ss) {
  var ws = ss.getSheetByName(SHEET_PROYECTOS);
  if (!ws) return [];
  var lastRow = ws.getLastRow();
  if (lastRow < 2) return [];
  var lastCol = Math.min(ws.getLastColumn(), PROJ_COLS);
  var data = ws.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var projects = [];
  data.forEach(function(row) {
    if (!row[1]) return; // empty name → skip
    projects.push({
      id: row[0],
      nombre: row[1] || '',
      pais: (row[2] || '').toString().trim(),
      lider: (row[3] || '').toString().trim(),
      responsable: (row[4] || '').toString().trim(),
      deadline: row[5] ? (row[5] instanceof Date
        ? Utilities.formatDate(row[5], 'America/Bogota', 'dd/MM/yyyy')
        : row[5].toString()) : '',
      deadlineISO: row[5] && row[5] instanceof Date
        ? Utilities.formatDate(row[5], 'America/Bogota', 'yyyy-MM-dd') : '',
      priority: row[6] || 'Media',
      status: row[7] || 'Activo',
      statusForced: (row[7] || '').toString().trim() === 'Cancelado',
      descripcion: row[8] || '',
      notas: row[9] || '',
      creado: row[10] ? Utilities.formatDate(new Date(row[10]), 'America/Bogota', 'dd/MM/yyyy') : '',
      semana: row[11] || '',
      participantes: (row[12] || '').toString().split(',')
        .map(function(s) { return s.trim(); }).filter(Boolean),
      tipoTrabajo: (row[13] || '').toString().trim(),
      riesgo: (row[14] || '').toString().trim(),
      pctDone: 0, tasks: [], taskStats: {}
    });
  });
  return projects;
}

function addProject(obj) {
  if (!obj || !obj.nombre) return { success: false, error: 'Nombre de proyecto requerido' };
  var ss = _trackerSpreadsheet();
  var ws = ss.getSheetByName(SHEET_PROYECTOS);
  if (!ws) {
    ws = ss.insertSheet(SHEET_PROYECTOS);
    ws.appendRow(['ID', 'Nombre', 'País', 'Líder', 'Responsable', 'Deadline',
      'Prioridad', 'Estado', 'Descripción', 'Notas', 'Creado', 'Semana',
      'Participantes', 'TipoTrabajo', 'Riesgo']);
    ws.getRange(1, 1, 1, PROJ_COLS).setFontWeight('bold')
      .setBackground('#FF4940').setFontColor('#FFFFFF');
    ws.setTabColor('#FF4940');
  }
  var lastRow = ws.getLastRow();
  var newId = lastRow >= 2 ? ws.getRange(lastRow, 1).getValue() + 1 : 1;
  var equipos = readEquipos(ss);
  var pais  = obj.pais  || getCountryForMember(obj.responsable, equipos);
  var lider = obj.lider || getLeaderForCountry(pais, equipos);
  ws.appendRow([
    newId, obj.nombre, pais, lider, obj.responsable || '',
    obj.deadline || '', obj.priority || 'Media', obj.status || 'Activo',
    obj.descripcion || '', obj.notas || '', new Date(), getCurrentWeekLabel(),
    obj.participantes || '', obj.tipoTrabajo || '', obj.riesgo || ''
  ]);
  return { success: true, id: newId, nombre: obj.nombre };
}

var _PROJECT_FIELD_MAP = {
  'nombre': 2, 'pais': 3, 'lider': 4, 'responsable': 5,
  'deadline': 6, 'priority': 7, 'status': 8,
  'descripcion': 9, 'notas': 10, 'participantes': 13
};

function updateProjectField(projId, field, value) {
  var col = _PROJECT_FIELD_MAP[field];
  if (!col) return { success: false, error: 'Invalid field: ' + field };
  var ss = _trackerSpreadsheet();
  var ws = ss.getSheetByName(SHEET_PROYECTOS);
  if (!ws) return { success: false, error: 'No projects sheet' };
  var lastRow = ws.getLastRow();
  if (lastRow < 2) return { success: false, error: 'No projects' };

  var data = ws.getRange(2, 1, lastRow - 1, Math.min(ws.getLastColumn(), PROJ_COLS)).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == projId) { // intentional loose equality (string vs number IDs)
      ws.getRange(i + 2, col).setValue(value);
      return { success: true };
    }
  }
  return { success: false, error: 'Project #' + projId + ' not found' };
}

// ════════════════════════════════════════════════════════════════
// TASKS
// ════════════════════════════════════════════════════════════════
function readTasks(ws) {
  if (!ws) return [];
  var lastRow = ws.getLastRow();
  if (lastRow < 4) return [];
  var lastCol = Math.min(ws.getLastColumn(), TASK_COLS);
  var data = ws.getRange(4, 1, lastRow - 3, lastCol).getValues();
  var tasks = [];
  data.forEach(function(row) {
    if (!row[1]) return;
    var proyVal = (row[11] || '').toString().trim();
    tasks.push({
      id: row[0],
      nombre: row[1] || '',
      resp: row[2] || '',
      acc: row[3] || '',
      deadline: row[4] ? (row[4] instanceof Date
        ? Utilities.formatDate(row[4], 'America/Bogota', 'dd/MM/yyyy')
        : row[4].toString()) : '',
      deadlineISO: row[4] && row[4] instanceof Date
        ? Utilities.formatDate(row[4], 'America/Bogota', 'yyyy-MM-dd') : '',
      priority: row[5] || 'Media',
      status: row[6] || 'Pendiente',
      semana: row[7] || '',
      creado: row[8] ? Utilities.formatDate(new Date(row[8]), 'America/Bogota', 'dd/MM/yyyy') : '',
      creadoRaw: row[8] ? new Date(row[8]).toISOString() : null,
      cerrado: row[9] ? Utilities.formatDate(new Date(row[9]), 'America/Bogota', 'dd/MM/yyyy') : '',
      notas: row[10] || '',
      proyectoId: isNaN(parseInt(proyVal, 10)) ? '' : parseInt(proyVal, 10),
      proyecto: proyVal,
      pais: (row[12] || '').toString().trim(),
      lider: (row[13] || '').toString().trim(),
      tipoTrabajo: (row[14] || '').toString().trim(),
      riesgo: (row[15] || '').toString().trim()
    });
  });
  tasks.sort(function(a, b) {
    return (PRIO_ORDER[a.priority] || 1) - (PRIO_ORDER[b.priority] || 1)
        || (STATUS_ORDER[a.status]  || 2) - (STATUS_ORDER[b.status]  || 2);
  });
  return tasks;
}

function addTask(taskObj) {
  if (!taskObj || !taskObj.nombre) return { success: false, error: 'Nombre de tarea requerido' };
  var ss = _trackerSpreadsheet();
  var ws = ss.getSheetByName(SHEET_ACTIVO);
  if (!ws) return { success: false, error: 'Sheet ' + SHEET_ACTIVO + ' no encontrado' };
  var lastRow = ws.getLastRow();
  var newId = lastRow >= 4 ? ws.getRange(lastRow, 1).getValue() + 1 : 1;
  var equipos = readEquipos(ss);
  var pais  = taskObj.pais  || getCountryForMember(taskObj.resp, equipos);
  var lider = taskObj.lider || getLeaderForCountry(pais, equipos);
  ws.appendRow([
    newId, taskObj.nombre, taskObj.resp || '', taskObj.acc || '',
    taskObj.deadline || '', taskObj.priority || 'Media', taskObj.status || 'Pendiente',
    taskObj.semana || getCurrentWeekLabel(), new Date(), '', taskObj.notas || '',
    taskObj.proyectoId || taskObj.proyecto || '', pais, lider,
    taskObj.tipoTrabajo || '', taskObj.riesgo || ''
  ]);
  return { success: true, id: newId };
}

var _TASK_FIELD_MAP = {
  'nombre': 2, 'resp': 3, 'acc': 4, 'deadline': 5, 'priority': 6, 'status': 7,
  'notas': 11, 'proyecto': 12, 'proyectoId': 12, 'pais': 13, 'lider': 14,
  'tipoTrabajo': 15, 'riesgo': 16
};

function updateTaskField(taskId, field, value) {
  var col = _TASK_FIELD_MAP[field];
  if (!col) return { success: false, error: 'Invalid field: ' + field };
  var ss = _trackerSpreadsheet();
  var ws = ss.getSheetByName(SHEET_ACTIVO);
  if (!ws) return { success: false, error: 'Sheet ' + SHEET_ACTIVO + ' no encontrado' };
  var lastRow = ws.getLastRow();
  if (lastRow < 4) return { success: false, error: 'No tasks' };

  var lastCol = Math.min(ws.getLastColumn(), TASK_COLS);
  var data = ws.getRange(4, 1, lastRow - 3, lastCol).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == taskId) {
      var row = i + 4;
      ws.getRange(row, col).setValue(value);
      if (field === 'status' && value === 'Listo') {
        ws.getRange(row, 10).setValue(new Date());
        moveToHistorial(ss, ws, row);
        return { success: true, moved: true, message: 'Tarea movida a Historial' };
      }
      return { success: true };
    }
  }
  return { success: false, error: 'Task #' + taskId + ' not found' };
}

function updateTaskStatus(taskId, newStatus) {
  return updateTaskField(taskId, 'status', newStatus);
}

// ════════════════════════════════════════════════════════════════
// EQUIPOS / CONFIG / HELPERS
// ════════════════════════════════════════════════════════════════

function readEquipos(ss) {
  var ws = ss.getSheetByName(SHEET_EQUIPOS);
  if (!ws) return getDefaultEquipos();
  var lr = ws.getLastRow();
  if (lr < 2) return getDefaultEquipos();

  var data = ws.getRange(2, 1, lr - 1, 8).getValues();
  var eq = [];
  data.forEach(function(r) {
    var c = (r[0] || '').toString().trim();
    if (!c) return;
    eq.push({
      code: c,
      country: (r[1] || '').toString().trim(),
      leader: (r[2] || '').toString().trim().replace(/\n/g, ''),
      leaderEmail: (r[3] || '').toString().trim(),
      members: (r[4] || '').toString().split(',')
        .map(function(s) { return s.trim(); }).filter(Boolean),
      emails: (r[5] || '').toString().split(',')
        .map(function(s) { return s.trim(); }).filter(Boolean),
      slackChannel: (r[6] || '').toString().trim(),
      notes: (r[7] || '').toString().trim()
    });
  });
  return eq.length > 0 ? eq : getDefaultEquipos();
}

function getDefaultEquipos() {
  return [{
    code: 'CO', country: 'Colombia',
    leader: 'Carlos Eduardo Fernández', leaderEmail: '',
    members: ['Isabela Zuluaga', 'Nicolás Naranjo', 'Juan Manuel Caicedo',
              'Juan Camilo Gallego', 'Valeria Rangel', 'David Gaviria'],
    emails: [], slackChannel: '', notes: ''
  }];
}

function getAllMembers(eq) {
  var n = {};
  eq.forEach(function(e) {
    if (e.leader) n[e.leader] = 1;
    e.members.forEach(function(m) { n[m] = 1; });
  });
  return Object.keys(n).sort();
}

function getCountryForMember(name, eq) {
  if (!name) return '';
  for (var i = 0; i < eq.length; i++) {
    if (eq[i].leader === name) return eq[i].code;
    if (eq[i].members.indexOf(name) >= 0) return eq[i].code;
  }
  return '';
}

function getLeaderForCountry(code, eq) {
  for (var i = 0; i < eq.length; i++) {
    if (eq[i].code === code) return eq[i].leader;
  }
  return '';
}

function readConfig(ss) {
  var ws = ss.getSheetByName(SHEET_CONFIG);
  if (!ws) return {};
  var lr = ws.getLastRow();
  if (lr < 3) return {};
  var data = ws.getRange(3, 1, lr - 2, 2).getValues();
  var c = {};
  data.forEach(function(r) { if (r[0]) c[r[0]] = r[1]; });
  return c;
}

function countBizDays(start, end) {
  var count = 0;
  var cur = new Date(start);
  while (cur < end) {
    cur.setDate(cur.getDate() + 1);
    var d = cur.getDay();
    if (d !== 0 && d !== 6) count++;
  }
  return count;
}

function moveToHistorial(ss, wsA, row) {
  var wsH = ss.getSheetByName(SHEET_HISTORIAL);
  if (!wsH) return; // nothing to do; preserve original row
  var lc = Math.min(wsA.getLastColumn(), TASK_COLS);
  var rd = wsA.getRange(row, 1, 1, lc).getValues()[0];
  while (rd.length < TASK_COLS) rd.push('');
  var hl = wsH.getLastRow();
  var nid = hl >= 4 ? wsH.getRange(hl, 1).getValue() + 1 : 1;
  rd[0] = nid;
  wsH.appendRow(rd);
  wsA.deleteRow(row);
  renumberTasks(wsA);
}

function renumberTasks(ws) {
  var lr = ws.getLastRow();
  if (lr < 4) return;
  for (var i = 4; i <= lr; i++) ws.getRange(i, 1).setValue(i - 3);
}

function getCurrentWeekLabel() {
  var now = new Date();
  var mon = new Date(now);
  mon.setDate(now.getDate() - (now.getDay() === 0 ? 6 : now.getDay() - 1));
  var fri = new Date(mon);
  fri.setDate(mon.getDate() + 4);
  var m = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  return mon.getDate() + '-' + fri.getDate() + ' ' + m[fri.getMonth()] + ' ' + fri.getFullYear();
}

// ════════════════════════════════════════════════════════════════
// SLACK HELPERS
// ----------------------------------------------------------------
// These return ContentService outputs because they're hit by the
// Slack webhook handler, not by google.script.run.
// ════════════════════════════════════════════════════════════════

function _slackJsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Find the best-matching active task by name (fuzzy "shared word
 * count"). Returns the index in `data` plus the score, or {idx: -1}.
 */
function _findTaskByName(data, query) {
  var st = (query || '').toLowerCase();
  if (!st) return { idx: -1, score: 0 };
  var words = st.split(/\s+/).filter(function(w) { return w.length > 2; });
  if (!words.length) return { idx: -1, score: 0 };
  var best = { idx: -1, score: 0 };
  for (var i = 0; i < data.length; i++) {
    var n = (data[i][1] || '').toLowerCase();
    if (!n) continue;
    var sc = 0;
    words.forEach(function(w) { if (n.indexOf(w) >= 0) sc++; });
    if (sc > best.score) { best = { idx: i, score: sc }; }
  }
  return best;
}

function handleCloseTask(params) {
  var ss = _trackerSpreadsheet();
  var ws = ss.getSheetByName(SHEET_ACTIVO);
  if (!ws) return _slackJsonResponse({ success: false, message: 'Sheet no encontrado' });
  var lr = ws.getLastRow();
  if (lr < 4) return _slackJsonResponse({ success: false, message: 'No hay tareas activas' });

  var lc = Math.min(ws.getLastColumn(), TASK_COLS);
  var data = ws.getRange(4, 1, lr - 3, lc).getValues();
  var match = _findTaskByName(data, params.task_name || params.message_text || '');
  if (match.idx >= 0 && match.score >= 1) {
    var row = match.idx + 4;
    var tn = data[match.idx][1];
    var tid = data[match.idx][0];
    ws.getRange(row, 7).setValue('Listo');
    ws.getRange(row, 10).setValue(new Date());
    moveToHistorial(ss, ws, row);
    return _slackJsonResponse({
      success: true,
      message: 'Tarea #' + tid + ' "' + tn + '" marcada como Listo y movida a Historial'
    });
  }
  return _slackJsonResponse({ success: false, message: 'No encontré una tarea que coincida' });
}

function handleBlockTask(params) {
  var ss = _trackerSpreadsheet();
  var ws = ss.getSheetByName(SHEET_ACTIVO);
  if (!ws) return _slackJsonResponse({ success: false, message: 'Sheet no encontrado' });
  var lr = ws.getLastRow();
  if (lr < 4) return _slackJsonResponse({ success: false, message: 'No hay tareas activas' });

  var lc = Math.min(ws.getLastColumn(), TASK_COLS);
  var data = ws.getRange(4, 1, lr - 3, lc).getValues();
  var match = _findTaskByName(data, params.task_name || '');
  if (match.idx >= 0 && match.score >= 1) {
    var row = match.idx + 4;
    var tn = data[match.idx][1];
    var tid = data[match.idx][0];
    ws.getRange(row, 7).setValue('Bloqueado');
    var existing = ws.getRange(row, 11).getValue() || '';
    var note = '⛔ ' + (params.reason || '') + ' (' + (params.slack_user || '') + ', '
             + new Date().toLocaleDateString('es-CO') + ')';
    ws.getRange(row, 11).setValue((existing ? existing + ' | ' : '') + note);
    return _slackJsonResponse({ success: true, message: 'Tarea bloqueada: #' + tid + ' "' + tn + '"' });
  }
  return _slackJsonResponse({ success: false, message: 'No encontré una tarea que coincida' });
}

/** Smoke test for the Apps Script editor. */
function testData() {
  var d = getTrackerData();
  Logger.log('Tasks: ' + d.tasks.length);
  Logger.log('Equipos: ' + d.equipos.length);
  Logger.log('Projects: ' + d.projects.length);
  Logger.log('Team: ' + d.team.length);
}
