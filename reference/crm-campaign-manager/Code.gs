// ============================================================
// CAMPAIGN MANAGER — Google Apps Script Web App
// ============================================================

var SPREADSHEET_ID  = '1yzcRTWhdVlm9G-M2--0KIS8_nnsGa93xg3xerL4ZQqc';
var SHEET_NAME      = 'Base';
var SLACK_BOT_TOKEN = 'xoxb-REDACTADO__ROTAR_Y_MOVER_A_SCRIPT_PROPERTIES';
var SLACK_CHANNEL   = 'C09S1BGKUQJ';
var TIMEZONE        = 'America/Mexico_City';

// ============================================================
// HEADERS
// ============================================================
var HEADERS = [
  'TICKET_ID',             // A   0
  'FECHA_SOLICITUD',       // B   1
  'EMAIL_SOLICITANTE',     // C   2
  'country',               // D   3
  'owner',                 // E   4
  'fecha_inicio',          // F   5
  'fecha_fin',             // G   6
  'hora_inicio',           // H   7
  'hora_fin',              // I   8
  'tipo_oferta',           // J   9
  'segmentacion',          // K   10
  'discount',              // L   11
  'minimo_compra',         // M   12
  'maximo_descuento',      // N   13
  'vertical',              // O   14
  'squad',                 // P   15
  'strategy',              // Q   16
  'descripcion',           // R   17
  'pct_alianza',           // S   18
  'id_alianza',            // T   19
  'id_orden_compra',       // U   20
  'store_type',            // V   21
  'brand_id',              // W   22
  'store_ids',             // X   23
  'product_id',            // Y   24
  'exclude_store_ids',     // Z   25
  'max_unidades_orden',    // AA  26
  'days_hours',            // AB  27
  'brand_name',            // AC  28
  'metodos_pago',          // AD  29
  'cc_type',               // AE  30
  'bin',                   // AF  31
  'max_ordenes_usuario',   // AG  32
  'presupuesto',           // AH  33
  'tipo_usuario_app',        // AI  34
  'user_not_in_tag',         // AJ  35
  'terminos_y_condiciones',  // AK  36
  'prime',                   // AL  37
  'is_deal_of_the_day',      // AM  38
  'is_offer_on_top',         // AN  39
  'max_quantity_global',     // AO  40
  'send_products_to_braze',  // AP  41
  'cashback_days_to_end',    // AQ  42
  'store_types_redencion',   // AR  43
  'fecha_inicio_redencion',  // AS  44
  'fecha_fin_redencion',     // AT  45
  'hora_inicio_redencion',   // AU  46
  'hora_fin_redencion',      // AV  47
  'stores_redencion_si',     // AW  48
  'stores_redencion_no',     // AX  49
  'item',                    // AY  50
  'STATUS',                  // AZ  51
  'COMENTARIOS',             // BA  52
  'GLOBAL_OFFER_ID',         // BB  53
  'THREAD_TS'                // BC  54
];

// ── Índices clave ────────────────────────────────────────────
var IDX = {
  TICKET_ID:              0,
  FECHA_SOLICITUD:        1,
  EMAIL_SOLICITANTE:      2,
  COUNTRY:                3,
  OWNER:                  4,
  FECHA_INICIO:           5,
  FECHA_FIN:              6,
  HORA_INICIO:            7,
  HORA_FIN:               8,
  TIPO_OFERTA:            9,
  SEGMENTACION:           10,
  DISCOUNT:               11,
  MINIMO_COMPRA:          12,
  MAXIMO_DESCUENTO:       13,
  VERTICAL:               14,
  SQUAD:                  15,
  STRATEGY:               16,
  DESCRIPCION:            17,
  PCT_ALIANZA:            18,
  ID_ALIANZA:             19,
  ID_ORDEN_COMPRA:        20,
  STORE_TYPE:             21,
  BRAND_ID:               22,
  STORE_IDS:              23,
  PRODUCT_ID:             24,
  EXCLUDE_STORE_IDS:      25,
  MAX_UNIDADES_ORDEN:     26,
  DAYS_HOURS:             27,
  BRAND_NAME:             28,
  METODOS_PAGO:           29,
  CC_TYPE:                30,
  BIN:                    31,
  MAX_ORDENES_USUARIO:    32,
  PRESUPUESTO:            33,
  TIPO_USUARIO_APP:         34,
  USER_NOT_IN_TAG:          35,
  TERMINOS_Y_CONDICIONES:   36,
  PRIME:                    37,
  IS_DEAL_OF_THE_DAY:       38,
  IS_OFFER_ON_TOP:          39,
  MAX_QUANTITY_GLOBAL:      40,
  SEND_PRODUCTS_TO_BRAZE:   41,
  CASHBACK_DAYS_TO_END:     42,
  STORE_TYPES_REDENCION:    43,
  FECHA_INICIO_REDENCION:   44,
  FECHA_FIN_REDENCION:      45,
  HORA_INICIO_REDENCION:    46,
  HORA_FIN_REDENCION:       47,
  STORES_REDENCION_SI:      48,
  STORES_REDENCION_NO:      49,
  ITEM:                     50,
  STATUS:                   51,
  COMENTARIOS:              52,
  GLOBAL_OFFER_ID:          53,
  THREAD_TS:                54
};

// ============================================================
// RAPPICREDITOS — HOJA Y ESQUEMA
// ============================================================
var SHEET_NAME_RC = 'RappiCreditos';

var HEADERS_RC = [
  'TICKET_ID', 'FECHA_SOLICITUD', 'EMAIL_SOLICITANTE',
  'country', 'owner', 'fecha_inicio', 'fecha_fin', 'hora_inicio', 'hora_fin',
  'squad', 'strategy', 'pago', 'pct_alianza', 'id_alianza', 'id_orden_compra',
  'budget', 'users_file_url', 'tope_pedido', 'vigencia_dias', 'vigencia_fecha',
  'store_ids', 'descripcion',
  'STATUS', 'COMENTARIOS', 'GLOBAL_OFFER_ID', 'THREAD_TS'
];

var IDX_RC = {
  TICKET_ID: 0, FECHA_SOLICITUD: 1, EMAIL_SOLICITANTE: 2,
  COUNTRY: 3, OWNER: 4, FECHA_INICIO: 5, FECHA_FIN: 6, HORA_INICIO: 7, HORA_FIN: 8,
  SQUAD: 9, STRATEGY: 10, PAGO: 11, PCT_ALIANZA: 12, ID_ALIANZA: 13, ID_ORDEN_COMPRA: 14,
  BUDGET: 15, USERS_FILE_URL: 16, TOPE_PEDIDO: 17, VIGENCIA_DIAS: 18, VIGENCIA_FECHA: 19,
  STORE_IDS: 20, DESCRIPCION: 21,
  STATUS: 22, COMENTARIOS: 23, GLOBAL_OFFER_ID: 24, THREAD_TS: 25
};

function openRCSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME_RC);
  if (!sheet) throw new Error('No existe la hoja "' + SHEET_NAME_RC + '"');
  return sheet;
}

function getRCSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME_RC);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME_RC);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS_RC.length).setValues([HEADERS_RC]);

    var hRange = sheet.getRange(1, 1, 1, HEADERS_RC.length);
    hRange.setBackground('#FF441F').setFontColor('#FFFFFF').setFontWeight('bold')
          .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 36);
    sheet.setFrozenRows(1);

    sheet.getRange(1, 1, 1, 3).setBackground('#1A1916');
    sheet.getRange(1, IDX_RC.STATUS + 1, 1, 3).setBackground('#4C1D95');

    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente','En Revisión','Aprobado','Rechazado','Completado'], true).build();
    sheet.getRange(2, IDX_RC.STATUS + 1, 1000, 1).setDataValidation(rule);

    sheet.setColumnWidth(IDX_RC.USERS_FILE_URL + 1, 260);
    sheet.setColumnWidth(IDX_RC.STORE_IDS + 1, 200);
    sheet.setColumnWidth(IDX_RC.DESCRIPCION + 1, 200);
  }
  return sheet;
}

// ============================================================
// ROUTER
// ============================================================
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';
  var template;
  if (page === 'form') {
    template = HtmlService.createTemplateFromFile('home');
  } else if (page === 'status') {
    template = HtmlService.createTemplateFromFile('status');
    template.ticketId = (e && e.parameter && e.parameter.ticket) ? e.parameter.ticket : '';
  } else if (page === 'admin') {
    template = HtmlService.createTemplateFromFile('admin');
  } else {
    template = HtmlService.createTemplateFromFile('home');
  }
  return template.evaluate()
    .setTitle('Campaign Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// ============================================================
// openSheet — lectura
// ============================================================
function openSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('No existe la hoja "' + SHEET_NAME + '"');
  return sheet;
}

// ============================================================
// getSheet — escritura, crea encabezados si vacía
// ============================================================
function getSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

    var hRange = sheet.getRange(1, 1, 1, HEADERS.length);
    hRange.setBackground('#FF441F').setFontColor('#FFFFFF').setFontWeight('bold')
          .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.setRowHeight(1, 36);
    sheet.setFrozenRows(1);

    // Gestión interna (A-C) en negro
    sheet.getRange(1, 1, 1, 3).setBackground('#1A1916');

    // Columnas que llena el equipo — amarillo oscuro
    var teamCols = [IDX.TIPO_USUARIO_APP + 1, IDX.USER_NOT_IN_TAG + 1, IDX.ITEM + 1];
    teamCols.forEach(function(col) {
      sheet.getRange(1, col).setBackground('#92400E').setFontColor('#FEF3C7');
    });

    // Respuesta del equipo (STATUS, COMENTARIOS, GLOBAL_OFFER_ID) — morado
    sheet.getRange(1, IDX.STATUS + 1, 1, 3).setBackground('#4C1D95');

    // Validación STATUS
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pendiente','En Revisión','Aprobado','Rechazado','Completado'], true).build();
    sheet.getRange(2, IDX.STATUS + 1, 1000, 1).setDataValidation(rule);

    // Anchos de columna
    sheet.setColumnWidth(IDX.STORE_IDS + 1, 200);
    sheet.setColumnWidth(IDX.SEGMENTACION + 1, 180);
    sheet.setColumnWidth(IDX.DESCRIPCION + 1, 200);
    sheet.setColumnWidth(IDX.DAYS_HOURS + 1, 200);
    sheet.setColumnWidth(IDX.STORES_REDENCION_SI + 1, 180);
    sheet.setColumnWidth(IDX.STORES_REDENCION_NO + 1, 180);
  }
  return sheet;
}

function initSheet() { getSheet(); }

// ============================================================
// saveRequest
// ============================================================
function saveRequest(data) {
  try {
    var sheet    = getSheet();
    var ticketId = generateTicketId();
    var fecha    = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    var email = '';
    try { email = Session.getActiveUser().getEmail(); } catch(ex) {}
    if (!email) email = data.ownerEmail || data.emailSolicitante || 'sin-email';

    var countryMap = {
      'Colombia':'co','México':'mx','Argentina':'ar','Chile':'cl',
      'Perú':'pe','Uruguay':'uy','Ecuador':'ec','Costa Rica':'cr'
    };
    var countryCode = countryMap[data.pais] || (data.pais || '').toLowerCase();

    var row = new Array(HEADERS.length).fill('');

    // ── GESTIÓN INTERNA ──
    row[IDX.TICKET_ID]         = ticketId;
    row[IDX.FECHA_SOLICITUD]   = fecha;
    row[IDX.EMAIL_SOLICITANTE] = email;

    // ── TEMPLATE ──
    row[IDX.COUNTRY]            = countryCode;
    row[IDX.OWNER]              = data.ownerEmail              || '';
    row[IDX.FECHA_INICIO]       = data.fechaInicio             || '';
    row[IDX.FECHA_FIN]          = data.fechaFin                || '';
    row[IDX.HORA_INICIO]        = data.horaInicio              || '';
    row[IDX.HORA_FIN]           = data.horaFin                 || '';
    row[IDX.TIPO_OFERTA]        = (data.tipoOferta || '').toLowerCase().replace(/ /g,'_');
    row[IDX.SEGMENTACION]       = data.segmento                || '';
    row[IDX.DISCOUNT]           = data.discount                || '';
    row[IDX.MINIMO_COMPRA]      = data.minimoCompra            || '';
    row[IDX.MAXIMO_DESCUENTO]   = data.maximoDescuento         || '';
    row[IDX.VERTICAL]           = data.vertical                || '';
    row[IDX.SQUAD]              = data.squad                   || '';
    row[IDX.STRATEGY]           = data.strategy                || '';
    row[IDX.DESCRIPCION]        = data.descripcion             || data.comentariosSolicitud || '';
    row[IDX.PCT_ALIANZA]        = data.pctAlianza              || '';
    row[IDX.ID_ALIANZA]         = data.idAlianza               || '';
    row[IDX.ID_ORDEN_COMPRA]    = data.ordenCompra ? String(data.ordenCompra) : '';
    row[IDX.STORE_TYPE]         = '';
    row[IDX.BRAND_ID]           = data.brandId                 || '';
    row[IDX.STORE_IDS]          = data.storeIds                || '';
    row[IDX.PRODUCT_ID]         = data.productIds              || '';
    row[IDX.EXCLUDE_STORE_IDS]  = data.excludeStoreIds         || '';
    row[IDX.MAX_UNIDADES_ORDEN] = data.maxUnidades             || '';
    row[IDX.DAYS_HOURS]         = data.daysHours               || '';
    row[IDX.BRAND_NAME]         = data.brandName               || '';
    row[IDX.METODOS_PAGO]       = data.metodosPago             || '';
    row[IDX.CC_TYPE]            = data.ccType                  || '';
    row[IDX.BIN]                = data.bin                     || '';
    row[IDX.MAX_ORDENES_USUARIO]= data.maxOrdenes              || '';
    row[IDX.PRESUPUESTO]        = data.budget                  || '';
    row[IDX.TIPO_USUARIO_APP]          = '';
    row[IDX.USER_NOT_IN_TAG]            = '';
    row[IDX.TERMINOS_Y_CONDICIONES]     = data.linkTyC                || '';
    row[IDX.PRIME]              = data.prime                   || '';
    row[IDX.IS_DEAL_OF_THE_DAY] = data.isDealOfTheDay === 'SI' ? 'SI' : '';
    row[IDX.IS_OFFER_ON_TOP]    = data.isOfferOnTop === 'SI'   ? 'SI' : '';
    row[IDX.MAX_QUANTITY_GLOBAL]= data.maxQuantityGlobal       || '';
    row[IDX.SEND_PRODUCTS_TO_BRAZE] = data.sendProductsToBraze === 'SI' ? 'SI' : '';
    row[IDX.CASHBACK_DAYS_TO_END]   = data.cashbackDaysToEnd   || '';
    row[IDX.STORE_TYPES_REDENCION]  = data.storeTypeRedencion  || '';
    row[IDX.FECHA_INICIO_REDENCION] = data.fechaInicioRedencion|| '';
    row[IDX.FECHA_FIN_REDENCION]    = data.fechaFinRedencion   || '';
    row[IDX.HORA_INICIO_REDENCION]  = data.horaInicioRedencion || '';
    row[IDX.HORA_FIN_REDENCION]     = data.horaFinRedencion    || '';
    row[IDX.STORES_REDENCION_SI]    = data.storesRedencionSi   || '';
    row[IDX.STORES_REDENCION_NO]    = data.storesRedencionNo   || '';
    row[IDX.ITEM]               = '';

    // ── RESPUESTA ──
    row[IDX.STATUS]             = 'Pendiente';
    row[IDX.COMENTARIOS]        = '';
    row[IDX.GLOBAL_OFFER_ID]    = '';
    row[IDX.THREAD_TS]          = '';

    sheet.appendRow(row);
    var newRow = sheet.getLastRow();

    sheet.getRange(newRow, IDX.STATUS + 1)
      .setBackground('#FEF3C7').setFontWeight('bold').setHorizontalAlignment('center');

    // Crear CSV en Drive
    try { createCampaignCSV(ticketId, row); } catch(e) { Logger.log('CSV error: ' + e.message); }

    try {
      var threadTs = notifySlack(ticketId, email, data, countryCode);
      if (threadTs) {
        var tsCell = sheet.getRange(newRow, IDX.THREAD_TS + 1);
        tsCell.setNumberFormat('@');
        tsCell.setValue("'" + String(threadTs));
      }
    } catch(e) { Logger.log('Slack notify error: ' + e.message); }

    return { success: true, ticketId: ticketId, fecha: fecha, email: email };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// getTicketStatus
// ============================================================
function getTicketStatus(ticketId) {
  try {
    var buscar = String(ticketId || '').trim().toUpperCase();
    if (!buscar) return { found: false, debug: 'ID vacío' };

    var normalizado = buscar.replace(/^(CAM-)(\d+)$/, function(m, p, n) {
      while (n.length < 4) n = '0' + n;
      return p + n;
    });

    var found = findTicketRow(buscar) || findTicketRow(normalizado);
    if (!found) return { found: false, debug: 'Ticket "' + buscar + '" no encontrado.' };

    if (found.kind === 'rappicreditos') {
      var r = found.rowData, idx = found.idx;
      return {
        found: true,
        kind: 'rappicreditos',
        ticket: {
          ticketId:      String(r[idx.TICKET_ID] || ''),
          fecha:         r[idx.FECHA_SOLICITUD] ? Utilities.formatDate(new Date(r[idx.FECHA_SOLICITUD]), TIMEZONE, 'yyyy-MM-dd HH:mm') : '',
          email:         String(r[idx.EMAIL_SOLICITANTE] || ''),
          pais:          String(r[idx.COUNTRY] || ''),
          owner:         String(r[idx.OWNER] || ''),
          fechaInicio:   String(r[idx.FECHA_INICIO] || ''),
          fechaFin:      String(r[idx.FECHA_FIN] || ''),
          horaInicio:    String(r[idx.HORA_INICIO] || ''),
          horaFin:       String(r[idx.HORA_FIN] || ''),
          squad:         String(r[idx.SQUAD] || ''),
          strategy:      String(r[idx.STRATEGY] || ''),
          pago:          String(r[idx.PAGO] || ''),
          pctAlianza:    String(r[idx.PCT_ALIANZA] || ''),
          idAlianza:     String(r[idx.ID_ALIANZA] || ''),
          ordenCompra:   String(r[idx.ID_ORDEN_COMPRA] || ''),
          budget:        String(r[idx.BUDGET] || ''),
          usersFileUrl:  String(r[idx.USERS_FILE_URL] || ''),
          topePedido:    String(r[idx.TOPE_PEDIDO] || ''),
          vigenciaDias:  String(r[idx.VIGENCIA_DIAS] || ''),
          vigenciaFecha: String(r[idx.VIGENCIA_FECHA] || ''),
          storeIds:      String(r[idx.STORE_IDS] || ''),
          descripcion:   String(r[idx.DESCRIPCION] || ''),
          status:        String(r[idx.STATUS] || ''),
          comentarios:   String(r[idx.COMENTARIOS] || ''),
          globalOfferId: String(r[idx.GLOBAL_OFFER_ID] || '')
        }
      };
    }

    var row = found.rowData;
    return {
      found: true,
      kind: 'campaign',
      ticket: {
        ticketId:             String(row[IDX.TICKET_ID]              || ''),
        fecha:                row[IDX.FECHA_SOLICITUD] ? Utilities.formatDate(new Date(row[IDX.FECHA_SOLICITUD]), TIMEZONE, 'yyyy-MM-dd HH:mm') : '',
        email:                String(row[IDX.EMAIL_SOLICITANTE]      || ''),
        pais:                 String(row[IDX.COUNTRY]                || ''),
        owner:                String(row[IDX.OWNER]                  || ''),
        fechaInicio:          String(row[IDX.FECHA_INICIO]           || ''),
        fechaFin:             String(row[IDX.FECHA_FIN]              || ''),
        horaInicio:           String(row[IDX.HORA_INICIO]            || ''),
        horaFin:              String(row[IDX.HORA_FIN]               || ''),
        tipoOferta:           String(row[IDX.TIPO_OFERTA]            || ''),
        segmento:             String(row[IDX.SEGMENTACION]           || ''),
        discount:             String(row[IDX.DISCOUNT]               || ''),
        minimoCompra:         String(row[IDX.MINIMO_COMPRA]          || ''),
        maximoDescuento:      String(row[IDX.MAXIMO_DESCUENTO]       || ''),
        vertical:             String(row[IDX.VERTICAL]               || ''),
        squad:                String(row[IDX.SQUAD]                  || ''),
        strategy:             String(row[IDX.STRATEGY]               || ''),
        descripcion:          String(row[IDX.DESCRIPCION]            || ''),
        pctAlianza:           String(row[IDX.PCT_ALIANZA]            || ''),
        idAlianza:            String(row[IDX.ID_ALIANZA]             || ''),
        ordenCompra:          String(row[IDX.ID_ORDEN_COMPRA]        || ''),
        storeType:            String(row[IDX.STORE_TYPE]             || ''),
        brandId:              String(row[IDX.BRAND_ID]               || ''),
        storeIds:             String(row[IDX.STORE_IDS]              || ''),
        productId:            String(row[IDX.PRODUCT_ID]             || ''),
        excludeStoreIds:      String(row[IDX.EXCLUDE_STORE_IDS]      || ''),
        maxUnidades:          String(row[IDX.MAX_UNIDADES_ORDEN]     || ''),
        daysHours:            String(row[IDX.DAYS_HOURS]             || ''),
        brandName:            String(row[IDX.BRAND_NAME]             || ''),
        metodosPago:          String(row[IDX.METODOS_PAGO]           || ''),
        ccType:               String(row[IDX.CC_TYPE]                || ''),
        bin:                  String(row[IDX.BIN]                    || ''),
        maxOrdenes:           String(row[IDX.MAX_ORDENES_USUARIO]    || ''),
        presupuesto:          String(row[IDX.PRESUPUESTO]            || ''),
        terminosYCondiciones: String(row[IDX.TERMINOS_Y_CONDICIONES] || ''),
        prime:                String(row[IDX.PRIME]                  || ''),
        isDealOfTheDay:       String(row[IDX.IS_DEAL_OF_THE_DAY]     || ''),
        isOfferOnTop:         String(row[IDX.IS_OFFER_ON_TOP]        || ''),
        maxQuantityGlobal:    String(row[IDX.MAX_QUANTITY_GLOBAL]    || ''),
        sendProductsToBraze:  String(row[IDX.SEND_PRODUCTS_TO_BRAZE]|| ''),
        cashbackDaysToEnd:    String(row[IDX.CASHBACK_DAYS_TO_END]   || ''),
        storeTypesRedencion:  String(row[IDX.STORE_TYPES_REDENCION]  || ''),
        fechaInicioRedencion: String(row[IDX.FECHA_INICIO_REDENCION] || ''),
        fechaFinRedencion:    String(row[IDX.FECHA_FIN_REDENCION]    || ''),
        horaInicioRedencion:  String(row[IDX.HORA_INICIO_REDENCION]  || ''),
        horaFinRedencion:     String(row[IDX.HORA_FIN_REDENCION]     || ''),
        storesRedencionSi:    String(row[IDX.STORES_REDENCION_SI]    || ''),
        storesRedencionNo:    String(row[IDX.STORES_REDENCION_NO]    || ''),
        status:               String(row[IDX.STATUS]                 || ''),
        comentarios:          String(row[IDX.COMENTARIOS]            || ''),
        globalOfferId:        String(row[IDX.GLOBAL_OFFER_ID]        || '')
      }
    };
  } catch (err) {
    return { found: false, debug: 'ERROR: ' + err.message };
  }
}

// ============================================================
// getTicketsByEmail
// ============================================================
function getTicketsByEmail(email) {
  try {
    var buscar = String(email || '').trim().toLowerCase();
    if (!buscar) return { found: false, tickets: [] };

    var results = [];

    var baseSheet = openSheet();
    if (baseSheet.getLastRow() >= 2) {
      var baseData = baseSheet.getDataRange().getValues();
      for (var i = 1; i < baseData.length; i++) {
        var row = baseData[i];
        var rowEmail = String(row[IDX.EMAIL_SOLICITANTE] || '').trim().toLowerCase();
        var rowOwner = String(row[IDX.OWNER]             || '').trim().toLowerCase();
        if (rowEmail.includes(buscar) || rowOwner.includes(buscar)) {
          results.push({
            kind:        'campaign',
            ticketId:    String(row[IDX.TICKET_ID]         || ''),
            fecha:       row[IDX.FECHA_SOLICITUD] ? Utilities.formatDate(new Date(row[IDX.FECHA_SOLICITUD]), TIMEZONE, 'yyyy-MM-dd HH:mm') : '',
            email:       String(row[IDX.EMAIL_SOLICITANTE] || ''),
            owner:       String(row[IDX.OWNER]             || ''),
            pais:        String(row[IDX.COUNTRY]           || ''),
            tipoOferta:  String(row[IDX.TIPO_OFERTA]       || ''),
            brandId:     String(row[IDX.BRAND_ID]          || ''),
            brandName:   String(row[IDX.BRAND_NAME]        || ''),
            fechaInicio: String(row[IDX.FECHA_INICIO]      || ''),
            fechaFin:    String(row[IDX.FECHA_FIN]         || ''),
            discount:    String(row[IDX.DISCOUNT]          || ''),
            presupuesto: String(row[IDX.PRESUPUESTO]       || ''),
            status:      String(row[IDX.STATUS]            || ''),
            comentarios: String(row[IDX.COMENTARIOS]       || ''),
            globalOfferId: String(row[IDX.GLOBAL_OFFER_ID] || '')
          });
        }
      }
    }

    var rcSheet;
    try { rcSheet = openRCSheet(); } catch (e) { rcSheet = null; }
    if (rcSheet && rcSheet.getLastRow() >= 2) {
      var rcData = rcSheet.getDataRange().getValues();
      for (var j = 1; j < rcData.length; j++) {
        var r = rcData[j];
        var rEmail = String(r[IDX_RC.EMAIL_SOLICITANTE] || '').trim().toLowerCase();
        var rOwner = String(r[IDX_RC.OWNER]              || '').trim().toLowerCase();
        if (rEmail.includes(buscar) || rOwner.includes(buscar)) {
          results.push({
            kind:        'rappicreditos',
            ticketId:    String(r[IDX_RC.TICKET_ID]         || ''),
            fecha:       r[IDX_RC.FECHA_SOLICITUD] ? Utilities.formatDate(new Date(r[IDX_RC.FECHA_SOLICITUD]), TIMEZONE, 'yyyy-MM-dd HH:mm') : '',
            email:       String(r[IDX_RC.EMAIL_SOLICITANTE] || ''),
            owner:       String(r[IDX_RC.OWNER]             || ''),
            pais:        String(r[IDX_RC.COUNTRY]           || ''),
            tipoOferta:  'rappicreditos',
            brandId:     '',
            brandName:   '',
            fechaInicio: String(r[IDX_RC.FECHA_INICIO]      || ''),
            fechaFin:    String(r[IDX_RC.FECHA_FIN]         || ''),
            discount:    '',
            presupuesto: String(r[IDX_RC.BUDGET]            || ''),
            status:      String(r[IDX_RC.STATUS]            || ''),
            comentarios: String(r[IDX_RC.COMENTARIOS]       || ''),
            globalOfferId: String(r[IDX_RC.GLOBAL_OFFER_ID] || '')
          });
        }
      }
    }

    results.reverse();
    return { found: results.length > 0, tickets: results };
  } catch (err) {
    return { found: false, tickets: [], debug: err.message };
  }
}

// ============================================================
// getAllTickets
// ============================================================
function getAllTickets() {
  try {
    var sheet = openSheet();
    if (sheet.getLastRow() < 2) return [];
    var data   = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || !row[IDX.TICKET_ID]) continue;
      function safe(idx) { return row.length > idx ? String(row[idx] || '') : ''; }
      result.push({
        ticketId:      safe(IDX.TICKET_ID),
        fecha:         safe(IDX.FECHA_SOLICITUD),
        email:         safe(IDX.EMAIL_SOLICITANTE),
        pais:          safe(IDX.COUNTRY),
        owner:         safe(IDX.OWNER),
        tipoOferta:    safe(IDX.TIPO_OFERTA),
        vertical:      safe(IDX.VERTICAL),
        squad:         safe(IDX.SQUAD),
        strategy:      safe(IDX.STRATEGY),
        brandId:       safe(IDX.BRAND_ID),
        brandName:     safe(IDX.BRAND_NAME),
        storeType:     safe(IDX.STORE_TYPE),
        discount:      safe(IDX.DISCOUNT),
        segmento:      safe(IDX.SEGMENTACION),
        presupuesto:   safe(IDX.PRESUPUESTO),
        fechaInicio:   safe(IDX.FECHA_INICIO),
        fechaFin:      safe(IDX.FECHA_FIN),
        status:        safe(IDX.STATUS),
        comentarios:   safe(IDX.COMENTARIOS),
        globalOfferId: safe(IDX.GLOBAL_OFFER_ID)
      });
    }
    result.reverse();
    return result;
  } catch (err) {
    Logger.log('getAllTickets error: ' + err.message);
    return [];
  }
}

// ============================================================
// findTicketRow — busca un ticket en Base o RappiCreditos
// ============================================================
function findTicketRow(ticketId) {
  var id = String(ticketId).trim().toUpperCase();

  var baseSheet = openSheet();
  var baseData  = baseSheet.getDataRange().getValues();
  for (var i = 1; i < baseData.length; i++) {
    if (String(baseData[i][IDX.TICKET_ID]).trim().toUpperCase() === id) {
      return { sheet: baseSheet, row: i + 1, rowData: baseData[i], idx: IDX, kind: 'campaign' };
    }
  }

  var rcSheet;
  try { rcSheet = openRCSheet(); } catch (e) { return null; }
  var rcData = rcSheet.getDataRange().getValues();
  for (var j = 1; j < rcData.length; j++) {
    if (String(rcData[j][IDX_RC.TICKET_ID]).trim().toUpperCase() === id) {
      return { sheet: rcSheet, row: j + 1, rowData: rcData[j], idx: IDX_RC, kind: 'rappicreditos' };
    }
  }

  return null;
}

// ============================================================
// updateStatus
// ============================================================
function updateStatus(ticketId, newStatus, comentarios, globalOfferId) {
  try {
    var found = findTicketRow(ticketId);
    if (!found) return { success: false, message: 'Ticket no encontrado.' };

    var sheet = found.sheet, row = found.row, idx = found.idx;

    sheet.getRange(row, idx.STATUS + 1).setValue(newStatus);
    if (comentarios !== undefined && comentarios !== null) {
      sheet.getRange(row, idx.COMENTARIOS + 1).setValue(comentarios);
    }
    if (globalOfferId) {
      sheet.getRange(row, idx.GLOBAL_OFFER_ID + 1).setValue(globalOfferId)
        .setBackground('#EDE9FE').setFontWeight('bold')
        .setFontColor('#4C1D95').setHorizontalAlignment('center');
    }
    var colors = {
      'Pendiente':'#FEF3C7','En Revisión':'#DBEAFE',
      'Aprobado':'#DCFCE7','Rechazado':'#FEE2E2','Completado':'#EDE9FE'
    };
    sheet.getRange(row, idx.STATUS + 1)
      .setBackground(colors[newStatus] || '#F3F4F6')
      .setFontWeight('bold').setHorizontalAlignment('center');
    try { notifySlackUpdate(ticketId, newStatus, comentarios, globalOfferId); } catch(e) {}
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// updateGlobalOffer
// ============================================================
function updateGlobalOffer(ticketId, globalOfferId) {
  try {
    var found = findTicketRow(ticketId);
    if (!found) return { success: false, message: 'Ticket no encontrado.' };
    found.sheet.getRange(found.row, found.idx.GLOBAL_OFFER_ID + 1).setValue(globalOfferId)
      .setBackground('#EDE9FE').setFontWeight('bold')
      .setFontColor('#4C1D95').setHorizontalAlignment('center');
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// Slack — utilidades base
// ============================================================
function slackApi(method, payload) {
  var res = UrlFetchApp.fetch('https://slack.com/api/' + method, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    headers: { 'Authorization': 'Bearer ' + SLACK_BOT_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return JSON.parse(res.getContentText());
}

// ============================================================
// saveThreadTs — guarda como texto puro para evitar corrupción
// ============================================================
function saveThreadTs(ticketId, threadTs) {
  try {
    var found = findTicketRow(ticketId);
    if (!found) return;
    var cell = found.sheet.getRange(found.row, found.idx.THREAD_TS + 1);
    cell.setNumberFormat('@');
    cell.setValue("'" + String(threadTs));
  } catch(e) { Logger.log('saveThreadTs error: ' + e.message); }
}

// ============================================================
// getThreadTs — lee y limpia correctamente, protege contra
//              notación científica y apóstrofes de texto
// ============================================================
function getThreadTs(ticketId) {
  try {
    var found = findTicketRow(ticketId);
    if (!found) return '';
    var ts = found.rowData[found.idx.THREAD_TS];
    if (!ts && ts !== 0) return '';
    var clean = String(ts)
      .replace(/^'+/, '')
      .replace(/\s/g, '')
      .trim();
    if (clean.indexOf('e') !== -1 || clean.indexOf('E') !== -1) {
      Logger.log('getThreadTs: ts en notación científica para ' + ticketId + ': ' + clean);
      return '';
    }
    return clean;
  } catch(e) { Logger.log('getThreadTs error: ' + e.message); }
  return '';
}

// ============================================================
// notifySlack — mensaje inicial al crear solicitud
// ============================================================
function notifySlack(ticketId, email, data, countryCode) {
  try {
    var paisFlags = { co:'🇨🇴', mx:'🇲🇽', ar:'🇦🇷', cl:'🇨🇱', pe:'🇵🇪', uy:'🇺🇾', ec:'🇪🇨', cr:'🇨🇷' };
    var flag      = paisFlags[countryCode] || '🌎';
    var tipo      = (data.tipoOferta || '').toUpperCase();
    var brand     = data.brandName || data.brandId || '—';
    var discount  = data.discount ? data.discount + (tipo === 'CASHBACK' || tipo === 'PERCENTAGE' ? '%' : '') : '—';
    var budget    = data.budget ? Number(data.budget).toLocaleString() : '—';

    var mainRes = slackApi('chat.postMessage', {
      channel: SLACK_CHANNEL,
      text: flag + ' [' + ticketId + '] ' + tipo + ' · ' + brand,
      blocks: [
        {
          type: 'section',
          text: { type: 'mrkdwn',
            text: flag + ' *' + tipo + '* — ' + brand + '\n' +
                  '*Ticket:* `' + ticketId + '`   *Descuento:* ' + discount + '   *Budget:* ' + budget + '\n' +
                  '*Fechas:* ' + (data.fechaInicio||'—') + ' → ' + (data.fechaFin||'—') + '\n' +
                  '*Solicitante:* ' + (data.ownerEmail||email)
          }
        }
      ]
    });

    if (!mainRes.ok) { Logger.log('Slack error: ' + mainRes.error); return null; }
    var threadTs = mainRes.ts;
    saveThreadTs(ticketId, threadTs);

    var lines = [];
    lines.push('*📋 Detalle — ' + ticketId + '*');
    lines.push('');
    lines.push('*Campaña*');
    lines.push('> País: ' + flag + ' ' + (data.pais || countryCode));
    lines.push('> Owner: ' + (data.ownerEmail || email));
    lines.push('> Tipo: ' + tipo);
    lines.push('> Vertical: ' + (data.vertical || '—'));
    lines.push('> Brand ID: ' + (data.brandId || '—') + (data.brandName ? ' (' + data.brandName + ')' : ''));
    lines.push('> Fechas: ' + (data.fechaInicio||'—') + ' → ' + (data.fechaFin||'—'));
    if (data.horaInicio || data.horaFin) lines.push('> Horario: ' + (data.horaInicio||'—') + ' - ' + (data.horaFin||'—'));
    if (data.daysHours) lines.push('> Days/Hours: `' + data.daysHours + '`');
    lines.push('');

    lines.push('*Descuento*');
    lines.push('> Discount: ' + discount);
    lines.push('> Mín. Compra: ' + (data.minimoCompra || '0'));
    lines.push('> Máx. Descuento: ' + (data.maximoDescuento || '—'));
    lines.push('> Máx. órdenes/usuario: ' + (data.maxOrdenes || '—'));
    lines.push('');

    lines.push('*Stores*');
    lines.push('> Cobertura: ' + (data.cobertura || 'Full Brand'));
    if (data.storeIds) lines.push('> Store IDs: `' + data.storeIds.substring(0,300) + (data.storeIds.length > 300 ? '...' : '') + '`');
    if (data.excludeStoreIds) lines.push('> Exclude Store IDs: ' + data.excludeStoreIds);
    lines.push('');

    lines.push('*Segmento & Config*');
    lines.push('> Segmento: ' + (data.segmento || '—'));
    lines.push('> PRIME: ' + (data.prime || 'Todos'));
    if (data.metodosPago) lines.push('> Métodos pago: ' + data.metodosPago);
    if (data.linkTyC) lines.push('> T&C: ' + data.linkTyC);
    lines.push('');

    lines.push('*Pago & Imputación*');
    var pagoLabel = {full_rappi:'Full Rappi', full_aliado:'Full Aliado'}[data.pago] || data.pago || '—';
    lines.push('> Esquema: ' + pagoLabel);
    if (data.squad)      lines.push('> Squad: ' + data.squad);
    if (data.strategy)   lines.push('> Strategy: ' + data.strategy);
    if (data.pctAlianza) lines.push('> % Alianza: ' + data.pctAlianza);
    if (data.idAlianza)  lines.push('> ID Alianza: ' + data.idAlianza);
    if (data.ordenCompra)lines.push('> Orden Compra: ' + data.ordenCompra);
    lines.push('');

    lines.push('*Budget*');
    lines.push('> ' + budget);

    if (data.descripcion) {
      lines.push('');
      lines.push('*💬 Descripción*');
      lines.push('> ' + data.descripcion);
    }

    if (tipo === 'CASHBACK') {
      lines.push('');
      lines.push('*Redención Cashback*');
      if (data.cashbackDaysToEnd) lines.push('> Days to end: ' + data.cashbackDaysToEnd + ' días');
      if (data.storeTypeRedencion) lines.push('> Store types redención: ' + data.storeTypeRedencion);
      if (data.fechaInicioRedencion) lines.push('> Fechas redención: ' + data.fechaInicioRedencion + ' → ' + (data.fechaFinRedencion||'—'));
      if (data.storesRedencionSi) lines.push('> Stores SI: ' + data.storesRedencionSi.substring(0,200));
      if (data.storesRedencionNo) lines.push('> Stores NO: ' + data.storesRedencionNo.substring(0,200));
    }

    slackApi('chat.postMessage', {
      channel: SLACK_CHANNEL,
      thread_ts: threadTs,
      text: 'Detalle ' + ticketId,
      blocks: [{ type:'section', text:{ type:'mrkdwn', text: lines.join('\n') } }]
    });

    var reviewerMsg = getReviewerMessage(countryCode);
    if (reviewerMsg) {
      slackApi('chat.postMessage', {
        channel: SLACK_CHANNEL, thread_ts: threadTs,
        text: reviewerMsg, mrkdwn: true
      });
    }

    return threadTs;
  } catch(e) { Logger.log('Slack error: ' + e.message); return null; }
}

// ============================================================
// notifySlackUpdate — responde en el thread al cambiar status.
//                     Si no hay thread_ts válido, omite el
//                     mensaje en lugar de enviarlo al canal.
// ============================================================
function notifySlackUpdate(ticketId, newStatus, comentarios, globalOfferId) {
  try {
    var threadTs = getThreadTs(ticketId);

    // Sin thread válido: loguea y sale sin enviar al canal
    if (!threadTs || threadTs.length === 0) {
      Logger.log('notifySlackUpdate: thread_ts no disponible para ' + ticketId + ', mensaje omitido.');
      return;
    }

    var resolved   = { Aprobado:'✅', Rechazado:'❌', Completado:'🎉' };
    var inProgress = { 'En Revisión':'👀', Pendiente:'⏳' };
    var text;

    if (resolved[newStatus]) {
      text = resolved[newStatus] + ' *' + newStatus.toUpperCase() + '* — `' + ticketId + '`';
      if (newStatus === 'Aprobado' && globalOfferId) {
        text += '\n🎯 *Global Offer ID:* `' + globalOfferId + '`';
      }
    } else {
      text = (inProgress[newStatus] || '📌') + ' Status actualizado → *' + newStatus + '* — `' + ticketId + '`';
    }
    if (comentarios) text += '\n💬 _' + comentarios + '_';

    slackApi('chat.postMessage', {
      channel:   SLACK_CHANNEL,
      thread_ts: threadTs,
      text:      text,
      blocks:    [{ type:'section', text:{ type:'mrkdwn', text: text } }]
    });
  } catch(e) {
    Logger.log('Slack update error: ' + e.message);
  }
}

// ============================================================
// getReviewerMessage
// ============================================================
function getReviewerMessage(countryCode) {
  var reviewers = {
    'co': '👀 <@U082C55UAG5> estará revisando esta solicitud.',
    'pe': '👀 <@U07D3LV7JKV> estará revisando esta solicitud.',
    'ec': '👀 <@U07D3LV7JKV> estará revisando esta solicitud.',
    'ar': '👀 <@U08BAQVR4EA> estará revisando esta solicitud.',
    'cl': '👀 <@U08BAQVR4EA> estará revisando esta solicitud.',
    'uy': '👀 <@U08BAQVR4EA> estará revisando esta solicitud.',
    'mx': '👀 <@U0AUV5H65EH> y <@U082DA3KE5R> estarán revisando esta solicitud.',
    'cr': '👀 <@U0ANNEQJGE7> estará revisando esta solicitud.'
  };
  return reviewers[countryCode] || '';
}

// ============================================================
// generateTicketId
// ============================================================
function generateTicketId() {
  var existing = {};
  var maxRow = 0;

  var baseSheet = getSheet();
  if (baseSheet.getLastRow() > 1) {
    var baseIds = baseSheet.getRange(2, IDX.TICKET_ID + 1, baseSheet.getLastRow() - 1, 1).getValues();
    baseIds.forEach(function(r) { if (r[0]) existing[String(r[0]).trim().toUpperCase()] = true; });
  }
  maxRow = Math.max(maxRow, baseSheet.getLastRow());

  var rcSheet = getRCSheet();
  if (rcSheet.getLastRow() > 1) {
    var rcIds = rcSheet.getRange(2, IDX_RC.TICKET_ID + 1, rcSheet.getLastRow() - 1, 1).getValues();
    rcIds.forEach(function(r) { if (r[0]) existing[String(r[0]).trim().toUpperCase()] = true; });
  }
  maxRow = Math.max(maxRow, rcSheet.getLastRow());

  var seq = maxRow;
  var candidate, key;
  do {
    var s = String(seq);
    while (s.length < 4) s = '0' + s;
    candidate = 'CAM-' + s;
    key = candidate.toUpperCase();
    seq++;
  } while (existing[key]);
  return candidate;
}

// ============================================================
// createCampaignCSV
// ============================================================
var CSV_FOLDER_ID = '1pZ03_RgDTVtudaFDAmc8vK6-0k4qNLNR';

var CSV_HEADERS = [
  'country', 'owner', 'fecha_inicio', 'fecha_fin', 'hora_inicio', 'hora_fin',
  'tipo_oferta', 'segmentacion', 'discount', 'minimo_compra', 'maximo_descuento',
  'vertical', 'squad', 'strategy', 'descripcion', 'pct_alianza', 'id_alianza',
  'id_orden_compra', 'store_type', 'brand_id', 'store_ids', 'product_id',
  'exclude_store_ids', 'max_unidades_orden', 'days_hours', 'brand_name',
  'metodos_pago', 'cc_type', 'bin', 'max_ordenes_usuario', 'presupuesto',
  'tipo_usuario_app', 'user_not_in_tag', 'terminos_y_condiciones', 'prime', 'is_deal_of_the_day',
  'is_offer_on_top', 'max_quantity_global', 'send_products_to_braze',
  'cashback_days_to_end', 'store_types_redencion', 'fecha_inicio_redencion',
  'fecha_fin_redencion', 'hora_inicio_redencion', 'hora_fin_redencion',
  'stores_redencion_si', 'stores_redencion_no', 'item'
];

var CSV_IDX_MAP = [
  IDX.COUNTRY, IDX.OWNER, IDX.FECHA_INICIO, IDX.FECHA_FIN,
  IDX.HORA_INICIO, IDX.HORA_FIN, IDX.TIPO_OFERTA, IDX.SEGMENTACION,
  IDX.DISCOUNT, IDX.MINIMO_COMPRA, IDX.MAXIMO_DESCUENTO,
  IDX.VERTICAL, IDX.SQUAD, IDX.STRATEGY, IDX.DESCRIPCION,
  IDX.PCT_ALIANZA, IDX.ID_ALIANZA, IDX.ID_ORDEN_COMPRA,
  IDX.STORE_TYPE, IDX.BRAND_ID, IDX.STORE_IDS, IDX.PRODUCT_ID,
  IDX.EXCLUDE_STORE_IDS, IDX.MAX_UNIDADES_ORDEN, IDX.DAYS_HOURS,
  IDX.BRAND_NAME, IDX.METODOS_PAGO, IDX.CC_TYPE, IDX.BIN,
  IDX.MAX_ORDENES_USUARIO, IDX.PRESUPUESTO, IDX.TIPO_USUARIO_APP,
  IDX.USER_NOT_IN_TAG, IDX.TERMINOS_Y_CONDICIONES, IDX.PRIME, IDX.IS_DEAL_OF_THE_DAY,
  IDX.IS_OFFER_ON_TOP, IDX.MAX_QUANTITY_GLOBAL, IDX.SEND_PRODUCTS_TO_BRAZE,
  IDX.CASHBACK_DAYS_TO_END, IDX.STORE_TYPES_REDENCION,
  IDX.FECHA_INICIO_REDENCION, IDX.FECHA_FIN_REDENCION,
  IDX.HORA_INICIO_REDENCION, IDX.HORA_FIN_REDENCION,
  IDX.STORES_REDENCION_SI, IDX.STORES_REDENCION_NO, IDX.ITEM
];

function csvEscape(val) {
  var s = String(val === null || val === undefined ? '' : val);
  if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) {
    s = '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function createCampaignCSV(ticketId, rowData) {
  try {
    var folder = DriveApp.getFolderById(CSV_FOLDER_ID);
    var values = CSV_IDX_MAP.map(function(idx) {
      return csvEscape(rowData[idx] || '');
    });
    var csvContent = CSV_HEADERS.join(',') + '\n' + values.join(',');
    var fileName = ticketId + '.csv';
    var blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
    folder.createFile(blob);
    Logger.log('CSV creado: ' + fileName);
  } catch(e) {
    Logger.log('Error creando CSV: ' + e.message);
  }
}

// ============================================================
// uploadRappiCreditosUsersFile — sube archivo o genera CSV
//                                 desde ingreso manual
// ============================================================
function uploadRappiCreditosUsersFile(ticketId, data) {
  var folder = DriveApp.getFolderById(CSV_FOLDER_ID);
  var blob;

  if (data.usersMode === 'file' && data.usersFileBase64) {
    var bytes = Utilities.base64Decode(data.usersFileBase64);
    var fileName = ticketId + '_' + (data.usersFileName || 'usuarios');
    blob = Utilities.newBlob(bytes, data.usersFileMime || 'text/csv', fileName);
  } else {
    var rows = data.userRows || [];
    var lines = ['user_id,monto'];
    rows.forEach(function(r) {
      lines.push(csvEscape(r.userId) + ',' + csvEscape(r.monto));
    });
    blob = Utilities.newBlob(lines.join('\n'), 'text/csv', ticketId + '_usuarios.csv');
  }

  var file = folder.createFile(blob);
  return file.getUrl();
}

// ============================================================
// createRappiCreditosCSV
// ============================================================
var CSV_HEADERS_RC = [
  'country','owner','fecha_inicio','fecha_fin','hora_inicio','hora_fin',
  'squad','strategy','pago','pct_alianza','id_alianza','id_orden_compra',
  'budget','users_file_url','tope_pedido','vigencia_dias','vigencia_fecha',
  'store_ids','descripcion'
];

var CSV_IDX_MAP_RC = [
  IDX_RC.COUNTRY, IDX_RC.OWNER, IDX_RC.FECHA_INICIO, IDX_RC.FECHA_FIN,
  IDX_RC.HORA_INICIO, IDX_RC.HORA_FIN, IDX_RC.SQUAD, IDX_RC.STRATEGY,
  IDX_RC.PAGO, IDX_RC.PCT_ALIANZA, IDX_RC.ID_ALIANZA, IDX_RC.ID_ORDEN_COMPRA,
  IDX_RC.BUDGET, IDX_RC.USERS_FILE_URL, IDX_RC.TOPE_PEDIDO,
  IDX_RC.VIGENCIA_DIAS, IDX_RC.VIGENCIA_FECHA, IDX_RC.STORE_IDS, IDX_RC.DESCRIPCION
];

function createRappiCreditosCSV(ticketId, rowData) {
  try {
    var folder = DriveApp.getFolderById(CSV_FOLDER_ID);
    var values = CSV_IDX_MAP_RC.map(function(idx) {
      return csvEscape(rowData[idx] || '');
    });
    var csvContent = CSV_HEADERS_RC.join(',') + '\n' + values.join(',');
    var fileName = ticketId + '_resumen.csv';
    var blob = Utilities.newBlob(csvContent, 'text/csv', fileName);
    folder.createFile(blob);
    Logger.log('CSV RappiCreditos creado: ' + fileName);
  } catch(e) {
    Logger.log('Error creando CSV RappiCreditos: ' + e.message);
  }
}

// ============================================================
// notifyRappiCreditosSlack — mensaje inicial al crear solicitud RC
// ============================================================
function notifyRappiCreditosSlack(ticketId, email, data, usersFileUrl) {
  try {
    var countryMap = {
      'Colombia':'co','México':'mx','Argentina':'ar','Chile':'cl',
      'Perú':'pe','Uruguay':'uy','Ecuador':'ec','Costa Rica':'cr'
    };
    var countryCode = countryMap[data.pais] || (data.pais || '').toLowerCase();
    var paisFlags = { co:'🇨🇴', mx:'🇲🇽', ar:'🇦🇷', cl:'🇨🇱', pe:'🇵🇪', uy:'🇺🇾', ec:'🇪🇨', cr:'🇨🇷' };
    var flag = paisFlags[countryCode] || '🌎';
    var budget = data.budget ? Number(data.budget).toLocaleString() : '—';
    var userCount = data.usersMode === 'manual' ? (data.userRows || []).length : null;

    var mainRes = slackApi('chat.postMessage', {
      channel: SLACK_CHANNEL,
      text: flag + ' [' + ticketId + '] RAPPICREDITOS · ' + (data.ownerEmail || email),
      blocks: [{
        type: 'section',
        text: { type: 'mrkdwn',
          text: flag + ' *RAPPICREDITOS*\n' +
                '*Ticket:* `' + ticketId + '`   *Budget:* ' + budget + '\n' +
                '*Fechas:* ' + (data.fechaInicio||'—') + ' → ' + (data.fechaFin||'—') + '\n' +
                '*Solicitante:* ' + (data.ownerEmail||email)
        }
      }]
    });

    if (!mainRes.ok) { Logger.log('Slack error: ' + mainRes.error); return null; }
    var threadTs = mainRes.ts;
    saveThreadTs(ticketId, threadTs);

    var lines = [];
    lines.push('*🎁 Detalle — ' + ticketId + '*');
    lines.push('');
    lines.push('> País: ' + flag + ' ' + (data.pais || countryCode));
    lines.push('> Owner: ' + (data.ownerEmail || email));
    lines.push('> Fechas: ' + (data.fechaInicio||'—') + ' → ' + (data.fechaFin||'—'));
    lines.push('> Usuarios: ' + (userCount !== null ? userCount + ' (ingreso manual)' : 'ver archivo adjunto'));
    if (usersFileUrl) lines.push('> Archivo: ' + usersFileUrl);
    if (data.topePedido) lines.push('> Tope por pedido: ' + data.topePedido);
    if (data.vigenciaDias) lines.push('> Vigencia: ' + data.vigenciaDias + ' días desde la carga');
    else if (data.vigenciaFecha) lines.push('> Vigencia: hasta ' + data.vigenciaFecha);
    lines.push('> Cobertura de redención: ' + (data.storeIds ? 'Store IDs: ' + data.storeIds.substring(0,200) : 'Abierto'));
    if (data.squad) lines.push('> Squad: ' + data.squad);
    if (data.strategy) lines.push('> Strategy: ' + data.strategy);
    var pagoLabel = {full_rappi:'Full Rappi', full_aliado:'Full Aliado'}[data.pago] || data.pago || '—';
    lines.push('> Esquema de pago: ' + pagoLabel);
    if (data.idAlianza) lines.push('> ID Alianza: ' + data.idAlianza);
    if (data.ordenCompra) lines.push('> Orden Compra: ' + data.ordenCompra);
    if (data.descripcion) { lines.push(''); lines.push('*💬 Descripción*'); lines.push('> ' + data.descripcion); }

    slackApi('chat.postMessage', {
      channel: SLACK_CHANNEL,
      thread_ts: threadTs,
      text: 'Detalle ' + ticketId,
      blocks: [{ type:'section', text:{ type:'mrkdwn', text: lines.join('\n') } }]
    });

    var reviewerMsg = getReviewerMessage(countryCode);
    if (reviewerMsg) {
      slackApi('chat.postMessage', { channel: SLACK_CHANNEL, thread_ts: threadTs, text: reviewerMsg, mrkdwn: true });
    }

    return threadTs;
  } catch(e) { Logger.log('Slack RappiCreditos error: ' + e.message); return null; }
}

// ============================================================
// saveRappiCreditosRequest
// ============================================================
function saveRappiCreditosRequest(data) {
  try {
    var sheet    = getRCSheet();
    var ticketId = generateTicketId();
    var fecha    = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    var email = '';
    try { email = Session.getActiveUser().getEmail(); } catch(ex) {}
    if (!email) email = data.ownerEmail || 'sin-email';

    var countryMap = {
      'Colombia':'co','México':'mx','Argentina':'ar','Chile':'cl',
      'Perú':'pe','Uruguay':'uy','Ecuador':'ec','Costa Rica':'cr'
    };
    var countryCode = countryMap[data.pais] || (data.pais || '').toLowerCase();

    var usersFileUrl = uploadRappiCreditosUsersFile(ticketId, data);

    var row = new Array(HEADERS_RC.length).fill('');
    row[IDX_RC.TICKET_ID]         = ticketId;
    row[IDX_RC.FECHA_SOLICITUD]   = fecha;
    row[IDX_RC.EMAIL_SOLICITANTE] = email;
    row[IDX_RC.COUNTRY]           = countryCode;
    row[IDX_RC.OWNER]             = data.ownerEmail        || '';
    row[IDX_RC.FECHA_INICIO]      = data.fechaInicio       || '';
    row[IDX_RC.FECHA_FIN]         = data.fechaFin          || '';
    row[IDX_RC.HORA_INICIO]       = data.horaInicio        || '';
    row[IDX_RC.HORA_FIN]          = data.horaFin           || '';
    row[IDX_RC.SQUAD]             = data.squad             || '';
    row[IDX_RC.STRATEGY]          = data.strategy          || '';
    row[IDX_RC.PAGO]              = data.pago              || '';
    row[IDX_RC.PCT_ALIANZA]       = data.pctAlianza        || '';
    row[IDX_RC.ID_ALIANZA]        = data.idAlianza         || '';
    row[IDX_RC.ID_ORDEN_COMPRA]   = data.ordenCompra ? String(data.ordenCompra) : '';
    row[IDX_RC.BUDGET]            = data.budget            || '';
    row[IDX_RC.USERS_FILE_URL]    = usersFileUrl            || '';
    row[IDX_RC.TOPE_PEDIDO]       = data.topePedido        || '';
    row[IDX_RC.VIGENCIA_DIAS]     = data.vigenciaDias      || '';
    row[IDX_RC.VIGENCIA_FECHA]    = data.vigenciaFecha     || '';
    row[IDX_RC.STORE_IDS]         = data.storeIds          || '';
    row[IDX_RC.DESCRIPCION]       = data.descripcion       || '';
    row[IDX_RC.STATUS]            = 'Pendiente';
    row[IDX_RC.COMENTARIOS]       = '';
    row[IDX_RC.GLOBAL_OFFER_ID]   = '';
    row[IDX_RC.THREAD_TS]         = '';

    sheet.appendRow(row);
    var newRow = sheet.getLastRow();

    sheet.getRange(newRow, IDX_RC.STATUS + 1)
      .setBackground('#FEF3C7').setFontWeight('bold').setHorizontalAlignment('center');

    try { createRappiCreditosCSV(ticketId, row); } catch(e) { Logger.log('CSV RC error: ' + e.message); }

    try {
      var threadTs = notifyRappiCreditosSlack(ticketId, email, data, usersFileUrl);
      if (threadTs) {
        var tsCell = sheet.getRange(newRow, IDX_RC.THREAD_TS + 1);
        tsCell.setNumberFormat('@');
        tsCell.setValue("'" + String(threadTs));
      }
    } catch(e) { Logger.log('Slack RC notify error: ' + e.message); }

    return { success: true, ticketId: ticketId, fecha: fecha, email: email };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================================
// getAllRappiCreditosTickets
// ============================================================
function getAllRappiCreditosTickets() {
  try {
    var sheet = openRCSheet();
    if (sheet.getLastRow() < 2) return [];
    var data   = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || !row[IDX_RC.TICKET_ID]) continue;
      function safe(idx) { return row.length > idx ? String(row[idx] || '') : ''; }
      result.push({
        ticketId:      safe(IDX_RC.TICKET_ID),
        fecha:         safe(IDX_RC.FECHA_SOLICITUD),
        email:         safe(IDX_RC.EMAIL_SOLICITANTE),
        pais:          safe(IDX_RC.COUNTRY),
        owner:         safe(IDX_RC.OWNER),
        squad:         safe(IDX_RC.SQUAD),
        strategy:      safe(IDX_RC.STRATEGY),
        budget:        safe(IDX_RC.BUDGET),
        usersFileUrl:  safe(IDX_RC.USERS_FILE_URL),
        topePedido:    safe(IDX_RC.TOPE_PEDIDO),
        vigenciaDias:  safe(IDX_RC.VIGENCIA_DIAS),
        vigenciaFecha: safe(IDX_RC.VIGENCIA_FECHA),
        storeIds:      safe(IDX_RC.STORE_IDS),
        status:        safe(IDX_RC.STATUS),
        comentarios:   safe(IDX_RC.COMENTARIOS),
        globalOfferId: safe(IDX_RC.GLOBAL_OFFER_ID)
      });
    }
    result.reverse();
    return result;
  } catch (err) {
    Logger.log('getAllRappiCreditosTickets error: ' + err.message);
    return [];
  }
}
