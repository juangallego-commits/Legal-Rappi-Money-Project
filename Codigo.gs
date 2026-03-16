// =================================================================
// MOTOR LEGAL RAPPI - CORE ENGINE (V2/V3)
// =================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('WebApp')
      .evaluate()
      .setTitle('Motor Legal Rappi')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function processWebPayload(payloadString) {
  try {
    const payload = JSON.parse(payloadString);
    let activeEmail = payload.userEmail;
    try { activeEmail = Session.getActiveUser().getEmail() || payload.userEmail; } catch(e) {}
      
    Logger.log('📥 Tipo de campaña: ' + payload.dynamicType);
    const result = coreEngineV2(payload, activeEmail);
    return JSON.stringify({ status: 'success', docUrl: result.docUrl, docName: result.docName });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message, stack: e.stack });
  }
}

// -----------------------------------------------------------------
// ENRUTADOR V3 (TEMPLATE ENGINE)
// -----------------------------------------------------------------
function coreEngineV2(payload, submitterEmail) {
  const vars = mapWebToEngine(payload);
  const countryCode = payload.countryCode || 'CO';
  const campaignType = vars['Tipo de Dinámica'] || 'Cashback';
  const vertical = payload.vertical || 'ALL';
  
  try {
    const typeConfig = _getCampaignTypeConfig(campaignType);
    const registry = _getTemplateRegistry();
    let template = registry.find(r => r.country_code === countryCode && r.campaign_type === campaignType && r.status === 'active' && (r.vertical === vertical || r.vertical === 'ALL' || !r.vertical));
    
    if (template) {
      if (typeConfig && typeConfig.processing_mode === 'template_only') {
        // V3.3: Smart Template con fallback a legacy si falla
        try {
          return _generateSmartTemplate(template, payload, vars, submitterEmail);
        } catch (smartErr) {
          Logger.log('⚠️ Smart Template falló, intentando legacy: ' + smartErr.message);
          // Solo intentar legacy si el tipo tiene lógica legacy (Cashback o Concurso)
          var campaignTypeName = vars['Tipo de Dinámica'] || campaignType;
          if (campaignTypeName === 'Cashback' || campaignTypeName === 'Concurso Mayor Comprador') {
            Logger.log('🔄 Fallback a _generateWithTemplate para: ' + campaignTypeName);
            return _generateWithTemplate(template, vars, submitterEmail);
          }
          // Si no es legacy-compatible, propagar el error original
          throw smartErr;
        }
      } else {
        return _generateWithTemplate(template, vars, submitterEmail);
      }
    }
  } catch (e) { Logger.log('⚠️ Engine falló: ' + e.message); }
  
  throw new Error("No se encontró una plantilla de documento activa para este tipo de dinámica en este país.");
}
// ---INICIO COPIAR---
// -----------------------------------------------------------------
// GENERADOR SMART (TEMPLATE_ONLY) — Dinámicas importadas via Wizard
// -----------------------------------------------------------------
function _generateSmartTemplate(template, payload, vars, submitterEmail) {
  // 1. Obtener campos definidos para esta combinación tipo+país
  var fieldsSheet = _getSheet(FIELDS_SHEET_NAME);
  if (!fieldsSheet) throw new Error('Sheet Template_Fields no encontrada.');

  var allFields = _sheetToObjects(fieldsSheet);
  var campaignType = vars['Tipo de Dinámica'] || template.campaign_type;
  var countryCode = payload.countryCode || template.country_code || 'CO';

  var relevantFields = allFields.filter(function(f) {
    var matchType = (f.campaign_type === 'ALL' || f.campaign_type === campaignType);
    var matchCountry = (f.country_code === 'ALL' || f.country_code === countryCode);
    return matchType && matchCountry;
  });

  // 2. Construir mapa de placeholders cruzando payload con Template_Fields
  var placeholders = {};

  relevantFields.forEach(function(field) {
    var fieldId = field.field_id;
    var placeholder = field.placeholder; // Ej: "{{FECHA_INICIO}}"
    if (!placeholder) return;

    // V3.2: Si tiene canonical_field_id, buscar primero en el campo base del payload
    var value = '';
    var canonicalId = (field.canonical_field_id || '').toString().trim();

    if (canonicalId && payload[canonicalId] !== undefined && payload[canonicalId] !== '') {
      // Señal principal: este placeholder corresponde a un campo base del formulario
      value = String(payload[canonicalId]);
    } else if (payload[fieldId] !== undefined && payload[fieldId] !== null) {
      value = String(payload[fieldId]);
    } else if (payload.dynamicFields && payload.dynamicFields[fieldId] !== undefined) {
      value = String(payload.dynamicFields[fieldId]);
    } else if (vars[field.label_es] !== undefined) {
      value = String(vars[field.label_es]);
    }

    // Aplicar formateo si tiene format_as
    if (value && field.format_as) {
      value = _applySmartFormat(value, field.format_as, countryCode);
    }

    var cleanKey = placeholder.replace(/^\{\{/, '').replace(/\}\}$/, '');
    placeholders['{{' + cleanKey + '}}'] = value;
  });
  // ──── V3.3: RESOLVER CAMPOS DERIVADOS ────
  // Estos son placeholders compuestos que se calculan del payload
  // (ej: TOPE_LETRAS, TEXTO_PORCENTAJE, UMBRAL_NUM, etc.)
  if (typeof DERIVED_FIELDS !== 'undefined') {
    var derivedKeys = Object.keys(DERIVED_FIELDS);
    for (var dk = 0; dk < derivedKeys.length; dk++) {
      var dKey = derivedKeys[dk];
      var phDerived = '{{' + dKey + '}}';
      // V3.3: DERIVED_FIELDS siempre sobreescribe — es la fuente autoritativa
      // para formato complejo (ej: "veinte por ciento (20%)" vs "20%")
      {
        try {
          var derivedVal = DERIVED_FIELDS[dKey](payload);
          if (derivedVal !== null && derivedVal !== undefined && derivedVal !== '') {
            placeholders[phDerived] = String(derivedVal);
          }
        } catch(de) {
          Logger.log('⚠️ Derived field ' + dKey + ': ' + de.message);
        }
      }
    }
  }

  // ──── V3.3: RESOLVER LEGAL DEFAULTS ────
  // Campos como jurisdicción, ley aplicable, se resuelven de Country_Settings
  if (typeof LEGAL_DEFAULTS_MAP !== 'undefined') {
    try {
      var csSheet = _getSheet(COUNTRY_SETTINGS_SHEET);
      if (csSheet) {
        var csData = _sheetToObjects(csSheet);
        var csRow = csData.find(function(c) { return c.country_code === countryCode; });
        if (csRow) {
          var legalKeys = Object.keys(LEGAL_DEFAULTS_MAP);
          for (var lk = 0; lk < legalKeys.length; lk++) {
            var lKey = legalKeys[lk];
            var phLegal = '{{' + lKey + '}}';
            if (!placeholders[phLegal] || placeholders[phLegal] === '') {
              var colName = LEGAL_DEFAULTS_MAP[lKey].column;
              if (csRow[colName]) {
                placeholders[phLegal] = String(csRow[colName]);
              }
            }
          }
        }
      }
    } catch(le) {
      Logger.log('⚠️ Legal defaults: ' + le.message);
    }
  }
  // 3. Nombre del documento
  var campaignName = payload.campaignName || vars['Nombre de Campaña (Opcional)'] || campaignType;
  var shopName = payload.shopName || vars['Tienda Participante'] || 'General';
  var today = new Date();
  var dateStr = today.getFullYear() + '-' +
    String(today.getMonth() + 1).padStart(2, '0') + '-' +
    String(today.getDate()).padStart(2, '0');
  var docName = 'T&C ' + shopName + ' - ' + campaignName + ' (' + dateStr + ')';

  // 4. Clonar el template Doc
  var newFile = DriveApp.getFileById(template.template_doc_id).makeCopy(docName);

  if (DRIVE_FOLDER_ID) {
    try {
      DriveApp.getFolderById(DRIVE_FOLDER_ID).addFile(newFile);
      DriveApp.getRootFolder().removeFile(newFile);
    } catch(e) {}
  }

  // 5. Abrir y hacer reemplazos
  var doc = DocumentApp.openById(newFile.getId());
  var body = doc.getBody();

  // 5a. Reemplazar placeholders con valor
  Object.keys(placeholders).forEach(function(key) {
    var value = placeholders[key];
    if (value !== null && value !== undefined && value !== '') {
      var escaped = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      body.replaceText(escaped, value);
    }
  });

  // 5b. SMART DELETION: Viñetas vacías (solo placeholder)
  var listItems = body.getListItems();
  for (var i = listItems.length - 1; i >= 0; i--) {
    if (listItems[i].getText().trim().match(/^\{\{[A-Z_0-9]+\}\}$/)) {
      listItems[i].removeFromParent();
    }
  }

  // 5c. SMART DELETION: Párrafos vacíos (solo placeholder)
  var paragraphs = body.getParagraphs();
  for (var j = paragraphs.length - 1; j >= 0; j--) {
    if (paragraphs[j].getText().trim().match(/^\{\{[A-Z_0-9]+\}\}$/)) {
      paragraphs[j].removeFromParent();
    }
  }

  // 5d. Limpiar placeholders sueltos residuales
  body.replaceText('\\{\\{[A-Z_0-9_]+\\}\\}', '');
  doc.saveAndClose();

  // 6. Permisos + tracking
  var publicUrl = setPublicViewPermissions(doc);

  var auditVars = {
    'Tipo de Dinámica': campaignType,
    'Tienda Participante': shopName,
    'Nombre de Campaña (Opcional)': campaignName,
    'Email Generador': submitterEmail
  };
  relevantFields.forEach(function(f) {
    var cleanKey = (f.placeholder || '').replace(/^\{\{/, '').replace(/\}\}$/, '');
    if (cleanKey && placeholders['{{' + cleanKey + '}}']) {
      auditVars[f.label_es || f.field_id] = placeholders['{{' + cleanKey + '}}'];
    }
  });

  try { saveResponseToSheet(auditVars, publicUrl); } catch(e) { Logger.log('⚠️ Audit: ' + e.message); }

  // 7. Email
  try {
    sendEmailNotification(submitterEmail, docName, publicUrl, {
      docName: docName,
      dinamica: campaignType,
      tiendaDisplay: shopName,
      textoVigenciaEmail: (payload.startDate || '') + ' → ' + (payload.endDate || ''),
      condicionesEspeciales: payload.specialConditions || ''
    });
  } catch(e) { Logger.log('⚠️ Email: ' + e.message); }

  return { docUrl: publicUrl, docName: docName };
}

// -----------------------------------------------------------------
// FORMATO INTELIGENTE para campos de templates importados
// -----------------------------------------------------------------
function _applySmartFormat(value, formatType, countryCode) {
  try {
    switch(formatType) {
      case 'date_legal':
        var parts = value.split('-');
        if (parts.length === 3) {
          var day = parseInt(parts[2]);
          var monthIdx = parseInt(parts[1]) - 1;
          if (monthIdx >= 0 && monthIdx < 12) {
            return day + ' de ' + MESES_ES[monthIdx] + ' de ' + parts[0];
          }
        }
        return value;

      case 'money':
        var symbol = '$';
        try {
          var cSheet = _getSheet(COUNTRY_SETTINGS_SHEET);
          if (cSheet) {
            var cData = _sheetToObjects(cSheet);
            var cc = cData.find(function(c) { return c.country_code === countryCode; });
            if (cc) symbol = cc.currency_symbol || '$';
          }
        } catch(e) {}
        var num = parseFloat(value);
        if (!isNaN(num)) return symbol + num.toLocaleString('es-CO');
        return value;

      case 'percentage':
        var pct = parseFloat(value);
        if (!isNaN(pct)) return pct + '%';
        return value;

      case 'number_words':
        var numWords = ['cero','uno','dos','tres','cuatro','cinco','seis','siete',
                        'ocho','nueve','diez','once','doce','trece','catorce','quince',
                        'dieciséis','diecisiete','dieciocho','diecinueve','veinte'];
        var n = parseInt(value);
        if (!isNaN(n) && n >= 0 && n <= 20) return numWords[n] + ' (' + n + ')';
        return value;

      default:
        return value;
    }
  } catch(e) { return value; }
}
// -----------------------------------------------------------------
// GENERADOR CON PLANTILLA
// -----------------------------------------------------------------
function _generateWithTemplate(template, vars, submitterEmail) {
  vars['Email Generador'] = submitterEmail;
  validateDates(vars);
  
  const tipoDinamica = vars['Tipo de Dinámica'] || 'Cashback';
  let data = tipoDinamica === 'Concurso Mayor Comprador' ? procesarConcurso(vars) : procesarCashback(vars);
  auditData(data);
  
  const placeholders = _buildPlaceholderMap(data);
  const newFile = DriveApp.getFileById(template.template_doc_id).makeCopy(data.docName);
  
  if (DRIVE_FOLDER_ID) {
    try { DriveApp.getFolderById(DRIVE_FOLDER_ID).addFile(newFile); DriveApp.getRootFolder().removeFile(newFile); } catch(e){}
  }
  
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();
  
  // 1. Reemplazar variables que SÍ tienen contenido
  Object.entries(placeholders).forEach(([key, value]) => {
    if (value !== null && value !== undefined && value !== '') {
      body.replaceText(key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), String(value));
    }
  });
  
  // 2. SMART DELETION: Eliminar viñetas enteras que quedaron con variables vacías
  const listItems = body.getListItems();
  for (let i = listItems.length - 1; i >= 0; i--) {
    if (listItems[i].getText().trim().match(/^\{\{[A-Z_0-9_]+\}\}$/)) {
      listItems[i].removeFromParent();
    }
  }
  
  // 3. SMART DELETION: Eliminar párrafos enteros que quedaron con variables vacías
  const paragraphs = body.getParagraphs();
  for (let i = paragraphs.length - 1; i >= 0; i--) {
    if (paragraphs[i].getText().trim().match(/^\{\{[A-Z_0-9_]+\}\}$/)) {
      paragraphs[i].removeFromParent();
    }
  }
  
  // 4. Limpiar cualquier otro placeholder que haya quedado suelto
  body.replaceText('\\{\\{[A-Z_0-9_]+\\}\\}', '');
  doc.saveAndClose();
  
  const publicUrl = setPublicViewPermissions(doc);
  saveResponseToSheet(vars, publicUrl);
  sendEmailNotification(submitterEmail, data.docName, publicUrl, data);
  return { docUrl: publicUrl, docName: data.docName };
}

// -----------------------------------------------------------------
// MAPEOS Y PROCESAMIENTO INTELIGENTE
// -----------------------------------------------------------------
function mapWebToEngine(webData) {
  return {
    'Tipo de Dinámica': webData.dynamicType, 'Tienda Participante': webData.shopName, 'Territorio': webData.territory,
    'Fecha de INICIO de Campaña': webData.startDate, 'Hora de INICIO de Campaña': webData.startTime,
    'Fecha de FIN de Campaña': webData.endDate, 'Hora de FIN de Campaña': webData.endTime,
    'Nombre de Campaña (Opcional)': webData.campaignName, 'Email(s) adicionales': webData.extraEmails,
    'Condiciones Especiales (Adicionales)': webData.specialConditions, 'Tipo de Usuarios Participantes': webData.userSegment,
    'Métodos de Pago Válidos': webData.paymentMethods, 'Máximo de Órdenes por Usuario': webData.maxOrders,
    'Valor Mínimo de Compra (Opcional)': webData.minPurchase, 'Lugar de Redención de Créditos': webData.redemptionPlace,
    'Porcentaje del Cashback': webData.cashbackPct, 'Presupuesto (Existencias)': webData.budget,
    'Tope Máximo de Cashback': webData.cap, 'Tipo de Carga de Créditos': webData.loadType,
    'Fecha de Carga de Créditos': webData.loadDate, '¿Cómo se define la vigencia de los Créditos?': webData.validityType,
    'Cantidad de Días de Vigencia': webData.validityDays, 'Fecha Inicio de Redención': webData.redemptionStart,
    'Fecha Fin de Redención': webData.redemptionEnd, 'Razón Social Organizador': webData.organizerLegalName,
    'Teléfono Contacto Organizador': webData.organizerPhone, 'Email Contacto Organizador': webData.organizerEmail,
    'Número de Ganadores': webData.numberOfWinners, 'Tipo de Premio': webData.prizeType,
    'Quién Entrega Premio': webData.prizeDeliveryBy, 'Monto Créditos Premio': webData.creditsAmount,
    'Días para Carga Premio': webData.creditLoadDays, 'Vigencia Créditos Premio': webData.creditsValidityDays,
    'Lugar Redención Premio': webData.creditsRedemptionPlace, 'Descripción Premio Físico': webData.physicalPrizeDescription,
    'Verticales/Secciones participantes': webData.verticals, 'Productos Participantes': webData.participatingProducts,
    'Criterio del Ganador': webData.winnerCriteria, 'Mínimo de Compra para participar': webData.minParticipation,
    'Lista de Premios': webData.prizes, 'Fecha de Anuncio de Ganadores': webData.announcementDate,
    'País Seleccionado': webData.countryCode, 'Código de Moneda': webData.currencyCode, 'Símbolo de Moneda': webData.currencySymbol
  };
}

function processCommonVariables(vars) {
  const data = {};
  data.tiendaBase = capitalize(vars['Tienda Participante'] ? vars['Tienda Participante'].trim() : "TIENDA");
  data.marcaAliada = data.tiendaBase;

  let rawTerritorio = vars['Territorio'];
  let territoriosSeleccionados = Array.isArray(rawTerritorio) ? rawTerritorio : (typeof rawTerritorio === 'string' ? rawTerritorio.split(',').map(s => s.trim()) : [rawTerritorio]);
  
  if (territoriosSeleccionados.includes('Nacional')) {
    data.textoTerritorio = 'el territorio nacional de la República de Colombia';
    data.territorioResumen = 'Nacional';
  } else {
    const listaFinalStr = formatListToText(territoriosSeleccionados);
    data.territorioResumen = listaFinalStr;
    let hayMunicipios = false;
    let hayCiudades = false;
    
    territoriosSeleccionados.forEach(lugar => {
      if (typeof LISTA_MUNICIPIOS !== 'undefined' && LISTA_MUNICIPIOS.includes(lugar)) {
        hayMunicipios = true;
      } else {
        hayCiudades = true;
      }
    });
    
    let prefix = "";
    if (hayCiudades && hayMunicipios) {
      prefix = "las ciudades y municipios de";
    } else if (hayMunicipios && !hayCiudades) {
      prefix = (territoriosSeleccionados.length > 1 || listaFinalStr.includes(' y ')) ? "los municipios de" : "el municipio de";
    } else {
      prefix = (territoriosSeleccionados.length > 1 || listaFinalStr.includes(' y ')) ? "las ciudades de" : "la ciudad de";
    }
    data.textoTerritorio = `${prefix} ${listaFinalStr}, dentro de la República de Colombia`;
  }

  const startDateStr = vars['Fecha de INICIO de Campaña'];
  const startTimeStr = vars['Hora de INICIO de Campaña'];
  const endDateStr = vars['Fecha de FIN de Campaña'];
  const endTimeStr = vars['Hora de FIN de Campaña'];
  
  const startDate = new Date(parseInt(startDateStr.split('-')[0]), parseInt(startDateStr.split('-')[1]) - 1, parseInt(startDateStr.split('-')[2]), parseInt(startTimeStr.split(':')[0]), parseInt(startTimeStr.split(':')[1]));
  const endDate = new Date(parseInt(endDateStr.split('-')[0]), parseInt(endDateStr.split('-')[1]) - 1, parseInt(endDateStr.split('-')[2]), parseInt(endTimeStr.split(':')[0]), parseInt(endTimeStr.split(':')[1]));
  
  data.startFmtDate = formatDateInSpanish(startDate); data.endFmtDate = formatDateInSpanish(endDate);
  data.startFmtTime = formatTimeInSpanish(startDate); data.endFmtTime = formatTimeInSpanish(endDate);
  data.isSameDay = (data.startFmtDate === data.endFmtDate);
  data.textoVigenciaEmail = data.isSameDay ? `El ${data.startFmtDate} (${data.startFmtTime} - ${data.endFmtTime})` : `Del ${data.startFmtDate} al ${data.endFmtDate}`;
  data.condicionesEspeciales = vars['Condiciones Especiales (Adicionales)'] || '';
  return data;
}

function procesarCashback(vars) {
  const data = processCommonVariables(vars);
  data.dinamica = 'CASHBACK';
  
  const redencionSeleccionada = vars['Lugar de Redención de Créditos'] || 'Únicamente en la Tienda Participante (Brand Credits)';
  const isGeneric = (vars['Tienda Participante'].toUpperCase().startsWith('TODAS') || vars['Tienda Participante'].toUpperCase().startsWith('TODOS'));

  if (isGeneric) {
    data.tiendaDisplay = "Aliados Comerciales"; 
    data.txtDefinicionTienda = ' (en adelante las "Tiendas Participantes") ';
    data.txtRefTienda = 'las Tiendas Participantes';
    data.txtDeclaracionTienda = 'son las Tiendas Participantes';
  } else {
    data.tiendaDisplay = `"${data.tiendaBase}"`;
    data.txtDefinicionTienda = ' (en adelante la "Tienda Participante") ';
    data.txtRefTienda = 'la Tienda Participante';
    data.txtDeclaracionTienda = 'es la Tienda Participante';
  }

  data.segmento = vars['Tipo de Usuarios Participantes'];
  data.metodosPago = vars['Métodos de Pago Válidos'];
  data.extraEmails = vars['Email(s) adicionales'];
  data.limiteOrdenes = vars['Máximo de Órdenes por Usuario'] || '1';
  const ordenNum = parseInt(data.limiteOrdenes);
  data.txtOrdenesGramatica = (!isNaN(ordenNum) && ordenNum === 1) ? 'orden' : 'órdenes';
  
  const pctNum = Number(vars['Porcentaje del Cashback'] || '99'); 
  data.textoPorcentaje = `${numeroALetras(Math.floor(pctNum))} por ciento (${pctNum}%)`;

  const topeNum = Number(vars['Tope Máximo de Cashback'] || 0);
  data.topeLetras = numeroALetras(topeNum);
  data.topeNumFmt = topeNum.toLocaleString('es-CO');
  data.presupuestoNumFmt = Number(vars['Presupuesto (Existencias)'] || 0).toLocaleString('es-CO');
  
  let umbralCompraNum = (pctNum > 0) ? Math.ceil(topeNum / (pctNum / 100)) : topeNum;
  data.umbralCompraFmt = umbralCompraNum.toLocaleString('es-CO');
  data.umbralCompraLetras = numeroALetras(umbralCompraNum); 
  
  const tipoCarga = vars['Tipo de Carga de Créditos'];
  const objFechaCarga = parseFormDate(vars['Fecha de Carga de Créditos']);
  let fechaCargaFmt = formatDateInSpanish(objFechaCarga); 
  if (tipoCarga === 'Inmediatamente (Al finalizar la orden)') {
    data.txtCargaCompleto = 'de manera inmediata una vez finalizada la orden';
  } else if (tipoCarga === 'Al día siguiente') {
    data.txtCargaCompleto = 'al día calendario siguiente de haber finalizado la orden';
  } else {
    data.txtCargaCompleto = `el ${fechaCargaFmt || "[FECHA PENDIENTE]"}`;
  }

  const tipoVigenciaCreditos = vars['¿Cómo se define la vigencia de los Créditos?'];
  if (tipoVigenciaCreditos === 'Por días calendario (Duración)') {
    const diasNum = vars['Cantidad de Días de Vigencia'] || '30';
    data.txtVigenciaCreditos = `tendrán una vigencia de ${numeroALetras(Number(diasNum))} (${diasNum}) días calendario contados a partir de su carga`;
  } else {
    let iniRed = formatDateInSpanish(parseFormDate(vars['Fecha Inicio de Redención'])) || "el momento en que sean cargados";
    let finRed = formatDateInSpanish(parseFormDate(vars['Fecha Fin de Redención'])) || "[FECHA FIN PENDIENTE]";
    data.txtVigenciaCreditos = `podrán ser utilizados entre ${iniRed} y el ${finRed}`;
  }

  // Texto segmento
  if (data.segmento === 'Pro y Pro Black') data.textoSegmento = 'Pueden participar los Usuarios/Consumidores que tengan activa la suscripción RappiPro y/o RappiPro Black.';
  else if (data.segmento === 'Nuevos Usuarios') data.textoSegmento = 'Campaña válida únicamente para Nuevos Usuarios/Consumidores.';
  else data.textoSegmento = 'Pueden participar todos los Usuarios/Consumidores de la Plataforma Rappi que sean mayores de edad.';

  // Texto método pago
  if (data.metodosPago === 'Todos excepto Efectivo') data.textoMetodoPago = 'Campaña válida para órdenes pagadas con todos los medios de pago habilitados en la Plataforma Rappi, excepto efectivo. No se obtendrá el Beneficio respecto de órdenes pagadas en efectivo o parcial/totalmente con Créditos.';
  else if (data.metodosPago && data.metodosPago.includes('Todos')) data.textoMetodoPago = 'Campaña válida para órdenes pagadas con todos los medios de pago habilitados en la Plataforma Rappi.';
  else data.textoMetodoPago = `Campaña válida únicamente para órdenes pagadas con ${(data.metodosPago || '').replace('Únicamente ', '')}.`;

  // Texto redención
  const suffix = " únicamente dentro del Territorio.";
  if (redencionSeleccionada.includes('Restaurantes')) data.textoLugarRedencion = `Se aclara que los Créditos otorgados podrán ser redimidos en cualquier tienda de la sección "Restaurantes"${suffix}`;
  else if (redencionSeleccionada.includes('Plataforma')) data.textoLugarRedencion = `Se aclara que los Créditos otorgados podrán ser redimidos en cualquier sección de la Plataforma Rappi (excepto Cajero ATM y RappiFavor)${suffix}`;
  else data.textoLugarRedencion = `Se aclara que los Créditos otorgados podrán ser redimidos únicamente en ${data.txtRefTienda} donde se originó el Beneficio${suffix}`;

  data.nombreCampana = vars['Nombre de Campaña (Opcional)'] || `CASHBACK ${pctNum}% ${data.tiendaArchivo || data.tiendaBase.toUpperCase()}`; 
  data.docName = `T&C - ${data.nombreCampana}`;
  return data;
}

function procesarConcurso(vars) {
  const data = processCommonVariables(vars);
  data.dinamica = 'TOP_SPENDER';
  data.razonSocialOrganizador = vars['Razón Social Organizador'] || data.tiendaBase;
  data.telefonoContacto = vars['Teléfono Contacto Organizador'] || 'Soporte In-App Rappi';
  data.emailContacto = vars['Email Contacto Organizador'] || 'servicioalcliente@rappi.com';
  data.numeroGanadores = Number(vars['Número de Ganadores']) || 1;
  data.numeroGanadoresLetras = numeroALetras(data.numeroGanadores);
  data.txtGanadoresPlural = data.numeroGanadores === 1 ? 'ganador' : 'ganadores';
  
  // Variables obligatorias que faltaban en el map
  data.criterioGanador = (vars['Criterio del Ganador'] && vars['Criterio del Ganador'].includes('$')) ? 'mayor valor acumulado (dinero) en compras' : 'mayor cantidad de órdenes finalizadas';
  data.verticales = vars['Verticales/Secciones participantes'] || "Restaurantes";
  data.productosParticipantes = vars['Productos Participantes'] || `todos los productos de la marca ${data.tiendaBase}`;
  data.fechaAnuncio = formatDateInSpanish(parseFormDate(vars['Fecha de Anuncio de Ganadores'])) || "[FECHA PENDIENTE]";
  data.minimoCompraTexto = vars['Mínimo de Compra para participar'] ? `$${vars['Mínimo de Compra para participar']} M/Cte` : 'No aplica un valor mínimo.';

  data.isPremioCreditos = vars['Tipo de Premio'] === 'credits';
  if (data.isPremioCreditos) {
    data.montoCreditosPremioFmt = Number(vars['Monto Créditos Premio'] || 50000).toLocaleString('es-CO');
    data.montoCreditosPremioLetras = numeroALetras(Number(vars['Monto Créditos Premio'] || 50000));
    data.diasCargaPremio = vars['Días para Carga Premio'] || 5;
    data.diasCargaPremioLetras = numeroALetras(Number(data.diasCargaPremio));
    data.vigenciaCreditosPremio = vars['Vigencia Créditos Premio'] || 30;
    data.vigenciaCreditosPremioLetras = numeroALetras(Number(data.vigenciaCreditosPremio));
    data.textoLugarRedencionPremio = `únicamente en la Tienda Participante "${data.tiendaBase}"`;
    data.listaPremios = `${data.montoCreditosPremioLetras} (${data.montoCreditosPremioFmt}) Créditos de la App por cada ganador.`;
    data.responsableEntrega = "Rappi";
  } else {
    data.listaPremios = cleanTechNames(vars['Descripción Premio Físico'] || "Premio sorpresa.");
    data.responsableEntrega = vars['Quién Entrega Premio'] === 'rappi' ? "Rappi" : "el Organizador";
  }

  data.textoSegmento = 'Pueden participar todos los Usuarios/Consumidores de la Plataforma Rappi que sean mayores de edad.';
  data.textoMetodoPago = 'Actividad válida para órdenes pagadas con todos los medios de pago habilitados en la Plataforma Rappi, excepto efectivo. No sumarán al acumulado las órdenes pagadas en efectivo o parcial/totalmente con Créditos.';

  data.nombreCampana = vars['Nombre de Campaña (Opcional)'] || `CONCURSO ${data.tiendaBase.toUpperCase()}`; 
  data.docName = `T&C - ${data.nombreCampana}`;
  return data;
}

function _buildPlaceholderMap(data) {
  const map = {
    '{{NOMBRE_CAMPANA_UPPER}}': (data.nombreCampana || '').toUpperCase(),
    '{{NOMBRE_CAMPANA}}': data.nombreCampana || '',
    '{{TEXTO_TERRITORIO}}': data.textoTerritorio || '',
    '{{FECHA_INICIO}}': data.startFmtDate || '',
    '{{FECHA_FIN}}': data.endFmtDate || '',
    '{{HORA_INICIO}}': data.startFmtTime || '',
    '{{HORA_FIN}}': data.endFmtTime || '',
    '{{TIENDA_BASE}}': data.tiendaBase || '',
    '{{TEXTO_SEGMENTO}}': data.textoSegmento || '',
    '{{TEXTO_METODO_PAGO}}': data.textoMetodoPago || '',
    '{{REF_TIENDA}}': data.txtRefTienda || 'la Tienda Participante',
    '{{TIENDA_DISPLAY}}': data.tiendaDisplay || `"${data.tiendaBase}"`,
    '{{DEFINICION_TIENDA}}': data.txtDefinicionTienda || ' (en adelante la "Tienda Participante") ',
    '{{TITULO_TIENDA}}': data.txtTituloTienda || 'IV. Tienda Participante: ',
    '{{DECLARACION_TIENDA}}': data.txtDeclaracionTienda || 'es la Tienda Participante',
    '{{CONDICIONES_ESPECIALES}}': data.condicionesEspeciales || ''
  };
  
  if (data.dinamica === 'CASHBACK') {
    map['{{TEXTO_PORCENTAJE}}'] = data.textoPorcentaje || '';
    map['{{TOPE_LETRAS}}'] = data.topeLetras || '';
    map['{{TOPE_NUM}}'] = data.topeNumFmt || '';
    map['{{PRESUPUESTO_NUM}}'] = data.presupuestoNumFmt || '';
    map['{{UMBRAL_LETRAS}}'] = data.umbralCompraLetras || '';
    map['{{UMBRAL_NUM}}'] = data.umbralCompraFmt || '';
    map['{{TEXTO_CARGA}}'] = data.txtCargaCompleto || '';
    map['{{TEXTO_VIGENCIA_CREDITOS}}'] = data.txtVigenciaCreditos || '';
    map['{{TEXTO_LUGAR_REDENCION}}'] = data.textoLugarRedencion || '';
    map['{{LIMITE_ORDENES}}'] = data.limiteOrdenes || '1';
    map['{{TEXTO_ORDENES}}'] = data.txtOrdenesGramatica || 'orden';
  } else {
    map['{{ORGANIZADOR}}'] = data.razonSocialOrganizador || '';
    map['{{TELEFONO_CONTACTO}}'] = data.telefonoContacto || '';
    map['{{EMAIL_CONTACTO}}'] = data.emailContacto || '';
    map['{{NUM_GANADORES}}'] = String(data.numeroGanadores || 1);
    map['{{NUM_GANADORES_LETRAS}}'] = data.numeroGanadoresLetras || '';
    map['{{PLURAL_GANADORES}}'] = data.txtGanadoresPlural || 'ganador';
    map['{{CRITERIO_GANADOR}}'] = data.criterioGanador || '';
    map['{{VERTICALES}}'] = data.verticales || '';
    map['{{PRODUCTOS_PARTICIPANTES}}'] = data.productosParticipantes || '';
    map['{{LISTA_PREMIOS}}'] = data.listaPremios || '';
    map['{{RESPONSABLE_ENTREGA}}'] = data.responsableEntrega || '';
    map['{{FECHA_ANUNCIO}}'] = data.fechaAnuncio || '';
    map['{{MINIMO_COMPRA_TEXTO}}'] = data.minimoCompraTexto || '';
    
    if (data.isPremioCreditos) {
      map['{{MONTO_CREDITOS_LETRAS}}'] = data.montoCreditosPremioLetras || '';
      map['{{MONTO_CREDITOS_NUM}}'] = data.montoCreditosPremioFmt || '';
      map['{{DIAS_CARGA_LETRAS}}'] = data.diasCargaPremioLetras || '';
      map['{{DIAS_CARGA_NUM}}'] = String(data.diasCargaPremio || 5);
      map['{{VIGENCIA_CREDITOS_LETRAS}}'] = data.vigenciaCreditosPremioLetras || '';
      map['{{VIGENCIA_CREDITOS_NUM}}'] = String(data.vigenciaCreditosPremio || 30);
      map['{{LUGAR_REDENCION_PREMIO}}'] = data.textoLugarRedencionPremio || '';
    }
  }
  return map;
}

// -----------------------------------------------------------------
// BASE DE DATOS Y CORREOS
// -----------------------------------------------------------------
function saveResponseToSheet(vars, docUrl) {
  try {
    const sheet = _getOrCreateSheet('Respuestas_Audit_V2', ['Timestamp', 'Email Generador', 'Link Documento', 'Tipo Dinámica', 'País', 'Tienda', 'Campaña', 'JSON_DATA']);
    sheet.appendRow([new Date(), vars['Email Generador'] || '', docUrl, vars['Tipo de Dinámica'] || '', vars['País Seleccionado'] || 'CO', vars['Tienda Participante'] || '', vars['Nombre de Campaña (Opcional)'] || '', JSON.stringify(vars)]);
  } catch (error) {}
}

function sendEmailNotification(submitterEmail, docName, docUrl, data) {
  const htmlBody = getEmailTemplate(data, docUrl, docName);
  MailApp.sendEmail({ to: submitterEmail, subject: `T&C: ${data.nombreCampana}`, htmlBody: htmlBody, name: "Motor Legal Rappi" });
}

function getEmailTemplate(data, docUrl, docName) {
  const color = (data.dinamica === 'CASHBACK') ? '#FF441F' : '#2962FF';
  const icon = (data.dinamica === 'CASHBACK') ? '💰' : '🏆';
  const typeLabel = (data.dinamica === 'CASHBACK') ? 'Campaña de Cashback' : 'Concurso / Top Spender';
  
  let tableRows = '';
  if (data.dinamica === 'CASHBACK') {
    tableRows = `
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🏪 Tienda:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.tiendaBase}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>💸 Beneficio:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.textoPorcentaje}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🛑 Tope:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.topeNumFmt}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>💰 Presupuesto:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.presupuestoNumFmt}</td></tr>
    `;
  } else {
    tableRows = `
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>👑 Organizador:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.razonSocialOrganizador}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🎯 Criterio:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.criterioGanador}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🎁 Premio:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.listaPremios}</td></tr>
    `;
  }

  return `
    <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
      <div style="background-color: ${color}; padding: 20px; text-align: center; color: white;">
        <h1 style="margin:0; font-size: 24px;">${icon} Documento Generado</h1>
        <p style="margin:5px 0 0 0; opacity: 0.9;">Motor Legal Rappi V2.4</p>
      </div>
      <div style="padding: 30px; background-color: white;">
        <p style="color: #374151; font-size: 16px; margin-top: 0;">Hola,</p>
        <p style="color: #4b5563; line-height: 1.5;">Se han generado exitosamente los Términos y Condiciones:</p>
        <div style="background-color: #f3f4f6; border-radius: 8px; padding: 15px; margin: 20px 0;">
          <h3 style="margin-top:0; color: ${color}; border-bottom: 2px solid #e5e7eb; padding-bottom: 10px;">${typeLabel}</h3>
          <table style="width:100%; border-collapse: collapse; font-size: 14px; color: #374151;">
            ${tableRows}
          </table>
        </div>
        <div style="text-align: center; margin: 30px 0;">
          <a href="${docUrl}" style="background-color: ${color}; color: white; padding: 14px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 16px; display: inline-block;">
            Abrir Documento
          </a>
        </div>
      </div>
    </div>
  `;
}

function _getTemplateRegistry() {
  return _sheetToObjects(_getSheet(TW_CONFIG.SHEET_REGISTRY) || []);
}
// -----------------------------------------------------------------
// ENDPOINTS PARA FRONTEND DINÁMICO (GOD MODE)
// -----------------------------------------------------------------
function getCampaignTypesForUser() {
  try {
    const sheet = _getSheet(CAMPAIGN_TYPES_SHEET);
    if (!sheet) return JSON.stringify([]);
    const all = _sheetToObjects(sheet);
    const active = all.filter(t => String(t.status).toLowerCase() === 'active');
    return JSON.stringify(active.map(t => ({
      type_id:    t.type_id,
      type_name:  t.type_name,
      description: t.description || '',
      icon:       t.icon || 'fa-bolt',
      color:      t.color || '#FF4500',
      countries:  t.countries || 'ALL'
    })));
  } catch (e) {
    Logger.log('❌ getCampaignTypesForUser: ' + e.message);
    return JSON.stringify([]);
  }
}

function getFieldsForUserForm(campaignType, countryCode) {
  try {
    const sheet = _getSheet(FIELDS_SHEET_NAME);
    if (!sheet) return JSON.stringify([]);
    const all = _sheetToObjects(sheet);
    const filtered = all.filter(f => {
      const matchType = (f.campaign_type === 'ALL' || f.campaign_type === campaignType);
      const matchCountry = (f.country_code === 'ALL' || f.country_code === countryCode);
      // V3.2: No enviar campos base (canonical) al frontend — ya se preguntan en el formulario estático
      // canonical_field_id es la señal principal; section='0' es redundancia de seguridad
      const canonicalVal = (f.canonical_field_id || '').toString().trim();
      // V3.3: Excluir campos base (section 0), legales auto-resueltos (section L), y canónicos
      var isNotBaseField = (String(f.section) !== '0') && (String(f.section) !== 'L') && (canonicalVal === '');
      return matchType && matchCountry && isNotBaseField;
    });
    // Ordenar por section y luego por order
    filtered.sort((a, b) => {
      const secDiff = Number(a.section || 99) - Number(b.section || 99);
      if (secDiff !== 0) return secDiff;
      return Number(a.order || 99) - Number(b.order || 99);
    });
    return JSON.stringify(filtered.map(f => ({
      field_id:       f.field_id,
      label:          f.label_es || f.field_id,
      field_type:     f.field_type || 'text',
      icon:           f.icon || '',
      required:       String(f.required).toUpperCase() === 'TRUE',
      section:        String(f.section || '3'),
      options:        f.options ? String(f.options).split('|').map(o => o.trim()) : [],
      default_value:  f.default_value || '',
      tooltip:        f.tooltip || '',
      depends_on:     f.depends_on || '',
      order:          Number(f.order || 0),
      group:          f.group || '',
      validation_rule: f.validation_rule || '',
      placeholder:    f.placeholder || ''
    })));
  } catch (e) {
    Logger.log('❌ getFieldsForUserForm: ' + e.message);
    return JSON.stringify([]);
  }
}
function _getCampaignTypeConfig(campaignType) {
  const sheet = _getSheet('Campaign_Types');
  if (!sheet) return null;
  return _sheetToObjects(sheet).find(t => t.type_name === campaignType && t.status === 'active') || null;
}
