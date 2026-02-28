// =================================================================
// MOTOR LEGAL RAPPI V2.1 - GLOBAL T&C GENERATOR
// Actualizado: Febrero 2026
// =================================================================

// -----------------------------------------------------------------
// CONSTANTES GLOBALES
// -----------------------------------------------------------------
const LEGAL_AUDIT_EMAIL = 'juan.gallego@rappi.com,david.gaviria@rappi.com'; 
const DRIVE_FOLDER_ID = ''; 
const MESES_ES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const LISTA_MUNICIPIOS = ['Soacha', 'Chía', 'Cajicá', 'Palmira', 'Bello', 'Buga', 'Envigado', 'Itagüí', 'Sabaneta', 'Jamundí', 'Yumbo', 'Floridablanca', 'Girón', 'Piedecuesta', 'Rionegro', 'Dosquebradas'];

// -----------------------------------------------------------------
// FUNCIÓN DOGET - PARA SERVIR LA WEB APP
// -----------------------------------------------------------------
function doGet(e) {
  return HtmlService.createTemplateFromFile('WebApp')
      .evaluate()
      .setTitle('Motor Legal Rappi - Web V2.1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// -----------------------------------------------------------------
// FUNCIÓN PARA RECIBIR DATOS DE LA WEB APP
// -----------------------------------------------------------------
function processWebPayload(payloadString) {
  try {
    const payload = JSON.parse(payloadString);
    
    Logger.log('========== MOTOR LEGAL V2.1 ==========');
    Logger.log('📥 Tipo de campaña: ' + payload.dynamicType);
    
    const result = coreEngineV2(payload, payload.userEmail);
    return JSON.stringify({
        status: 'success',
        docUrl: result.docUrl,
        docName: result.docName
    });
  } catch (e) {
    Logger.log('❌ ERROR: ' + e.message);
    return JSON.stringify({ 
      status: 'error', 
      message: e.message, 
      stack: e.stack 
    });
  }
}

// -----------------------------------------------------------------
// DICCIONARIO DE TRADUCCIÓN (WEB → MOTOR)
// -----------------------------------------------------------------
function mapWebToEngine(webData) {
  return {
    // Campos comunes
    'Tipo de Dinámica': webData.dynamicType,
    'Tienda Participante': webData.shopName,
    'Territorio': webData.territory,
    'Fecha de INICIO de Campaña': webData.startDate,
    'Hora de INICIO de Campaña': webData.startTime,
    'Fecha de FIN de Campaña': webData.endDate,
    'Hora de FIN de Campaña': webData.endTime,
    'Nombre de Campaña (Opcional)': webData.campaignName,
    'Email(s) adicionales': webData.extraEmails,
    'Condiciones Especiales (Adicionales)': webData.specialConditions,
    'Tipo de Usuarios Participantes': webData.userSegment,
    'Métodos de Pago Válidos': webData.paymentMethods,
    'Máximo de Órdenes por Usuario': webData.maxOrders,
    'Valor Mínimo de Compra (Opcional)': webData.minPurchase,
    
    // Campos de Cashback
    'Lugar de Redención de Créditos': webData.redemptionPlace,
    'Porcentaje del Cashback': webData.cashbackPct,
    'Presupuesto (Existencias)': webData.budget,
    'Tope Máximo de Cashback': webData.cap,
    'Tipo de Carga de Créditos': webData.loadType,
    'Fecha de Carga de Créditos': webData.loadDate,
    '¿Cómo se define la vigencia de los Créditos?': webData.validityType,
    'Cantidad de Días de Vigencia': webData.validityDays,
    'Fecha Inicio de Redención': webData.redemptionStart,
    'Fecha Fin de Redención': webData.redemptionEnd,
    
    // Campos de Concurso (NUEVOS V2.1)
    'Razón Social Organizador': webData.organizerLegalName,
    'Teléfono Contacto Organizador': webData.organizerPhone,
    'Email Contacto Organizador': webData.organizerEmail,
    'Número de Ganadores': webData.numberOfWinners,
    'Tipo de Premio': webData.prizeType,
    'Quién Entrega Premio': webData.prizeDeliveryBy,
    'Monto Créditos Premio': webData.creditsAmount,
    'Días para Carga Premio': webData.creditLoadDays,
    'Vigencia Créditos Premio': webData.creditsValidityDays,
    'Lugar Redención Premio': webData.creditsRedemptionPlace,
    'Descripción Premio Físico': webData.physicalPrizeDescription,
    'Verticales/Secciones participantes': webData.verticals,
    'Productos Participantes': webData.participatingProducts,
    'Criterio del Ganador': webData.winnerCriteria,
    'Mínimo de Compra para participar': webData.minParticipation,
    'Lista de Premios': webData.prizes,
    'Fecha de Anuncio de Ganadores': webData.announcementDate,
    
    // Multi-país
    'País Seleccionado': webData.countryCode,
    'Código de Moneda': webData.currencyCode,
    'Símbolo de Moneda': webData.currencySymbol
  };
}

// -----------------------------------------------------------------
// TRIGGER PARA GOOGLE FORMS (Legacy)
// -----------------------------------------------------------------
function onFormSubmit(e) {
  try {
    const responses = e.response.getItemResponses();
    const submitterEmail = e.response.getRespondentEmail();
    const vars = getResponsesAsMap(responses);
    const result = coreEngine(vars, submitterEmail);
    saveLinkToSheet(result.docUrl);
  } catch (error) {
    Logger.log(error);
    const submitterEmail = e ? e.response.getRespondentEmail() : 'unknown';
    if (submitterEmail !== 'unknown') {
      MailApp.sendEmail(submitterEmail, '⚠️ ERROR T&C', `Hubo un error:\n\n❌ ${error.message}`);
    }
  }
}

// -----------------------------------------------------------------
// NÚCLEO CENTRAL (CORE ENGINE)
// -----------------------------------------------------------------
function coreEngine(vars, submitterEmail) {
    // --- NUEVA LÍNEA PARA AUDITORÍA ---
  vars['Email Generador'] = submitterEmail;
  // ----------------------------------
    validateDates(vars);

    const tipoDinamica = vars['Tipo de Dinámica'] || 'Cashback'; 
    Logger.log('🎯 Tipo de Dinámica: ' + tipoDinamica);
    
    let data = {};

    if (tipoDinamica === 'Concurso Mayor Comprador') {
        Logger.log('   → Procesando CONCURSO');
        data = procesarConcurso(vars); 
    } else {
        Logger.log('   → Procesando CASHBACK');
        data = procesarCashback(vars); 
    }

    auditData(data);
    
    const doc = (data.dinamica === 'TOP_SPENDER') ? createConcursoDoc(data) : createCashbackDoc(data);
    const publicUrl = setPublicViewPermissions(doc); 

    // 1. Guardar auditoría en Google Sheets
    saveResponseToSheet(vars, publicUrl);

    // 2. Enviar correo de notificación
    sendEmailNotification(submitterEmail, data.docName, publicUrl, data);

    // 3. RETORNAR EL RESULTADO (Esto soluciona el error undefined)
    return {
        docUrl: publicUrl,
        docName: data.docName
    };
} 
    // =================================================================
// GUARDAR EN GOOGLE SHEETS (AUDITORÍA COMPLETA V2.1)
// =================================================================
function saveResponseToSheet(vars, docUrl) {
  try {
    // Tu ID de hoja original
    const SHEET_ID = '1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI';
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    // Nombre de la nueva pestaña para no dañar la anterior
    const sheetName = 'Respuestas_Audit_V2'; 
    let sheet = ss.getSheetByName(sheetName);
    
    // Si la hoja no existe, la crea con los encabezados
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = [
        'Timestamp', 
        'Email Generador', 
        'Link Documento',
        'Tipo Dinámica', 
        'País', 
        'Tienda/Organizador', 
        'Nombre Campaña', 
        'Territorio', 
        'Fecha Inicio', 
        'Fecha Fin',
        // Datos Financieros
        '% Cashback', 'Tope', 'Presupuesto',
        // Datos Concurso
        '# Ganadores', 'Premio', 'Criterio', 
        // Restricciones
        'Segmento Usuarios', 'Métodos Pago', 'Min Compra', 'Condiciones Extra',
        // --- CAJA NEGRA (AUDITORÍA TOTAL) ---
        'JSON_DATA_FULL' 
      ];
      
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      headerRange.setBackground('#FF441F'); // Naranja Rappi
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    // Preparamos los datos (usamos || para evitar espacios vacíos)
    const row = [
      new Date(),
      vars['Email Generador'] || 'No registrado',
      docUrl,
      vars['Tipo de Dinámica'] || 'Cashback',
      vars['País Seleccionado'] || 'CO',
      vars['Tienda Participante'] || vars['Razón Social Organizador'] || 'N/A',
      vars['Nombre de Campaña (Opcional)'] || '',
      vars['Territorio'] || '',
      vars['Fecha de INICIO de Campaña'] || '',
      vars['Fecha de FIN de Campaña'] || '',
      
      // Financieros
      vars['Porcentaje del Cashback'] || '-',
      vars['Tope Máximo de Cashback'] || '-',
      vars['Presupuesto (Existencias)'] || '-',
      
      // Concurso
      vars['Número de Ganadores'] || '-',
      vars['Tipo de Premio'] === 'credits' ? (vars['Monto Créditos Premio'] + ' créditos') : (vars['Descripción Premio Físico'] || '-'),
      vars['Criterio del Ganador'] || '-',
      
      // Restricciones
      vars['Tipo de Usuarios Participantes'] || '',
      vars['Métodos de Pago Válidos'] || '',
      vars['Valor Mínimo de Compra (Opcional)'] || vars['Mínimo de Compra para participar'] || '0',
      vars['Condiciones Especiales (Adicionales)'] || '',
      
      // --- AQUÍ GUARDAMOS TODO LO DEMÁS ---
      JSON.stringify(vars)
    ];
    
    sheet.appendRow(row);
    Logger.log('✅ Respuesta guardada en pestaña de auditoría');
    return true;
    
  } catch (error) {
    Logger.log('❌ Error guardando en Sheet: ' + error.message);
    return false;
  }
}
// =================================================================
// PROCESAMIENTO - CASHBACK
// =================================================================
function procesarCashback(vars) {
  const data = processCommonVariables(vars);
  data.dinamica = 'CASHBACK';

  const redencionSeleccionada = vars['Lugar de Redención de Créditos'] || 'Únicamente en la Tienda Participante (Brand Credits)';
  data.isTurbo = (redencionSeleccionada.includes('Turbo')) || (data.tiendaBase.toLowerCase().includes('turbo'));
  data.isTurboSubsection = (redencionSeleccionada === 'Únicamente en una Subsección de Turbo (Ej: Turbo Nutresa)');
  data.isTurboMarket = redencionSeleccionada.includes('Turbo Market') || redencionSeleccionada.includes('Híbrido: Tienda Marketplace + Turbo Market') || data.isTurboSubsection;
  const isHybrid = redencionSeleccionada.includes('Híbrido');
  const isGeneric = (vars['Tienda Participante'].toUpperCase().startsWith('TODAS') || vars['Tienda Participante'].toUpperCase().startsWith('TODOS'));

  if (data.isTurboSubsection) {
    data.nombreSubseccion = `"${data.tiendaBase}"`;
    data.tiendaArchivo = data.tiendaBase;
    data.txtTituloTienda = 'IV. Tienda y Sección Participante: ';
    data.useSubsectionWording = true;
    data.txtRefTienda = 'la Sección Participante';
    data.txtDeclaracionTienda = 'es la Tienda Participante';
  } else if (isGeneric) {
    data.isGenericCampaign = true;
    data.tiendaDisplay = "Aliados Comerciales"; 
    if (redencionSeleccionada.includes('Turbo Restaurantes')) {
      data.nombreVertical = 'la sección "Turbo Restaurantes"';
      data.tiendaArchivo = "TODOS TURBO RESTAURANTES";
      data.prefixTiendas = 'Participarán todas las tiendas virtuales que se encuentren activas en ';
    } else if (redencionSeleccionada.includes('Turbo Market')) {
      data.nombreVertical = 'la sección "Turbo"';
      data.tiendaArchivo = "TODOS TURBO MARKET";
      data.prefixTiendas = 'Participarán todas las tiendas virtuales (Aliados Comerciales) que se encuentren activas en ';
    } else {
      data.nombreVertical = 'la sección "Restaurantes"';
      data.tiendaArchivo = "TODOS RESTAURANTES";
      data.prefixTiendas = 'Participarán todas las tiendas virtuales (Aliados Comerciales) que se encuentren activas en ';
    }
    data.txtTituloTienda = 'IV. Tiendas Participantes: ';
    data.txtDefinicionTienda = ' (en adelante las "Tiendas Participantes") ';
    data.txtRefTienda = 'las Tiendas Participantes';
    data.txtDeclaracionTienda = 'son las Tiendas Participantes';
  } else if (isHybrid) {
    data.tiendaDisplay = `"${data.tiendaBase}" y "${data.tiendaBase} Turbo"`;
    data.tiendaArchivo = data.tiendaBase;
    data.txtTituloTienda = 'IV. Tiendas Participantes: ';
    data.txtDefinicionTienda = ' (en adelante las "Tiendas Participantes") ';
    data.txtRefTienda = 'las Tiendas Participantes';
    data.txtDeclaracionTienda = 'son las Tiendas Participantes';
    data.isGenericCampaign = false;
  } else {
    if (data.tiendaBase.includes(',') || data.tiendaBase.includes(' y ')) {
      data.tiendaDisplay = `"${data.tiendaBase}"`;
      data.tiendaArchivo = data.tiendaBase;
      data.txtTituloTienda = 'IV. Tiendas Participantes: ';
      data.txtDefinicionTienda = ' (en adelante las "Tiendas Participantes") ';
      data.txtRefTienda = 'las Tiendas Participantes';
      data.txtDeclaracionTienda = 'son las Tiendas Participantes';
    } else {
      data.tiendaDisplay = `"${data.tiendaBase}"`;
      data.tiendaArchivo = data.tiendaBase;
      data.txtTituloTienda = 'IV. Tienda Participante: ';
      data.txtDefinicionTienda = ' (en adelante la "Tienda Participante") ';
      data.txtRefTienda = 'la Tienda Participante';
      data.txtDeclaracionTienda = 'es la Tienda Participante';
    }
    data.isGenericCampaign = false;
  }

  data.segmento = vars['Tipo de Usuarios Participantes']; 
  data.metodosPago = vars['Métodos de Pago Válidos'];
  data.extraEmails = vars['Email(s) adicionales'];
  data.limiteOrdenes = vars['Máximo de Órdenes por Usuario'] || '1';
  const ordenNum = parseInt(data.limiteOrdenes);
  data.txtOrdenesGramatica = (!isNaN(ordenNum) && ordenNum === 1) ? 'orden' : 'órdenes';
  if (isNaN(ordenNum)) data.isOrdenesTexto = true;

  const minCompraInput = vars['Valor Mínimo de Compra (Opcional)'];
  if (minCompraInput && Number(minCompraInput) > 0) {
    data.minCompra = Number(minCompraInput).toLocaleString('es-CO');
    data.hasMinCompra = true;
  }
  
  const condicionesExtra = vars['Condiciones Especiales (Adicionales)'];
  if (condicionesExtra) {
    const cleanCond = condicionesExtra.trim();
    const lowerCond = cleanCond.toLowerCase().replace(/\.$/, ""); 
    const invalidos = ['n/a', 'na', 'no', 'ninguna', 'ninguno', 'no aplica', 'sin condiciones', '0', '-'];
    if (cleanCond !== '' && !invalidos.includes(lowerCond)) {
      data.condicionesEspeciales = cleanCond;
      data.hasCondiciones = true;
    }
  }

  const pctInput = vars['Porcentaje del Cashback'] || '99'; 
  const pctNum = Number(pctInput); 
  const pctLetras = numeroALetras(Math.floor(pctNum));
  data.textoPorcentaje = `${pctLetras} por ciento (${pctNum}%)`;

  const keyPresupuesto = vars['Presupuesto (Existencias)'] ? 'Presupuesto (Existencias)' : 'Presupuesto (Existencias) - EN NÚMEROS';
  const keyTope = vars['Tope Máximo de Cashback'] ? 'Tope Máximo de Cashback' : 'Tope Máximo de Cashback - EN NÚMEROS';
  const presupuestoNum = Number(vars[keyPresupuesto]);
  const topeNum = Number(vars[keyTope]);
  
  data.presupuestoLetras = numeroALetras(presupuestoNum);
  data.topeLetras = numeroALetras(topeNum);
  data.presupuestoNumFmt = isNaN(presupuestoNum) ? "0" : presupuestoNum.toLocaleString('es-CO');
  data.topeNumFmt = isNaN(topeNum) ? "0" : topeNum.toLocaleString('es-CO');

  let umbralCompraNum = (pctNum > 0) ? Math.ceil(topeNum / (pctNum / 100)) : topeNum;
  data.umbralCompraFmt = umbralCompraNum.toLocaleString('es-CO');
  data.umbralCompraLetras = numeroALetras(umbralCompraNum); 

  const tipoCarga = vars['Tipo de Carga de Créditos'];
  const objFechaCarga = parseFormDate(vars['Fecha de Carga de Créditos']);
  let fechaCargaFmt = formatDateInSpanish(objFechaCarga); 
  if (tipoCarga === 'Inmediatamente (Al finalizar la orden)') {
    data.txtCargaCompleto = 'de manera inmediata una vez finalizada la orden';
    data.txtInicioUso_Fallback = 'el momento en que sean cargados';
  } else if (tipoCarga === 'Al día siguiente') {
    data.txtCargaCompleto = 'al día calendario siguiente de haber finalizado la orden';
    data.txtInicioUso_Fallback = 'el momento en que sean cargados';
  } else {
    data.txtCargaCompleto = `el ${fechaCargaFmt || "[FECHA PENDIENTE]"}`;
    data.txtInicioUso_Fallback = `el ${fechaCargaFmt || "[FECHA PENDIENTE]"}`;
  }
  data.fechaCarga = fechaCargaFmt;

  const tipoVigenciaCreditos = vars['¿Cómo se define la vigencia de los Créditos?'];
  if (tipoVigenciaCreditos === 'Por días calendario (Duración)') {
    const diasNum = vars['Cantidad de Días de Vigencia'] || '30';
    const diasLetras = numeroALetras(Number(diasNum));
    data.txtVigenciaCreditos = `tendrán una vigencia de ${diasLetras} (${diasNum}) días calendario contados a partir de su carga`;
    data.usarTextoVigenciaEnEjemplo = true;
  } else {
    let iniRed = formatDateInSpanish(parseFormDate(vars['Fecha Inicio de Redención']));
    if (!iniRed) {
      iniRed = fechaCargaFmt ? `el ${fechaCargaFmt}` : "el momento en que sean cargados";
    }
    if (iniRed.includes('undefined') || iniRed === "null") {
      iniRed = "el momento en que sean cargados";
    }
    let finRed = formatDateInSpanish(parseFormDate(vars['Fecha Fin de Redención'])) || "[FECHA FIN PENDIENTE]";
    data.txtVigenciaCreditos = `podrán ser utilizados entre ${iniRed} y el ${finRed}`;
    data.fechaFinRedencion = finRed; 
  }
  if (!data.fechaFinRedencion || data.fechaFinRedencion === "PENDIENTE") {
    data.usarTextoVigenciaEnEjemplo = true;
  }

  // Texto de segmento
  if (data.segmento === 'Pro y Pro Black') {
    data.textoSegmento = 'Pueden participar los Usuarios/Consumidores que durante la Vigencia de la presente Campaña tengan activa la suscripción RappiPro y/o RappiPro Black.';
  } else if (data.segmento === 'Nuevos Usuarios') {
    data.textoSegmento = 'Campaña válida únicamente para Nuevos Usuarios/Consumidores. Se entiende por estos aquellos que se registren por primera vez en la Plataforma Rappi durante la Vigencia de la Campaña y/o que nunca hayan realizado una orden al interior de la Plataforma Rappi.';
  } else if (data.segmento === 'Usuarios Reactivos (Inactivos 28 días)') {
    data.textoSegmento = 'Campaña válida únicamente para Usuarios/Consumidores Reactivos (aquellos que no hayan realizado ninguna orden al interior de la Plataforma Rappi durante los 28 días calendario anteriores).';
  } else if (data.segmento === 'Nuevos Usuarios en la Vertical') {
    data.textoSegmento = 'Campaña válida únicamente para Nuevos Usuarios/Consumidores en la Vertical (aquellos que nunca hayan realizado una orden en la sección de la Plataforma Rappi a la cual pertenece la Tienda Participante).';
  } else {
    data.textoSegmento = 'Pueden participar todos los Usuarios/Consumidores de la Plataforma Rappi que sean mayores de edad y cumplan con los requisitos descritos.';
  }

  // Texto de método de pago
  if (data.metodosPago === 'Todos excepto Efectivo') {
    data.textoMetodoPago = 'Campaña válida para órdenes pagadas con todos los medios de pago habilitados en la Plataforma Rappi, excepto efectivo. No se obtendrá el Beneficio respecto de órdenes pagadas en efectivo o parcial/totalmente con Créditos.';
  } else if (data.metodosPago && data.metodosPago.includes('Todos') && !data.metodosPago.includes('excepto')) {
    data.textoMetodoPago = 'Campaña válida para órdenes pagadas con todos los medios de pago habilitados en la Plataforma Rappi. No se obtendrá el Beneficio respecto de órdenes pagadas parcial/totalmente con Créditos.';
  } else {
    const metodoLimpio = (data.metodosPago || '').replace('Únicamente ', '');
    data.textoMetodoPago = `Campaña válida únicamente para órdenes pagadas con ${metodoLimpio}. No se obtendrá el Beneficio respecto de órdenes pagadas con otros medios de pago o parcial/totalmente con Créditos.`;
  }

  // Texto de lugar de redención
  const suffixTerritorio = " únicamente dentro del Territorio.";
  switch (redencionSeleccionada) {
    case 'Únicamente en la Tienda Participante (Brand Credits)':
      data.textoLugarRedencion = `Se aclara que los Créditos otorgados en virtud de la presente Campaña podrán ser utilizados/redimidos únicamente en ${data.txtRefTienda} donde se originó el Beneficio${suffixTerritorio}`;
      break;
    case 'En cualquier tienda de la sección "Restaurantes"':
      data.textoLugarRedencion = `Se aclara que los Créditos otorgados en virtud de la presente Campaña podrán ser utilizados/redimidos en cualquier tienda que se encuentre en la sección "Restaurantes" al interior de la Plataforma Rappi${suffixTerritorio}`;
      break;
    case 'En cualquier sección de la Plataforma (Créditos Generales)':
      data.textoLugarRedencion = `Se aclara que los Créditos otorgados en virtud de la presente Campaña podrán ser utilizados/redimidos en cualquier sección de la Plataforma Rappi (excepto Cajero ATM, RappiFavor, RappiTravel y RappiApuestas)${suffixTerritorio}`;
      break;
    default:
      data.textoLugarRedencion = `Se aclara que los Créditos otorgados en virtud de la presente Campaña podrán ser utilizados/redimidos únicamente en ${data.txtRefTienda} donde se originó el Beneficio${suffixTerritorio}`;
  }

  data.nombreCampana = vars['Nombre de Campaña (Opcional)'] || `CASHBACK ${pctNum}% ${data.tiendaArchivo.toUpperCase()}`; 
  data.nombreCampanaLower = vars['Nombre de Campaña (Opcional)'] ? vars['Nombre de Campaña (Opcional)'] : `cashback ${pctNum}% ${data.tiendaArchivo}`; 
  data.docName = `T&C - ${data.nombreCampana}`;

  return data;
}

// =================================================================
// PROCESAMIENTO - CONCURSO (MEJORADO V2.1)
// =================================================================
function procesarConcurso(vars) {
  const data = processCommonVariables(vars);
  data.dinamica = 'TOP_SPENDER';
  data.tiendaArchivo = data.tiendaBase;
  
  // Datos del Organizador (NUEVO V2.1)
  data.razonSocialOrganizador = vars['Razón Social Organizador'] || data.tiendaBase;
  data.telefonoContacto = vars['Teléfono Contacto Organizador'] || 'Soporte In-App Rappi';
  data.emailContacto = vars['Email Contacto Organizador'] || 'servicioalcliente@rappi.com';
  
  // Verticales y productos
  data.verticales = vars['Verticales/Secciones participantes'] || "Restaurantes";
  data.productosParticipantes = vars['Productos Participantes'] || `todos los productos de la marca ${data.tiendaBase}`;
  
  // Número de ganadores
  const numGanadores = Number(vars['Número de Ganadores']) || 1;
  data.numeroGanadores = numGanadores;
  data.numeroGanadoresLetras = numeroALetras(numGanadores);
  data.txtGanadoresPlural = numGanadores === 1 ? 'ganador' : 'ganadores';
  
  // Criterio de selección
  const criterioRaw = vars['Criterio del Ganador'];
  if (criterioRaw && (criterioRaw.includes('Valor') || criterioRaw.includes('$') || criterioRaw.toLowerCase().includes('dinero') || criterioRaw.toLowerCase().includes('mayor venta'))) {
    data.criterioGanador = 'mayor valor acumulado (dinero) en compras';
    data.txtDesempate1 = 'Quien haya realizado la mayor cantidad de órdenes finalizadas de Productos Participantes durante la Vigencia.';
  } else {
    data.criterioGanador = 'mayor cantidad de órdenes finalizadas';
    data.txtDesempate1 = 'Quien haya acumulado el mayor valor (dinero) en compras de Productos Participantes durante la Vigencia.';
  }

  // Mínimo de compra para participar
  const minCompraRaw = vars['Mínimo de Compra para participar'];
  if (!minCompraRaw || minCompraRaw.toLowerCase().includes('no') || minCompraRaw === '0' || minCompraRaw.toLowerCase() === 'n/a' || minCompraRaw.trim() === '') {
    data.minimoCompraTexto = 'No aplica un valor mínimo de compra por orden, sin embargo, cada orden suma al acumulado.';
    data.hasMinCompra = false;
  } else {
    const minVal = Number(minCompraRaw.replace(/\D/g,'')); 
    if (minVal > 0) {
      data.minimoCompraTexto = `$${minVal.toLocaleString('es-CO')} pesos M/Cte.`;
      data.minCompraNum = minVal;
      data.hasMinCompra = true;
    } else {
      data.minimoCompraTexto = 'No aplica un valor mínimo de compra por orden, sin embargo, cada orden suma al acumulado.';
      data.hasMinCompra = false;
    }
  }

  // Tipo de premio (NUEVO V2.1)
  data.tipoPremio = vars['Tipo de Premio'] || 'credits';
  data.isPremioCreditos = data.tipoPremio === 'credits';
  data.isPremioFisico = data.tipoPremio === 'physical';
  
  if (data.isPremioCreditos) {
    const montoCreditos = Number(vars['Monto Créditos Premio']) || 50000;
    data.montoCreditosPremio = montoCreditos;
    data.montoCreditosPremioFmt = montoCreditos.toLocaleString('es-CO');
    data.montoCreditosPremioLetras = numeroALetras(montoCreditos);
    
    const diasCarga = Number(vars['Días para Carga Premio']) || 5;
    data.diasCargaPremio = diasCarga;
    data.diasCargaPremioLetras = numeroALetras(diasCarga);
    
    const vigenciaCreditos = Number(vars['Vigencia Créditos Premio']) || 30;
    data.vigenciaCreditosPremio = vigenciaCreditos;
    data.vigenciaCreditosPremioLetras = numeroALetras(vigenciaCreditos);
    
    const lugarRedencion = vars['Lugar Redención Premio'] || 'Únicamente en la Tienda Participante';
    data.lugarRedencionPremio = lugarRedencion;
    
    if (lugarRedencion.includes('toda la Plataforma') || lugarRedencion.includes('cualquier sección')) {
      data.textoLugarRedencionPremio = 'en cualquier sección de la Plataforma Rappi (excepto Cajero ATM, RappiFavor, RappiTravel y RappiApuestas)';
    } else if (lugarRedencion.includes('Restaurante')) {
      data.textoLugarRedencionPremio = 'en cualquier tienda de la sección "Restaurantes" de la Plataforma Rappi';
    } else {
      data.textoLugarRedencionPremio = `únicamente en la Tienda Participante "${data.tiendaBase}"`;
    }
    
    data.listaPremios = `${data.montoCreditosPremioLetras} (${data.montoCreditosPremioFmt}) Créditos de la App por cada ganador.`;
    data.responsableEntrega = "Rappi";
    data.requiereAcuerdoTransferencia = false;
  } else {
    // Premio Físico
    let premioRaw = vars['Descripción Premio Físico'] || vars['Lista de Premios'] || "Premio sorpresa.";
    data.listaPremios = cleanTechNames(premioRaw);
    
    // Quién entrega el premio (NUEVO V2.1)
    const quienEntrega = vars['Quién Entrega Premio'] || 'organizer';
    if (quienEntrega === 'rappi') {
      data.responsableEntrega = "Rappi";
      data.requiereAcuerdoTransferencia = false;
    } else {
      data.responsableEntrega = "el Organizador";
      data.requiereAcuerdoTransferencia = true;
    }
  }
  
  // Fecha de anuncio
  const fechaAnuncioObj = parseFormDate(vars['Fecha de Anuncio de Ganadores']);
  data.fechaAnuncio = formatDateInSpanish(fechaAnuncioObj) || "[FECHA PENDIENTE]";

  // Restricciones comunes
  data.segmento = vars['Tipo de Usuarios Participantes'] || 'Todos los usuarios';
  data.metodosPago = vars['Métodos de Pago Válidos'] || 'Todos excepto Efectivo';
  data.limiteOrdenes = vars['Máximo de Órdenes por Usuario'] || 'Sin límite';
  data.extraEmails = vars['Email(s) adicionales'];
  
  // Texto de segmento para concurso
  if (data.segmento === 'Pro y Pro Black') { 
    data.textoSegmento = 'Pueden participar los Usuarios/Consumidores que durante la Vigencia de la presente Actividad Promocional tengan activa la suscripción RappiPro y/o RappiPro Black.'; 
  } else if (data.segmento === 'Nuevos Usuarios') { 
    data.textoSegmento = 'Actividad válida únicamente para Nuevos Usuarios/Consumidores.'; 
  } else if (data.segmento === 'Usuarios Reactivos (Inactivos 28 días)') { 
    data.textoSegmento = 'Actividad válida únicamente para Usuarios/Consumidores Reactivos (aquellos que no hayan realizado ninguna orden durante los 28 días calendario anteriores).'; 
  } else { 
    data.textoSegmento = 'Pueden participar todos los Usuarios/Consumidores de la Plataforma Rappi que sean mayores de edad y cumplan con los requisitos descritos.'; 
  }
  
  // Texto de método de pago para concurso
  if (data.metodosPago === 'Todos excepto Efectivo') { 
    data.textoMetodoPago = 'Actividad válida para órdenes pagadas con todos los medios de pago habilitados en la Plataforma Rappi, excepto efectivo. No sumarán al acumulado las órdenes pagadas en efectivo o parcial/totalmente con Créditos.'; 
  } else if (data.metodosPago && data.metodosPago.includes('Todos') && !data.metodosPago.includes('excepto')) { 
    data.textoMetodoPago = 'Actividad válida para órdenes pagadas con todos los medios de pago habilitados en la Plataforma Rappi. No sumarán al acumulado las órdenes pagadas parcial/totalmente con Créditos.'; 
  } else { 
    const metodoLimpio = (data.metodosPago || '').replace('Únicamente ', ''); 
    data.textoMetodoPago = `Actividad válida únicamente para órdenes pagadas con ${metodoLimpio}. No sumarán al acumulado las órdenes pagadas con otros medios de pago o parcial/totalmente con Créditos.`; 
  }
  
  // Condiciones especiales
  const condicionesExtra = vars['Condiciones Especiales (Adicionales)'];
  if (condicionesExtra) {
    const cleanCond = condicionesExtra.trim();
    const lowerCond = cleanCond.toLowerCase().replace(/\.$/, ""); 
    const invalidos = ['n/a', 'na', 'no', 'ninguna', 'ninguno', 'no aplica', 'sin condiciones', '0', '-'];
    if (cleanCond !== '' && !invalidos.includes(lowerCond)) { 
      data.condicionesEspeciales = cleanCond; 
      data.hasCondiciones = true; 
    }
  }

  data.nombreCampana = vars['Nombre de Campaña (Opcional)'] || `CONCURSO ${data.tiendaBase.toUpperCase()}`; 
  data.docName = `T&C - ${data.nombreCampana}`;

  return data;
}

// =================================================================
// VARIABLES COMUNES
// =================================================================
function processCommonVariables(vars) {
  const data = {};
  let tiendaInput = vars['Tienda Participante'] ? vars['Tienda Participante'].trim() : "TIENDA";
  data.tiendaBase = capitalize(tiendaInput);
  data.marcaAliada = data.tiendaBase;

  let rawTerritorio = vars['Territorio'];
  let territoriosSeleccionados = [];
  if (Array.isArray(rawTerritorio)) {
    territoriosSeleccionados = rawTerritorio;
  } else if (typeof rawTerritorio === 'string') { 
    if (rawTerritorio.includes(',')) {
      territoriosSeleccionados = rawTerritorio.split(',').map(s => s.trim());
    } else {
      territoriosSeleccionados = [rawTerritorio];
    }
  }
  
  if (territoriosSeleccionados.includes('Nacional')) {
    data.textoTerritorio = 'el territorio nacional de la República de Colombia';
    data.territorioResumen = 'Nacional';
  } else {
    const listaFinalStr = formatListToText(territoriosSeleccionados);
    data.territorioResumen = listaFinalStr;
    let hayMunicipios = false;
    let hayCiudades = false;
    territoriosSeleccionados.forEach(lugar => {
      if (LISTA_MUNICIPIOS.includes(lugar)) {
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
  
  const startDate = new Date(
    parseInt(startDateStr.split('-')[0]),
    parseInt(startDateStr.split('-')[1]) - 1,
    parseInt(startDateStr.split('-')[2]),
    parseInt(startTimeStr.split(':')[0]),
    parseInt(startTimeStr.split(':')[1])
  );
  const endDate = new Date(
    parseInt(endDateStr.split('-')[0]),
    parseInt(endDateStr.split('-')[1]) - 1,
    parseInt(endDateStr.split('-')[2]),
    parseInt(endTimeStr.split(':')[0]),
    parseInt(endTimeStr.split(':')[1])
  );
  
  data.startFmtDate = formatDateInSpanish(startDate);
  data.endFmtDate = formatDateInSpanish(endDate);
  data.startFmtTime = formatTimeInSpanish(startDate);
  data.endFmtTime = formatTimeInSpanish(endDate);
  data.isSameDay = (data.startFmtDate === data.endFmtDate);
  
  if (data.isSameDay) {
    data.textoVigenciaEmail = `El ${data.startFmtDate} (${data.startFmtTime} - ${data.endFmtTime})`;
  } else {
    data.textoVigenciaEmail = `Del ${data.startFmtDate} al ${data.endFmtDate}`;
  }

  return data;
}

// =================================================================
// GENERADOR DE DOCUMENTO - CASHBACK
// =================================================================
function createCashbackDoc(data) {
  const doc = DocumentApp.create(data.docName);
  const body = doc.getBody();
  
  if (DRIVE_FOLDER_ID) {
    try {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const file = DriveApp.getFileById(doc.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (fError) {
      Logger.log('Nota: No se movió a carpeta.');
    }
  }

  // TÍTULO
  const titulo = body.appendParagraph(`TÉRMINOS Y CONDICIONES – CAMPAÑA "${data.nombreCampana.toUpperCase()}"`);
  titulo.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER); 
  body.appendParagraph(''); 

  // INTRODUCCIÓN
  appendRichParagraph(body, ['Por medio del presente documento se dan a conocer los términos y condiciones de la campaña denominada "', {text: data.nombreCampanaLower, bold: true}, '" (en adelante la "Campaña"). La participación en la Campaña constituye la aceptación total e incondicional de los presentes Términos y Condiciones, los cuales resultan definitivos y vinculantes para los Usuarios/Consumidores participantes, que cumplan con los requisitos aquí dispuestos.']);
  
  // I. TERRITORIO
  appendRichSection(body, 'I. Territorio: ', ['La Campaña será válida únicamente para las órdenes realizadas dentro de las zonas de cobertura de ', data.txtRefTienda, ' en la Plataforma Rappi, en ', {text: data.textoTerritorio, bold: true}, '.']);

  // II. VIGENCIA
  let vigenciaParts = [];
  if (data.isSameDay) {
    vigenciaParts = ['Campaña válida el ', {text: data.startFmtDate, bold: true}, ' desde las ', {text: data.startFmtTime, bold: true}, ' hasta las ', {text: data.endFmtTime, bold: true}];
  } else {
    vigenciaParts = ['Campaña válida desde las ', {text: data.startFmtTime, bold: true}, ' del ', {text: data.startFmtDate, bold: true}, ' hasta las ', {text: data.endFmtTime, bold: true}, ' del ', {text: data.endFmtDate, bold: true}];
  }
  vigenciaParts.push(' y/o hasta agotar existencias, lo que primero ocurra. Para efectos de la Campaña, se ha establecido un valor total máximo de ');
  vigenciaParts.push({text: data.presupuestoLetras, bold: true});
  vigenciaParts.push(' (');
  vigenciaParts.push({text: data.presupuestoNumFmt, bold: true});
  vigenciaParts.push(') de Créditos a ser entregados como Cashback a los Usuarios/Consumidores que participen en la Campaña ("Existencias"). Por consiguiente, el agotamiento de las Existencias con anterioridad al final de la Vigencia indicada implicará la terminación de la Campaña.');
  appendRichSection(body, 'II. Vigencia: ', vigenciaParts);

  // III. TIPO DE USUARIOS
  appendRichSection(body, 'III. Tipo de Usuarios Participantes: ', [data.textoSegmento]);

  // IV. TIENDA PARTICIPANTE
  if (data.useSubsectionWording) {
    appendRichSection(body, data.txtTituloTienda, ['Participarán todas las tiendas virtuales denominadas "Turbo" (en adelante la "Tienda Participante") donde se encuentre la subsección ', {text: data.nombreSubseccion, bold: true}, ' (en adelante la "Sección Participante") al interior de la Plataforma Rappi, ubicadas dentro del Territorio.']);
  } else {
    let tiendasText = [];
    if (data.isGenericCampaign) {
      tiendasText = [data.prefixTiendas, data.nombreVertical, data.txtDefinicionTienda, 'al interior de la Plataforma Rappi ubicadas dentro del Territorio.'];
    } else {
      tiendasText = ['Participarán todas las tiendas virtuales de ', {text: data.tiendaDisplay, bold: true}, data.txtDefinicionTienda, 'al interior de la Plataforma Rappi ubicadas dentro del Territorio.'];
    }
    appendRichSection(body, data.txtTituloTienda, tiendasText);
  }

  // V. PRODUCTOS PARTICIPANTES
  let productosText = [];
  if (data.useSubsectionWording) {
    productosText = ['Participarán todos los productos que hacen parte del catálogo de la Sección Participante, dentro de la Tienda Participante, al interior de la Plataforma Rappi'];
  } else {
    productosText = ['Participarán todos los productos que hacen parte del catálogo de ', data.txtRefTienda, ' al interior de la Plataforma Rappi'];
  }
  if (data.isTurboMarket) {
    productosText.push(', excepto los medicamentos de venta bajo fórmula médica y los medicamentos de venta libre que, de acuerdo con la normativa vigente, no puedan ser objeto de este tipo de campañas.');
  } else {
    productosText.push('.');
  }
  appendRichSection(body, 'V. Productos Participantes: ', productosText);

  // VI. BENEFICIO
  const pBeneficio = body.appendParagraph('');
  pBeneficio.appendText('VI. Beneficio: ').setBold(true);
  let fraseVigencia = `. Los Créditos ${data.txtVigenciaCreditos}, entendiéndose que si el Usuario/Consumidor no hace uso de ellos dentro del término estipulado los perderá, sin poder hacer uso de ellos posteriormente.`;
  appendRichTextToParagraph(pBeneficio, [
    'Los Usuarios/Consumidores Participantes que durante la Vigencia de la Campaña compren cualquiera de los Productos Participantes de ', data.txtRefTienda,
    ' recibirán en Créditos el ', {text: data.textoPorcentaje, bold: true},
    ' del valor de dichos Productos Participantes (en adelante el "Cashback"). Dichos Créditos serán cargados a su cuenta al interior de la Plataforma Rappi. Se aclara que el monto máximo del Cashback que se otorgará en Créditos es de ',
    {text: data.topeLetras, bold: true}, ' (', {text: data.topeNumFmt, bold: true}, ')',
    ' Créditos. Por lo tanto, en caso de que el Usuario/Consumidor realice una compra en ', data.txtRefTienda,
    ' por un valor superior a los ', {text: data.umbralCompraLetras, bold: true}, ' pesos M/Cte ($', {text: data.umbralCompraFmt, bold: true},
    '), recibirá un monto máximo en Créditos de ', {text: data.topeLetras, bold: true}, ' (', {text: data.topeNumFmt, bold: true}, ').',
    ' Los Créditos serán cargados a la cuenta de los Usuarios/Consumidores Participantes ', {text: data.txtCargaCompleto, bold: true}, fraseVigencia
  ]);
  pBeneficio.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  
  // VII. CONDICIONES Y RESTRICCIONES
  appendRichSection(body, 'VII. Condiciones y Restricciones: ', ['Podrán participar gratuitamente todas las personas naturales que sean Usuarios/Consumidores de la Plataforma Rappi que se encuentren en el Territorio y que cumplan las siguientes condiciones:']);
  
  const restricciones = [
    'Campaña válida únicamente para órdenes realizadas a través de ' + data.txtRefTienda + ', al interior de la Plataforma Rappi. Se aclara que no serán tenidas en cuenta las órdenes que se realicen a través de cualquier otra tienda y/o sección de la Plataforma Rappi.',
    'La presente Campaña se encuentra sujeta a los horarios de operación de los puntos de venta de ' + data.txtRefTienda + ', que ofrecen y exhiben sus productos al interior de la Plataforma Rappi.',
    'Los descuentos de los productos y/o servicios objeto de la Campaña no son intercambiables ni transferibles.',
    'Campaña válida para todas las órdenes que cumplan las condiciones y restricciones establecidas en los presentes Términos y Condiciones.',
    'Campaña válida durante la Vigencia y/o hasta agotar existencias, lo que primero ocurra.',
    'El Beneficio obtenido en virtud de la presente Campaña no es acumulable con otras promociones exhibidas en la Plataforma Rappi.'
  ];
  
  if (data.isOrdenesTexto) {
    restricciones.push(data.limiteOrdenes + ' de órdenes por Usuario/Consumidor.');
  } else {
    restricciones.push(['Máximo ', {text: data.limiteOrdenes, bold: true}, ' ', data.txtOrdenesGramatica, ' por Usuario/Consumidor.']);
  }
  
  restricciones.push(['Se aclara que el monto máximo de Créditos a recibir por el Usuario/Consumidor Participante es de ', {text: data.topeLetras, bold: true}, ' (', {text: data.topeNumFmt, bold: true}, '), de acuerdo con lo indicado en la sección VI (Beneficio) de los presentes Términos y Condiciones.']);
  restricciones.push('En caso de cancelación total o parcial de la orden, el Usuario/Consumidor Participante no tendrá derecho al Beneficio.');
  restricciones.push(data.textoLugarRedencion);
  restricciones.push('Se aclara que los Créditos no tienen algún valor monetario, ni constituyen un medio de pago, instrumento crediticio o financiero y la finalidad de su redención no es recibir una cantidad de dinero en efectivo.');
  restricciones.push('Se aclara que los Créditos no pueden ser utilizados en las secciones denominadas "Cajero ATM" y "RappiFavor" de la Plataforma Rappi, ni tampoco pueden ser utilizados para pagar el valor del costo de envío, la tarifa de servicio de un pedido realizado a través de la Plataforma Rappi y/o la propina otorgada de forma voluntaria a los repartidores independientes.');

  if (data.hasMinCompra) {
    restricciones.splice(0, 0, `Campaña válida únicamente para órdenes por un valor igual o superior a ${data.minCompra} pesos M/Cte (valor no incluye costo de envío ni propina).`);
  }
  if (data.hasCondiciones) {
    restricciones.push(data.condicionesEspeciales);
  }
  if (data.isTurbo) {
    restricciones.push('Aplican términos y condiciones del Botón Turbo, disponibles en el siguiente enlace: https://legal.rappi.com.co/colombia/terminos-y-condiciones-del-boton-turbo/ El tiempo de entrega es estimado y depende de condiciones climáticas y de tráfico.');
  }
  if (data.isTurboMarket) {
    restricciones.push('La Campaña no aplica para medicamentos de venta bajo fórmula médica ni para aquellos de venta libre que no puedan ser objeto de promociones, conforme a la normativa vigente.');
  }
  
  restricciones.forEach(restriccion => { appendRichListItem(body, restriccion); });
  
  // VIII. MEDIO DE PAGO
  appendRichSection(body, 'VIII. Medio de Pago: ', [data.textoMetodoPago]);
  
  // IX. MODIFICACIONES
  appendRichSection(body, 'IX. Modificaciones e Interpretación: ', ['Rappi se reserva el derecho de cancelar órdenes si detecta un comportamiento irregular por parte del Usuario/Consumidor en la Plataforma Rappi. Rappi se reserva el derecho de rechazar y cancelar cualquier orden que, por sus características, Rappi determine que no aplica para el Beneficio de la presente Campaña, sin previo aviso al Usuario/Consumidor.']);
  
  // X. DECLARACIÓN
  appendRichSection(body, 'X. Declaración: ', ['El Usuario/Consumidor reconoce y acepta que quien exhibe, ofrece, promociona y comercializa los productos adquiridos a través de la Plataforma Rappi ', data.txtDeclaracionTienda, '. Rappi no comercializa productos puesto que es solo una plataforma tecnológica de contacto.']);
  
  appendRichParagraph(body, ['Se aclara que los presentes Términos y Condiciones se encuentran sujetos a los Términos y Condiciones de Uso de la Plataforma Rappi, los cuales se encuentran en la siguiente dirección electrónica: https://legal.rappi.com.co/colombia/terminos-y-condiciones-de-uso-de-plataforma-rappi-2/.']); 
  
  // XI. JURISDICCIÓN
  appendRichSection(body, 'XI. Jurisdicción y Solución de conflictos: ', ['Los presentes Términos y Condiciones se regirán por las leyes de Colombia. Toda controversia surgida en razón de la Campaña o de los presentes Términos y Condiciones intentará ser resuelta por arreglo directo de las partes y/o a través de la conciliación como mecanismo alternativo de solución de conflictos. Si lo anterior no fuere posible, la controversia se someterá a la justicia ordinaria.']);

  doc.saveAndClose();
  return doc;
}

// =================================================================
// GENERADOR DE DOCUMENTO - CONCURSO (V2.1 MEJORADO)
// =================================================================
function createConcursoDoc(data) {
  const doc = DocumentApp.create(data.docName);
  const body = doc.getBody();
  
  if (DRIVE_FOLDER_ID) {
    try {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      const file = DriveApp.getFileById(doc.getId());
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    } catch (fError) {
      Logger.log('Nota: No se movió a carpeta.');
    }
  }

  // TÍTULO
  const titulo = body.appendParagraph(`TÉRMINOS Y CONDICIONES – ACTIVIDAD PROMOCIONAL "${data.nombreCampana.toUpperCase()}"`);
  titulo.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER); 
  body.appendParagraph(''); 

  // INTRODUCCIÓN
  appendRichParagraph(body, [
    'Por medio del presente documento se dan a conocer los términos y condiciones de la actividad promocional denominada "',
    {text: data.nombreCampana, bold: true},
    '" (en adelante la "Actividad Promocional"). La sola participación en esta actividad implica el conocimiento y aceptación total e incondicional de estos Términos y Condiciones, no pudiendo alegar el participante el desconocimiento de estos.'
  ]);
  
  // I. ORGANIZADOR Y TERRITORIO
  appendRichSection(body, 'I. Organizador y Territorio: ', [
    'La presente Actividad Promocional es organizada y financiada exclusivamente por ',
    {text: data.razonSocialOrganizador, bold: true},
    ' (en adelante el "Organizador"), quien es el responsable de la entrega del beneficio. RAPPI S.A.S. actúa únicamente como plataforma de contacto y medio de difusión. La Actividad Promocional será válida únicamente para pedidos realizados a través de la Plataforma Rappi en las zonas de cobertura habilitadas en ',
    {text: data.textoTerritorio, bold: true},
    ' (en adelante el "Territorio").'
  ]);

  // II. VIGENCIA
  let vigenciaTxt = [];
  if (data.isSameDay) {
    vigenciaTxt = [
      'La Actividad Promocional estará vigente el ',
      {text: data.startFmtDate, bold: true},
      ' desde las ', {text: data.startFmtTime, bold: true},
      ' hasta las ', {text: data.endFmtTime, bold: true},
      ' (en adelante, la "Vigencia").'
    ];
  } else {
    vigenciaTxt = [
      'La Actividad Promocional estará vigente desde las ',
      {text: data.startFmtTime, bold: true}, ' del ', {text: data.startFmtDate, bold: true},
      ' hasta las ', {text: data.endFmtTime, bold: true}, ' del ', {text: data.endFmtDate, bold: true},
      ' (en adelante, la "Vigencia").'
    ];
  }
  vigenciaTxt.push(' Los pedidos realizados fuera de este periodo no serán tenidos en cuenta.');
  appendRichSection(body, 'II. Vigencia: ', vigenciaTxt);

  // III. PARTICIPANTES
  appendRichSection(body, 'III. Participantes: ', [data.textoSegmento, ' Adicionalmente, para participar se debe:']);
  const participantes = [
    '(i) Ser mayor de edad y residir en el Territorio.',
    '(ii) Tener una cuenta activa y vigente en la Plataforma Rappi.',
    '(iii) Realizar compras de los Productos Participantes durante la Vigencia, cumpliendo con la mecánica descrita.'
  ];
  participantes.forEach(p => appendRichListItem(body, p));

  // IV. TIENDA Y PRODUCTOS PARTICIPANTES
  appendRichSection(body, 'IV. Tienda y Productos Participantes: ', [
    'Participará la tienda virtual ', {text: `"${data.tiendaBase}"`, bold: true},
    ' (en adelante la "Tienda Participante") al interior de la Plataforma Rappi, ubicada dentro del Territorio. Participarán ',
    {text: data.productosParticipantes, bold: true}, ' disponibles en las secciones/verticales de ',
    {text: data.verticales, bold: true}, ' (en adelante, los "Productos Participantes").'
  ]);
  
  // V. MECÁNICA
  appendRichSection(body, 'V. Mecánica de la Actividad Promocional: ', [
    'La presente actividad es un concurso de destreza comercial. Serán seleccionados como ',
    {text: data.txtGanadoresPlural, bold: true}, ' los ',
    {text: `${data.numeroGanadoresLetras} (${data.numeroGanadores})`, bold: true},
    ' Usuarios/Consumidores que, durante la Vigencia, registren el ',
    {text: data.criterioGanador, bold: true},
    ' de Productos Participantes en la Tienda Participante.'
  ]);
  
  appendRichParagraph(body, ['Condiciones de participación:']);
  const condicionesMecanica = [
    'Solo se tendrán en cuenta órdenes finalizadas y entregadas. No suman órdenes canceladas, devueltas o con contracargo.',
    `Compra Mínima por orden: ${data.minimoCompraTexto}`,
    'No se tendrán en cuenta pagos realizados parcialmente con RappiCréditos, ni órdenes manipuladas fraudulentamente.'
  ];
  condicionesMecanica.forEach(c => appendRichListItem(body, c));

  // VI. CRITERIOS DE DESEMPATE
  appendRichSection(body, 'VI. Criterios de Desempate: ', ['En caso de presentarse un empate entre dos o más participantes, se definirá el ganador aplicando los siguientes criterios en estricto orden:']);
  const desempate = [
    data.txtDesempate1,
    'Si persiste el empate, ganará el Usuario/Consumidor que haya alcanzado la cifra registrada primero en el tiempo, según los registros del sistema de la Plataforma Rappi (criterio cronológico).'
  ];
  desempate.forEach(c => appendRichListItem(body, c));

  // VII. PREMIO
  appendRichSection(body, 'VII. Premio: ', [
    'Se entregarán ', {text: `${data.numeroGanadoresLetras} (${data.numeroGanadores})`, bold: true}, ' premios consistentes en:'
  ]);
  appendRichListItem(body, {text: data.listaPremios, bold: true});
  
  if (data.isPremioCreditos) {
    appendRichParagraph(body, [
      'Los Créditos serán cargados a la cuenta del ganador dentro de los ',
      {text: `${data.diasCargaPremioLetras} (${data.diasCargaPremio})`, bold: true},
      ' días hábiles siguientes a la confirmación del ganador. Los Créditos tendrán una vigencia de ',
      {text: `${data.vigenciaCreditosPremioLetras} (${data.vigenciaCreditosPremio})`, bold: true},
      ' días calendario contados a partir de su carga.'
    ]);
    
    appendRichParagraph(body, [
      'Se aclara que los Créditos otorgados podrán ser utilizados/redimidos ',
      {text: data.textoLugarRedencionPremio, bold: true}, ', únicamente dentro del Territorio.'
    ]);
    
    appendRichParagraph(body, [
      'Se aclara que los Créditos no tienen algún valor monetario, ni constituyen un medio de pago, instrumento crediticio o financiero. Los Créditos no pueden ser utilizados en las secciones denominadas "Cajero ATM" y "RappiFavor" de la Plataforma Rappi, ni para pagar el costo de envío, tarifa de servicio o propina.'
    ]);
  } else {
    appendRichParagraph(body, [
      'Los premios son personales e intransferibles. No son canjeables por dinero en efectivo ni por otros bienes o servicios. La gestión, logística, costos de envío y entrega efectiva de los premios correrán por cuenta y responsabilidad exclusiva del Organizador.'
    ]);
  }

  // VIII. CONDICIONES Y RESTRICCIONES
  appendRichSection(body, 'VIII. Condiciones y Restricciones: ', ['Podrán participar gratuitamente todas las personas naturales que sean Usuarios/Consumidores de la Plataforma Rappi que se encuentren en el Territorio y que cumplan las siguientes condiciones:']);
  
  const restricciones = [
    'Actividad válida únicamente para órdenes realizadas a través de la Tienda Participante, al interior de la Plataforma Rappi. Se aclara que no serán tenidas en cuenta las órdenes que se realicen a través de cualquier otra tienda y/o sección de la Plataforma Rappi.',
    'La presente Actividad Promocional se encuentra sujeta a los horarios de operación de los puntos de venta de la Tienda Participante.',
    'Actividad válida para todas las órdenes que cumplan las condiciones y restricciones establecidas en los presentes Términos y Condiciones.',
    'Actividad válida durante la Vigencia.',
    'El Beneficio obtenido en virtud de la presente Actividad Promocional no es acumulable con otras promociones exhibidas en la Plataforma Rappi.'
  ];
  
  if (data.limiteOrdenes && data.limiteOrdenes !== 'Sin límite') {
    restricciones.push(`Máximo ${data.limiteOrdenes} orden(es) por Usuario/Consumidor que sumen al acumulado.`);
  }
  
  if (data.hasMinCompra) {
    restricciones.splice(0, 0, `Actividad válida únicamente para órdenes por un valor igual o superior a $${data.minCompraNum.toLocaleString('es-CO')} pesos M/Cte (valor no incluye costo de envío ni propina).`);
  }
  
  restricciones.push('En caso de cancelación total o parcial de la orden, dicha orden no sumará al acumulado del Usuario/Consumidor.');
  
  if (data.hasCondiciones) {
    restricciones.push(data.condicionesEspeciales);
  }
  
  restricciones.forEach(r => appendRichListItem(body, r));

  // IX. MEDIO DE PAGO
  appendRichSection(body, 'IX. Medio de Pago: ', [data.textoMetodoPago]);

  // X. ANUNCIO Y ENTREGA
  appendRichSection(body, 'X. Anuncio y Entrega de Premios: ', [
    'Los ganadores serán anunciados el día ', {text: data.fechaAnuncio, bold: true},
    ' a través de los canales oficiales definidos por el Organizador (con el apoyo de difusión de la Plataforma Rappi, de ser requerido).'
  ]);
  
  const logistica = [
    `La gestión, logística, costos de envío y entrega efectiva de los premios correrán por cuenta y responsabilidad exclusiva de ${data.responsableEntrega}.`,
    'El ganador tendrá un plazo máximo de cinco (5) días hábiles para responder al contacto. Si no responde en dicho plazo, se entenderá que renuncia al premio y se procederá a contactar al siguiente participante en el ranking.'
  ];
  logistica.forEach(l => appendRichListItem(body, l));
  
  // XI. EXCLUSIONES Y FRAUDE
  appendRichSection(body, 'XI. Exclusiones y Fraude: ', ['El Organizador y Rappi se reservan el derecho de excluir a cualquier participante y cancelar la entrega del premio si detectan:']);
  const exclusiones = [
    '(i) Compras inusuales, autocompras, fraude, o uso de cuentas múltiples.',
    '(ii) Pagos disputados o contracargos.',
    '(iii) Violación a los Términos y Condiciones generales de la Plataforma Rappi.',
    '(iv) Cualquier comportamiento irregular que atente contra la naturaleza de la Actividad Promocional.'
  ];
  exclusiones.forEach(e => appendRichListItem(body, e));
  
  // XII. DECLARACIÓN
  appendRichSection(body, 'XII. Declaración sobre el Rol de Rappi: ', [
    'El Usuario/Consumidor reconoce y acepta que quien exhibe, ofrece, promociona y comercializa los productos adquiridos a través de la Plataforma Rappi es la Tienda Participante (Aliado Comercial). Rappi no comercializa productos, ya que es una plataforma de contacto. Rappi actúa únicamente como medio de difusión y comunicación de material publicitario para la realización de la Actividad Promocional.'
  ]);
  
  // XIII. LIMITACIÓN DE RESPONSABILIDAD
  appendRichSection(body, 'XIII. Limitación de Responsabilidad: ', [
    'La actividad se brinda como un medio de esparcimiento y ocio para el público en general. El Organizador y/o Rappi no asumen responsabilidad por ninguna consecuencia que resulte directa o indirectamente de cualquier acción o falta de acción que el participante emprenda.'
  ]);

  // XIV. TRATAMIENTO DE DATOS
  appendRichSection(body, 'XIV. Tratamiento de Datos Personales: ', [
    'El participante autoriza el tratamiento de sus datos personales para fines de la actividad. Si la entrega del premio requiere de un tercero (el Organizador), Rappi transferirá los datos de contacto necesarios bajo un acuerdo de transmisión de datos seguro, conforme a la Ley 1581 de 2012.'
  ]);

  // XV. CONTACTO Y PQR
  appendRichSection(body, 'XV. Contacto y Atención de PQR: ', [
    'Cualquier duda o inquietud sobre los alcances e interpretación de los presentes Términos y Condiciones, podrá ser consultada a través de los siguientes canales de contacto del Organizador:'
  ]);
  appendRichListItem(body, ['Teléfono de Atención: ', {text: data.telefonoContacto, bold: true}]);
  appendRichListItem(body, ['Correo electrónico: ', {text: data.emailContacto, bold: true}]);
  appendRichParagraph(body, ['Para dudas relacionadas con el funcionamiento de la Plataforma Rappi, el Usuario podrá contactar al Centro de Ayuda de Rappi a través de la aplicación.']);

  // XVI. MODIFICACIONES
  appendRichSection(body, 'XVI. Modificaciones: ', [
    'El Organizador se reserva el derecho de modificar los presentes Términos y Condiciones, así como suspender o cancelar la Actividad Promocional por motivos de fuerza mayor o caso fortuito, informando previamente a los usuarios/consumidores y sin perjuicio de los derechos adquiridos de los participantes.'
  ]);

  // XVII. JURISDICCIÓN
  appendRichSection(body, 'XVII. Jurisdicción y Solución de Conflictos: ', [
    'Los presentes Términos y Condiciones se regirán por las leyes de Colombia. Toda controversia surgida en razón de la Actividad Promocional o de los presentes Términos y Condiciones intentará ser resuelta por arreglo directo de las partes y/o a través de la conciliación como mecanismo alternativo de solución de conflictos. Si lo anterior no fuere posible, la controversia se someterá a la justicia ordinaria.'
  ]);
  
  appendRichParagraph(body, ['Se aclara que los presentes Términos y Condiciones se encuentran sujetos a los Términos y Condiciones de Uso de la Plataforma Rappi, disponibles en: https://legal.rappi.com/colombia/terminos-y-condiciones-de-uso-de-plataforma-rappi-2/.']);

  doc.saveAndClose();
  return doc;
}

// =================================================================
// SISTEMA DE EMAILS
// =================================================================
function sendEmailNotification(submitterEmail, docName, docUrl, data) {
  let ccEmails = [LEGAL_AUDIT_EMAIL];
  
  if (data.extraEmails && data.extraEmails.trim() !== '') {
    const extras = data.extraEmails.split(',').map(e => e.trim());
    ccEmails = ccEmails.concat(extras);
  }
  
  const subject = `✅ T&C Generados: ${data.nombreCampana} [${data.dinamica}]`;
  const htmlBody = getEmailTemplate(data, docUrl, docName);

  MailApp.sendEmail({
    to: submitterEmail,
    cc: ccEmails.join(','),
    subject: subject,
    htmlBody: htmlBody,
    name: "Motor Legal Rappi"
  });
}

function getEmailTemplate(data, docUrl, docName) {
  const color = (data.dinamica === 'CASHBACK') ? '#FF441F' : '#2962FF';
  const icon = (data.dinamica === 'CASHBACK') ? '💰' : '🏆';
  const typeLabel = (data.dinamica === 'CASHBACK') ? 'Campaña de Cashback' : 'Concurso / Top Spender';
  
  let tableRows = '';
  
  if (data.dinamica === 'CASHBACK') {
    tableRows = `
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🏪 Tienda:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.tiendaArchivo}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>💸 Beneficio:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.textoPorcentaje}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🛑 Tope:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.topeNumFmt}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>💰 Presupuesto:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.presupuestoNumFmt}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>📅 Vigencia:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.textoVigenciaEmail}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🌍 Territorio:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.territorioResumen}</td></tr>
    `;
  } else {
    tableRows = `
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>👑 Organizador:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.razonSocialOrganizador}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🏪 Tienda:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.tiendaBase}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🎯 Criterio:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.criterioGanador}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>👥 Ganadores:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.numeroGanadores}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>🎁 Premio:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.listaPremios}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>📢 Anuncio:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.fechaAnuncio}</td></tr>
      <tr><td style="padding:8px; border-bottom:1px solid #eee;"><strong>📅 Vigencia:</strong></td><td style="padding:8px; border-bottom:1px solid #eee;">${data.textoVigenciaEmail}</td></tr>
    `;
  }

  return `
    <div style="font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #f9fafb; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden;">
      <div style="background-color: ${color}; padding: 20px; text-align: center; color: white;">
        <h1 style="margin:0; font-size: 24px;">${icon} Documento Generado</h1>
        <p style="margin:5px 0 0 0; opacity: 0.9;">Motor Legal Rappi V2.1</p>
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
        <hr style="border: 0; border-top: 1px solid #e5e7eb; margin: 30px 0;">
        <p style="font-size: 12px; color: #9ca3af; text-align: center;">Generado por el Motor Legal de Rappi<br>${new Date().toLocaleString('es-CO')}</p>
      </div>
    </div>
  `;
}

// =================================================================
// UTILIDADES Y HELPERS
// =================================================================
function appendRichTextToParagraph(paragraph, parts) {
  if (!parts) return;
  if (!Array.isArray(parts)) parts = [parts];
  parts.forEach(part => {
    if (!part) return;
    let text = typeof part === 'string' ? part : part.text;
    if (!text) return;
    let isBold = typeof part === 'string' ? false : part.bold;
    let txtObj = paragraph.appendText(text);
    txtObj.setBold(isBold);
  });
}

function appendRichParagraph(body, parts) {
  const p = body.appendParagraph('');
  appendRichTextToParagraph(p, parts);
  p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
}

function appendRichSection(body, title, parts) {
  const p = body.appendParagraph('');
  p.appendText(title).setBold(true);
  if (Array.isArray(parts)) {
    parts.forEach(part => {
      if (!part) return;
      let text = typeof part === 'string' ? part : part.text;
      if (!text) return;
      let isBold = typeof part === 'string' ? false : part.bold;
      let txtObj = p.appendText(text);
      txtObj.setBold(isBold);
    });
  } else {
    if (parts) p.appendText(parts).setBold(false);
  }
  p.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
}

function appendRichListItem(body, parts) {
  if (!parts) return;
  if (typeof parts === 'string') {
    var item = body.appendListItem(parts);
    item.setGlyphType(DocumentApp.GlyphType.BULLET);
    item.setBold(false);
    item.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  } else {
    var item = body.appendListItem('');
    item.setBold(false);
    appendRichTextToParagraph(item, parts);
    item.setGlyphType(DocumentApp.GlyphType.BULLET);
    item.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  }
}

function saveLinkToSheet(url) {
  try {
    const form = FormApp.getActiveForm();
    const destId = form.getDestinationId();
    if (!destId) return;
    const ss = SpreadsheetApp.openById(destId);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let colIndex = headers.indexOf('Link T&C Generado (AUTO)');
    if (colIndex === -1) {
      colIndex = headers.length;
      sheet.getRange(1, colIndex + 1).setValue('Link T&C Generado (AUTO)');
    }
    sheet.getRange(lastRow, colIndex + 1).setValue(url);
  } catch (err) {
    Logger.log(err);
  }
}

function auditData(data) {
  const keys = Object.keys(data);
  for (const key of keys) {
    if (String(data[key]).includes('undefined')) {
      throw new Error(`Dato faltante: ${key}`);
    }
  }
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

function getResponsesAsMap(responses) {
  const map = {};
  responses.forEach(r => {
    map[r.getItem().getTitle().trim()] = r.getResponse();
  });
  return map;
}

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
// =================================================================
// GUARDAR FEEDBACK
// =================================================================
function saveFeedback(feedbackText) {
  try {
    const SHEET_ID = '1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI';
    const ss = SpreadsheetApp.openById(SHEET_ID);
    
    let sheet = ss.getSheetByName('Feedback');
    
    if (!sheet) {
      sheet = ss.insertSheet('Feedback');
      sheet.appendRow(['Timestamp', 'Feedback']);
      const headerRange = sheet.getRange(1, 1, 1, 2);
      headerRange.setBackground('#FF5C00');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setFontWeight('bold');
    }
    
    sheet.appendRow([new Date(), feedbackText]);
    
    Logger.log('✅ Feedback guardado');
    return true;
    
  } catch (error) {
    Logger.log('❌ Error guardando feedback: ' + error.message);
    throw error;
  }
}
// =================================================================
// CHATBOT SOPORTE CON GEMINI (MODELO DETECTADO: FLASH-LATEST)
// =================================================================
function askGemini(userQuestion) {
  // 👇 USA LA MISMA API KEY QUE TE FUNCIONÓ EN EL DIAGNÓSTICO
  const API_KEY = 'AIzaSyB-HWnobl4UwQDGnk3zqN915sLvIJ_qOnM'; 
  
  const systemPrompt = `
    Eres el asistente de soporte experto del "Motor Legal Rappi V2.2".
    Ayudas a equipos de marketing a crear Términos y Condiciones.
    
    TUS CONOCIMIENTOS:
    - La herramienta genera T&C para "Cashback" y "Concursos" (Top Spender).
    - Actualmente solo funciona para Colombia.
    - El usuario llena el formulario > Genera Doc > Copia a Squarespace.
    - Si es Concurso con premio físico entregado por el Organizador, recuerda que necesitan el "Acuerdo de Transferencia de Datos".
    
    TONO:
    - Amable, breve, eficiente y profesional.
    - Responde siempre en español.
  `;

  // ✅ USAMOS EL MODELO QUE APARECIÓ EN TU LISTA: 'gemini-flash-latest'
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${API_KEY}`;
  
  const payload = {
    "contents": [{
      "parts": [{
        "text": systemPrompt + "\n\nPregunta del usuario: " + userQuestion
      }]
    }]
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      Logger.log("Error Gemini: " + JSON.stringify(json.error));
      
      // Si falla por cuota (429), intentamos con el modelo Lite como respaldo
      if (json.error.code === 429) {
         return "Lo siento, estoy recibiendo muchas consultas. Intenta con gemini-2.0-flash-lite en el código.";
      }
      return "Error técnico (" + json.error.message + ").";
    }
    
    return json.candidates[0].content.parts[0].text;
    
  } catch (e) {
    return "Error de conexión. Intenta de nuevo.";
  }
}
function testChatbot() {
  Logger.log("⏳ Enviando pregunta a Gemini...");
  
  // Hacemos una pregunta de prueba
  const respuesta = askGemini("Hola, quiero crear un concurso para Rappi, ¿qué debo hacer?");
  
  Logger.log("🤖 RESPUESTA RECIBIDA:");
  Logger.log(respuesta);
}// =================================================================
// RAPPIMIND V2.2 — TEMPLATE ENGINE + ADMIN PANEL
// Pegar TODO al final de Código.gs (NO borrar nada existente)
// =================================================================

// -----------------------------------------------------------------
// CONSTANTES DEL TEMPLATE ENGINE
// -----------------------------------------------------------------
const REGISTRY_SHEET_NAME = 'Template_Registry';
const FIELDS_SHEET_NAME = 'Template_Fields';
const AUDIT_SHEET_ID = '1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI';
const ADMIN_EMAILS_LIST = ['juan.gallego@rappi.com', 'david.gaviria@rappi.com'];

// =================================================================
// 1. CORE ENGINE V2 — NUEVO PUNTO DE ENTRADA
//    Reemplaza las 2 líneas en processWebPayload
// =================================================================
function coreEngineV2(payload, submitterEmail) {
  const vars = mapWebToEngine(payload);
  
  // Intentar Template Engine primero
  try {
    const countryCode = payload.countryCode || 'CO';
    const campaignType = vars['Tipo de Dinámica'] || 'Cashback';
    
    const registry = _getTemplateRegistry();
    const template = registry.find(r => 
      r.country_code === countryCode && 
      r.campaign_type === campaignType && 
      r.status === 'active'
    );
    
    if (template) {
      Logger.log('🚀 Template Engine: Usando template ' + countryCode + '/' + campaignType);
      return _generateWithTemplate(template, vars, submitterEmail);
    }
  } catch (e) {
    Logger.log('⚠️ Template Engine falló, usando legacy: ' + e.message);
  }
  
  // Fallback al motor legacy
  Logger.log('🔄 Usando motor legacy (hardcoded)');
  return coreEngine(vars, submitterEmail);
}

// =================================================================
// 2. GENERAR CON TEMPLATE — Clona doc + reemplaza placeholders
// =================================================================
function _generateWithTemplate(template, vars, submitterEmail) {
  vars['Email Generador'] = submitterEmail;
  validateDates(vars);
  
  const tipoDinamica = vars['Tipo de Dinámica'] || 'Cashback';
  let data = {};
  
  if (tipoDinamica === 'Concurso Mayor Comprador') {
    data = procesarConcurso(vars);
  } else {
    data = procesarCashback(vars);
  }
  
  auditData(data);
  
  // Construir mapa de placeholders desde data
  const placeholders = _buildPlaceholderMap(data);
  
  // Clonar template
  const templateFile = DriveApp.getFileById(template.template_doc_id);
  const newFile = templateFile.makeCopy(data.docName);
  
  // Mover a carpeta si existe
  if (DRIVE_FOLDER_ID) {
    try {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      folder.addFile(newFile);
      DriveApp.getRootFolder().removeFile(newFile);
    } catch (fError) {
      Logger.log('Nota: No se movió a carpeta.');
    }
  }
  
  // Abrir y reemplazar
  const doc = DocumentApp.openById(newFile.getId());
  const body = doc.getBody();
  
  // Reemplazar todos los placeholders
  Object.entries(placeholders).forEach(([key, value]) => {
    const safeValue = (value !== null && value !== undefined) ? String(value) : '';
    // Escapar regex chars en el key
    const escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    body.replaceText(escapedKey, safeValue);
  });
  
  // Limpiar placeholders no utilizados
  body.replaceText('\\{\\{[A-Z_0-9]+\\}\\}', '');
  
  doc.saveAndClose();
  
  const publicUrl = setPublicViewPermissions(doc);
  
  // Auditoría
  saveResponseToSheet(vars, publicUrl);
  
  // Email
  sendEmailNotification(submitterEmail, data.docName, publicUrl, data);
  
  return {
    docUrl: publicUrl,
    docName: data.docName
  };
}

// =================================================================
// 3. MAPA DE PLACEHOLDERS — Convierte data object a {{KEY}}: value
// =================================================================
function _buildPlaceholderMap(data) {
  const map = {};
  
  // --- COMUNES ---
  map['{{NOMBRE_CAMPANA_UPPER}}'] = (data.nombreCampana || '').toUpperCase();
  map['{{NOMBRE_CAMPANA}}'] = data.nombreCampana || '';
  map['{{NOMBRE_CAMPANA_LOWER}}'] = data.nombreCampanaLower || data.nombreCampana || '';
  map['{{TEXTO_TERRITORIO}}'] = data.textoTerritorio || '';
  map['{{FECHA_INICIO}}'] = data.startFmtDate || '';
  map['{{FECHA_FIN}}'] = data.endFmtDate || '';
  map['{{HORA_INICIO}}'] = data.startFmtTime || '';
  map['{{HORA_FIN}}'] = data.endFmtTime || '';
  map['{{TEXTO_SEGMENTO}}'] = data.textoSegmento || '';
  map['{{TEXTO_METODO_PAGO}}'] = data.textoMetodoPago || '';
  map['{{TIENDA_BASE}}'] = data.tiendaBase || '';
  
  // --- TIENDA (variantes) ---
  map['{{REF_TIENDA}}'] = data.txtRefTienda || 'la Tienda Participante';
  map['{{TIENDA_DISPLAY}}'] = data.tiendaDisplay || `"${data.tiendaBase}"`;
  map['{{DEFINICION_TIENDA}}'] = data.txtDefinicionTienda || ' (en adelante la "Tienda Participante") ';
  map['{{TITULO_TIENDA}}'] = data.txtTituloTienda || 'IV. Tienda Participante: ';
  map['{{DECLARACION_TIENDA}}'] = data.txtDeclaracionTienda || 'es la Tienda Participante';
  
  // --- CASHBACK ---
  if (data.dinamica === 'CASHBACK') {
    map['{{TEXTO_PORCENTAJE}}'] = data.textoPorcentaje || '';
    map['{{TOPE_LETRAS}}'] = data.topeLetras || '';
    map['{{TOPE_NUM}}'] = data.topeNumFmt || '';
    map['{{PRESUPUESTO_LETRAS}}'] = data.presupuestoLetras || '';
    map['{{PRESUPUESTO_NUM}}'] = data.presupuestoNumFmt || '';
    map['{{UMBRAL_LETRAS}}'] = data.umbralCompraLetras || '';
    map['{{UMBRAL_NUM}}'] = data.umbralCompraFmt || '';
    map['{{TEXTO_CARGA}}'] = data.txtCargaCompleto || '';
    map['{{TEXTO_VIGENCIA_CREDITOS}}'] = data.txtVigenciaCreditos || '';
    map['{{TEXTO_LUGAR_REDENCION}}'] = data.textoLugarRedencion || '';
    map['{{LIMITE_ORDENES}}'] = data.limiteOrdenes || '1';
    map['{{TEXTO_ORDENES}}'] = data.txtOrdenesGramatica || 'orden';
    map['{{MIN_COMPRA}}'] = data.minCompra || '';
    map['{{CONDICIONES_ESPECIALES}}'] = data.condicionesEspeciales || '';
  }
  
  // --- CONCURSO ---
  if (data.dinamica === 'TOP_SPENDER') {
    map['{{ORGANIZADOR}}'] = data.razonSocialOrganizador || '';
    map['{{TELEFONO_CONTACTO}}'] = data.telefonoContacto || '';
    map['{{EMAIL_CONTACTO}}'] = data.emailContacto || '';
    map['{{NUM_GANADORES}}'] = String(data.numeroGanadores || 1);
    map['{{NUM_GANADORES_LETRAS}}'] = data.numeroGanadoresLetras || '';
    map['{{PLURAL_GANADORES}}'] = data.txtGanadoresPlural || 'ganador';
    map['{{CRITERIO_GANADOR}}'] = data.criterioGanador || '';
    map['{{DESEMPATE_1}}'] = data.txtDesempate1 || '';
    map['{{VERTICALES}}'] = data.verticales || '';
    map['{{PRODUCTOS_PARTICIPANTES}}'] = data.productosParticipantes || '';
    map['{{LISTA_PREMIOS}}'] = data.listaPremios || '';
    map['{{RESPONSABLE_ENTREGA}}'] = data.responsableEntrega || '';
    map['{{FECHA_ANUNCIO}}'] = data.fechaAnuncio || '';
    map['{{MINIMO_COMPRA_TEXTO}}'] = data.minimoCompraTexto || '';
    map['{{CONDICIONES_ESPECIALES}}'] = data.condicionesEspeciales || '';
    
    // Créditos premio
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

// =================================================================
// 4. LEER TEMPLATE REGISTRY DESDE SHEET
// =================================================================
function _getTemplateRegistry() {
  try {
    const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    const sheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
    if (!sheet) return [];
    return _getSheetAsObjects(sheet);
  } catch (e) {
    Logger.log('Error leyendo registry: ' + e.message);
    return [];
  }
}

function _getSheetAsObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// =================================================================
// 5. SEED TEMPLATES — Ejecutar UNA vez para crear Google Docs
// =================================================================

// ---- MASTER: Corre esto y se hace todo ----
function setupTemplateEngine() {
  Logger.log('🚀 === SETUP TEMPLATE ENGINE ===');
  
  // Paso 1: Crear template Cashback
  const cashbackDocId = seedCashbackTemplate();
  Logger.log('✅ Template Cashback creado: ' + cashbackDocId);
  
  // Paso 2: Crear template Concurso
  const concursoDocId = seedConcursoTemplate();
  Logger.log('✅ Template Concurso creado: ' + concursoDocId);
  
  // Paso 3: Registrar en Template_Registry
  _seedRegistry(cashbackDocId, concursoDocId);
  Logger.log('✅ Registry creado');
  
  // Paso 4: Crear campos en Template_Fields
  seedColombiaFields();
  Logger.log('✅ Campos creados');
  
  Logger.log('🎉 === SETUP COMPLETO ===');
  Logger.log('📋 Cashback Doc ID: ' + cashbackDocId);
  Logger.log('📋 Concurso Doc ID: ' + concursoDocId);
  Logger.log('');
  Logger.log('👉 SIGUIENTE PASO: Cambia 2 líneas en processWebPayload');
  Logger.log('   ANTES:  const v67Map = mapWebToEngine(payload);');
  Logger.log('           const result = coreEngine(v67Map, payload.userEmail);');
  Logger.log('   DESPUÉS: const result = coreEngineV2(payload, payload.userEmail);');
  
  return { cashbackDocId, concursoDocId };
}

// ---- CASHBACK TEMPLATE ----
function seedCashbackTemplate() {
  const doc = DocumentApp.create('Template_CO_Cashback');
  const body = doc.getBody();
  
  // TÍTULO
  const titulo = body.appendParagraph('TÉRMINOS Y CONDICIONES – CAMPAÑA "{{NOMBRE_CAMPANA_UPPER}}"');
  titulo.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');
  
  // INTRODUCCIÓN
  appendRichParagraph(body, [
    'Por medio del presente documento se dan a conocer los términos y condiciones de la campaña denominada "',
    {text: '{{NOMBRE_CAMPANA_LOWER}}', bold: true},
    '" (en adelante la "Campaña"). La participación en la Campaña constituye la aceptación total e incondicional de los presentes Términos y Condiciones, los cuales resultan definitivos y vinculantes para los Usuarios/Consumidores participantes, que cumplan con los requisitos aquí dispuestos.'
  ]);
  
  // I. TERRITORIO
  appendRichSection(body, 'I. Territorio: ', [
    'La Campaña será válida únicamente para las órdenes realizadas dentro de las zonas de cobertura de {{REF_TIENDA}} en la Plataforma Rappi, en ',
    {text: '{{TEXTO_TERRITORIO}}', bold: true}, '.'
  ]);
  
  // II. VIGENCIA
  appendRichSection(body, 'II. Vigencia: ', [
    'Campaña válida desde las ', {text: '{{HORA_INICIO}}', bold: true},
    ' del ', {text: '{{FECHA_INICIO}}', bold: true},
    ' hasta las ', {text: '{{HORA_FIN}}', bold: true},
    ' del ', {text: '{{FECHA_FIN}}', bold: true},
    ' y/o hasta agotar existencias, lo que primero ocurra. Para efectos de la Campaña, se ha establecido un valor total máximo de ',
    {text: '{{PRESUPUESTO_LETRAS}}', bold: true}, ' (', {text: '{{PRESUPUESTO_NUM}}', bold: true},
    ') de Créditos a ser entregados como Cashback a los Usuarios/Consumidores que participen en la Campaña ("Existencias"). Por consiguiente, el agotamiento de las Existencias con anterioridad al final de la Vigencia indicada implicará la terminación de la Campaña.'
  ]);
  
  // III. TIPO DE USUARIOS
  appendRichSection(body, 'III. Tipo de Usuarios Participantes: ', ['{{TEXTO_SEGMENTO}}']);
  
  // IV. TIENDA PARTICIPANTE
  appendRichSection(body, '{{TITULO_TIENDA}}', [
    'Participarán todas las tiendas virtuales de ', {text: '{{TIENDA_DISPLAY}}', bold: true},
    '{{DEFINICION_TIENDA}}al interior de la Plataforma Rappi ubicadas dentro del Territorio.'
  ]);
  
  // V. PRODUCTOS PARTICIPANTES
  appendRichSection(body, 'V. Productos Participantes: ', [
    'Participarán todos los productos que hacen parte del catálogo de {{REF_TIENDA}} al interior de la Plataforma Rappi.'
  ]);
  
  // VI. BENEFICIO
  const pBeneficio = body.appendParagraph('');
  pBeneficio.appendText('VI. Beneficio: ').setBold(true);
  appendRichTextToParagraph(pBeneficio, [
    'Los Usuarios/Consumidores Participantes que durante la Vigencia de la Campaña compren cualquiera de los Productos Participantes de {{REF_TIENDA}} recibirán en Créditos el ',
    {text: '{{TEXTO_PORCENTAJE}}', bold: true},
    ' del valor de dichos Productos Participantes (en adelante el "Cashback"). Dichos Créditos serán cargados a su cuenta al interior de la Plataforma Rappi. Se aclara que el monto máximo del Cashback que se otorgará en Créditos es de ',
    {text: '{{TOPE_LETRAS}}', bold: true}, ' (', {text: '{{TOPE_NUM}}', bold: true}, ')',
    ' Créditos. Por lo tanto, en caso de que el Usuario/Consumidor realice una compra en {{REF_TIENDA}} por un valor superior a los ',
    {text: '{{UMBRAL_LETRAS}}', bold: true}, ' pesos M/Cte ($', {text: '{{UMBRAL_NUM}}', bold: true},
    '), recibirá un monto máximo en Créditos de ',
    {text: '{{TOPE_LETRAS}}', bold: true}, ' (', {text: '{{TOPE_NUM}}', bold: true}, ').',
    ' Los Créditos serán cargados a la cuenta de los Usuarios/Consumidores Participantes ',
    {text: '{{TEXTO_CARGA}}', bold: true},
    '. Los Créditos {{TEXTO_VIGENCIA_CREDITOS}}, entendiéndose que si el Usuario/Consumidor no hace uso de ellos dentro del término estipulado los perderá, sin poder hacer uso de ellos posteriormente.'
  ]);
  pBeneficio.setAlignment(DocumentApp.HorizontalAlignment.JUSTIFY);
  
  // VII. CONDICIONES Y RESTRICCIONES
  appendRichSection(body, 'VII. Condiciones y Restricciones: ', [
    'Podrán participar gratuitamente todas las personas naturales que sean Usuarios/Consumidores de la Plataforma Rappi que se encuentren en el Territorio y que cumplan las siguientes condiciones:'
  ]);
  
  const restricciones = [
    'Campaña válida únicamente para órdenes realizadas a través de {{REF_TIENDA}}, al interior de la Plataforma Rappi. Se aclara que no serán tenidas en cuenta las órdenes que se realicen a través de cualquier otra tienda y/o sección de la Plataforma Rappi.',
    'La presente Campaña se encuentra sujeta a los horarios de operación de los puntos de venta de {{REF_TIENDA}}, que ofrecen y exhiben sus productos al interior de la Plataforma Rappi.',
    'Los descuentos de los productos y/o servicios objeto de la Campaña no son intercambiables ni transferibles.',
    'Campaña válida para todas las órdenes que cumplan las condiciones y restricciones establecidas en los presentes Términos y Condiciones.',
    'Campaña válida durante la Vigencia y/o hasta agotar existencias, lo que primero ocurra.',
    'El Beneficio obtenido en virtud de la presente Campaña no es acumulable con otras promociones exhibidas en la Plataforma Rappi.'
  ];
  restricciones.forEach(r => appendRichListItem(body, r));
  
  appendRichListItem(body, ['Máximo ', {text: '{{LIMITE_ORDENES}}', bold: true}, ' {{TEXTO_ORDENES}} por Usuario/Consumidor.']);
  
  appendRichListItem(body, ['Se aclara que el monto máximo de Créditos a recibir por el Usuario/Consumidor Participante es de ',
    {text: '{{TOPE_LETRAS}}', bold: true}, ' (', {text: '{{TOPE_NUM}}', bold: true},
    '), de acuerdo con lo indicado en la sección VI (Beneficio) de los presentes Términos y Condiciones.']);
  
  appendRichListItem(body, 'En caso de cancelación total o parcial de la orden, el Usuario/Consumidor Participante no tendrá derecho al Beneficio.');
  appendRichListItem(body, '{{TEXTO_LUGAR_REDENCION}}');
  appendRichListItem(body, 'Se aclara que los Créditos no tienen algún valor monetario, ni constituyen un medio de pago, instrumento crediticio o financiero y la finalidad de su redención no es recibir una cantidad de dinero en efectivo.');
  appendRichListItem(body, 'Se aclara que los Créditos no pueden ser utilizados en las secciones denominadas "Cajero ATM" y "RappiFavor" de la Plataforma Rappi, ni tampoco pueden ser utilizados para pagar el valor del costo de envío, la tarifa de servicio de un pedido realizado a través de la Plataforma Rappi y/o la propina otorgada de forma voluntaria a los repartidores independientes.');
  
  // VIII. MEDIO DE PAGO
  appendRichSection(body, 'VIII. Medio de Pago: ', ['{{TEXTO_METODO_PAGO}}']);
  
  // IX. MODIFICACIONES
  appendRichSection(body, 'IX. Modificaciones e Interpretación: ', [
    'Rappi se reserva el derecho de cancelar órdenes si detecta un comportamiento irregular por parte del Usuario/Consumidor en la Plataforma Rappi. Rappi se reserva el derecho de rechazar y cancelar cualquier orden que, por sus características, Rappi determine que no aplica para el Beneficio de la presente Campaña, sin previo aviso al Usuario/Consumidor.'
  ]);
  
  // X. DECLARACIÓN
  appendRichSection(body, 'X. Declaración: ', [
    'El Usuario/Consumidor reconoce y acepta que quien exhibe, ofrece, promociona y comercializa los productos adquiridos a través de la Plataforma Rappi {{DECLARACION_TIENDA}}. Rappi no comercializa productos puesto que es solo una plataforma tecnológica de contacto.'
  ]);
  
  appendRichParagraph(body, [
    'Se aclara que los presentes Términos y Condiciones se encuentran sujetos a los Términos y Condiciones de Uso de la Plataforma Rappi, los cuales se encuentran en la siguiente dirección electrónica: https://legal.rappi.com.co/colombia/terminos-y-condiciones-de-uso-de-plataforma-rappi-2/.'
  ]);
  
  // XI. JURISDICCIÓN
  appendRichSection(body, 'XI. Jurisdicción y Solución de conflictos: ', [
    'Los presentes Términos y Condiciones se regirán por las leyes de Colombia. Toda controversia surgida en razón de la Campaña o de los presentes Términos y Condiciones intentará ser resuelta por arreglo directo de las partes y/o a través de la conciliación como mecanismo alternativo de solución de conflictos. Si lo anterior no fuere posible, la controversia se someterá a la justicia ordinaria.'
  ]);
  
  doc.saveAndClose();
  Logger.log('📄 Template Cashback creado: ' + doc.getId());
  return doc.getId();
}

// ---- CONCURSO TEMPLATE ----
function seedConcursoTemplate() {
  const doc = DocumentApp.create('Template_CO_Concurso');
  const body = doc.getBody();
  
  // TÍTULO
  const titulo = body.appendParagraph('TÉRMINOS Y CONDICIONES – ACTIVIDAD PROMOCIONAL "{{NOMBRE_CAMPANA_UPPER}}"');
  titulo.setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');
  
  // INTRODUCCIÓN
  appendRichParagraph(body, [
    'Por medio del presente documento se dan a conocer los términos y condiciones de la actividad promocional denominada "',
    {text: '{{NOMBRE_CAMPANA}}', bold: true},
    '" (en adelante la "Actividad Promocional"). La sola participación en esta actividad implica el conocimiento y aceptación total e incondicional de estos Términos y Condiciones, no pudiendo alegar el participante el desconocimiento de estos.'
  ]);
  
  // I. ORGANIZADOR Y TERRITORIO
  appendRichSection(body, 'I. Organizador y Territorio: ', [
    'La presente Actividad Promocional es organizada y financiada exclusivamente por ',
    {text: '{{ORGANIZADOR}}', bold: true},
    ' (en adelante el "Organizador"), quien es el responsable de la entrega del beneficio. RAPPI S.A.S. actúa únicamente como plataforma de contacto y medio de difusión. La Actividad Promocional será válida únicamente para pedidos realizados a través de la Plataforma Rappi en las zonas de cobertura habilitadas en ',
    {text: '{{TEXTO_TERRITORIO}}', bold: true},
    ' (en adelante el "Territorio").'
  ]);
  
  // II. VIGENCIA
  appendRichSection(body, 'II. Vigencia: ', [
    'La Actividad Promocional estará vigente desde las ',
    {text: '{{HORA_INICIO}}', bold: true}, ' del ', {text: '{{FECHA_INICIO}}', bold: true},
    ' hasta las ', {text: '{{HORA_FIN}}', bold: true}, ' del ', {text: '{{FECHA_FIN}}', bold: true},
    ' (en adelante, la "Vigencia"). Los pedidos realizados fuera de este periodo no serán tenidos en cuenta.'
  ]);
  
  // III. PARTICIPANTES
  appendRichSection(body, 'III. Participantes: ', ['{{TEXTO_SEGMENTO}} Adicionalmente, para participar se debe:']);
  appendRichListItem(body, '(i) Ser mayor de edad y residir en el Territorio.');
  appendRichListItem(body, '(ii) Tener una cuenta activa y vigente en la Plataforma Rappi.');
  appendRichListItem(body, '(iii) Realizar compras de los Productos Participantes durante la Vigencia, cumpliendo con la mecánica descrita.');
  
  // IV. TIENDA Y PRODUCTOS
  appendRichSection(body, 'IV. Tienda y Productos Participantes: ', [
    'Participará la tienda virtual ', {text: '"{{TIENDA_BASE}}"', bold: true},
    ' (en adelante la "Tienda Participante") al interior de la Plataforma Rappi, ubicada dentro del Territorio. Participarán ',
    {text: '{{PRODUCTOS_PARTICIPANTES}}', bold: true}, ' disponibles en las secciones/verticales de ',
    {text: '{{VERTICALES}}', bold: true}, ' (en adelante, los "Productos Participantes").'
  ]);
  
  // V. MECÁNICA
  appendRichSection(body, 'V. Mecánica de la Actividad Promocional: ', [
    'La presente actividad es un concurso de destreza comercial. Serán seleccionados como ',
    {text: '{{PLURAL_GANADORES}}', bold: true}, ' los ',
    {text: '{{NUM_GANADORES_LETRAS}} ({{NUM_GANADORES}})', bold: true},
    ' Usuarios/Consumidores que, durante la Vigencia, registren el ',
    {text: '{{CRITERIO_GANADOR}}', bold: true},
    ' de Productos Participantes en la Tienda Participante.'
  ]);
  
  appendRichParagraph(body, ['Condiciones de participación:']);
  appendRichListItem(body, 'Solo se tendrán en cuenta órdenes finalizadas y entregadas. No suman órdenes canceladas, devueltas o con contracargo.');
  appendRichListItem(body, 'Compra Mínima por orden: {{MINIMO_COMPRA_TEXTO}}');
  appendRichListItem(body, 'No se tendrán en cuenta pagos realizados parcialmente con RappiCréditos, ni órdenes manipuladas fraudulentamente.');
  
  // VI. DESEMPATE
  appendRichSection(body, 'VI. Criterios de Desempate: ', [
    'En caso de presentarse un empate entre dos o más participantes, se definirá el ganador aplicando los siguientes criterios en estricto orden:'
  ]);
  appendRichListItem(body, '{{DESEMPATE_1}}');
  appendRichListItem(body, 'Si persiste el empate, ganará el Usuario/Consumidor que haya alcanzado la cifra registrada primero en el tiempo, según los registros del sistema de la Plataforma Rappi (criterio cronológico).');
  
  // VII. PREMIO
  appendRichSection(body, 'VII. Premio: ', [
    'Se entregarán ', {text: '{{NUM_GANADORES_LETRAS}} ({{NUM_GANADORES}})', bold: true}, ' premios consistentes en:'
  ]);
  appendRichListItem(body, {text: '{{LISTA_PREMIOS}}', bold: true});
  
  // Bloque créditos (se limpia si es premio físico)
  appendRichParagraph(body, [
    'Los Créditos serán cargados a la cuenta del ganador dentro de los ',
    {text: '{{DIAS_CARGA_LETRAS}} ({{DIAS_CARGA_NUM}})', bold: true},
    ' días hábiles siguientes a la confirmación del ganador. Los Créditos tendrán una vigencia de ',
    {text: '{{VIGENCIA_CREDITOS_LETRAS}} ({{VIGENCIA_CREDITOS_NUM}})', bold: true},
    ' días calendario contados a partir de su carga.'
  ]);
  
  appendRichParagraph(body, [
    'Se aclara que los Créditos otorgados podrán ser utilizados/redimidos ',
    {text: '{{LUGAR_REDENCION_PREMIO}}', bold: true}, ', únicamente dentro del Territorio.'
  ]);
  
  appendRichParagraph(body, [
    'Se aclara que los Créditos no tienen algún valor monetario, ni constituyen un medio de pago, instrumento crediticio o financiero. Los Créditos no pueden ser utilizados en las secciones denominadas "Cajero ATM" y "RappiFavor" de la Plataforma Rappi, ni para pagar el costo de envío, tarifa de servicio o propina.'
  ]);
  
  // VIII. CONDICIONES
  appendRichSection(body, 'VIII. Condiciones y Restricciones: ', [
    'Podrán participar gratuitamente todas las personas naturales que sean Usuarios/Consumidores de la Plataforma Rappi que se encuentren en el Territorio y que cumplan las siguientes condiciones:'
  ]);
  
  appendRichListItem(body, 'Actividad válida únicamente para órdenes realizadas a través de la Tienda Participante, al interior de la Plataforma Rappi. Se aclara que no serán tenidas en cuenta las órdenes que se realicen a través de cualquier otra tienda y/o sección de la Plataforma Rappi.');
  appendRichListItem(body, 'La presente Actividad Promocional se encuentra sujeta a los horarios de operación de los puntos de venta de la Tienda Participante.');
  appendRichListItem(body, 'Actividad válida para todas las órdenes que cumplan las condiciones y restricciones establecidas en los presentes Términos y Condiciones.');
  appendRichListItem(body, 'Actividad válida durante la Vigencia.');
  appendRichListItem(body, 'El Beneficio obtenido en virtud de la presente Actividad Promocional no es acumulable con otras promociones exhibidas en la Plataforma Rappi.');
  appendRichListItem(body, 'En caso de cancelación total o parcial de la orden, dicha orden no sumará al acumulado del Usuario/Consumidor.');
  
  // IX. MEDIO DE PAGO
  appendRichSection(body, 'IX. Medio de Pago: ', ['{{TEXTO_METODO_PAGO}}']);
  
  // X. ANUNCIO
  appendRichSection(body, 'X. Anuncio y Entrega de Premios: ', [
    'Los ganadores serán anunciados el día ', {text: '{{FECHA_ANUNCIO}}', bold: true},
    ' a través de los canales oficiales definidos por el Organizador (con el apoyo de difusión de la Plataforma Rappi, de ser requerido).'
  ]);
  appendRichListItem(body, 'La gestión, logística, costos de envío y entrega efectiva de los premios correrán por cuenta y responsabilidad exclusiva de {{RESPONSABLE_ENTREGA}}.');
  appendRichListItem(body, 'El ganador tendrá un plazo máximo de cinco (5) días hábiles para responder al contacto. Si no responde en dicho plazo, se entenderá que renuncia al premio y se procederá a contactar al siguiente participante en el ranking.');
  
  // XI. EXCLUSIONES
  appendRichSection(body, 'XI. Exclusiones y Fraude: ', [
    'El Organizador y Rappi se reservan el derecho de excluir a cualquier participante y cancelar la entrega del premio si detectan:'
  ]);
  appendRichListItem(body, '(i) Compras inusuales, autocompras, fraude, o uso de cuentas múltiples.');
  appendRichListItem(body, '(ii) Pagos disputados o contracargos.');
  appendRichListItem(body, '(iii) Violación a los Términos y Condiciones generales de la Plataforma Rappi.');
  appendRichListItem(body, '(iv) Cualquier comportamiento irregular que atente contra la naturaleza de la Actividad Promocional.');
  
  // XII. DECLARACIÓN
  appendRichSection(body, 'XII. Declaración sobre el Rol de Rappi: ', [
    'El Usuario/Consumidor reconoce y acepta que quien exhibe, ofrece, promociona y comercializa los productos adquiridos a través de la Plataforma Rappi es la Tienda Participante (Aliado Comercial). Rappi no comercializa productos, ya que es una plataforma de contacto. Rappi actúa únicamente como medio de difusión y comunicación de material publicitario para la realización de la Actividad Promocional.'
  ]);
  
  // XIII. LIMITACIÓN
  appendRichSection(body, 'XIII. Limitación de Responsabilidad: ', [
    'La actividad se brinda como un medio de esparcimiento y ocio para el público en general. El Organizador y/o Rappi no asumen responsabilidad por ninguna consecuencia que resulte directa o indirectamente de cualquier acción o falta de acción que el participante emprenda.'
  ]);
  
  // XIV. DATOS
  appendRichSection(body, 'XIV. Tratamiento de Datos Personales: ', [
    'El participante autoriza el tratamiento de sus datos personales para fines de la actividad. Si la entrega del premio requiere de un tercero (el Organizador), Rappi transferirá los datos de contacto necesarios bajo un acuerdo de transmisión de datos seguro, conforme a la Ley 1581 de 2012.'
  ]);
  
  // XV. CONTACTO
  appendRichSection(body, 'XV. Contacto y Atención de PQR: ', [
    'Cualquier duda o inquietud sobre los alcances e interpretación de los presentes Términos y Condiciones, podrá ser consultada a través de los siguientes canales de contacto del Organizador:'
  ]);
  appendRichListItem(body, ['Teléfono de Atención: ', {text: '{{TELEFONO_CONTACTO}}', bold: true}]);
  appendRichListItem(body, ['Correo electrónico: ', {text: '{{EMAIL_CONTACTO}}', bold: true}]);
  appendRichParagraph(body, ['Para dudas relacionadas con el funcionamiento de la Plataforma Rappi, el Usuario podrá contactar al Centro de Ayuda de Rappi a través de la aplicación.']);
  
  // XVI. MODIFICACIONES
  appendRichSection(body, 'XVI. Modificaciones: ', [
    'El Organizador se reserva el derecho de modificar los presentes Términos y Condiciones, así como suspender o cancelar la Actividad Promocional por motivos de fuerza mayor o caso fortuito, informando previamente a los usuarios/consumidores y sin perjuicio de los derechos adquiridos de los participantes.'
  ]);
  
  // XVII. JURISDICCIÓN
  appendRichSection(body, 'XVII. Jurisdicción y Solución de Conflictos: ', [
    'Los presentes Términos y Condiciones se regirán por las leyes de Colombia. Toda controversia surgida en razón de la Actividad Promocional o de los presentes Términos y Condiciones intentará ser resuelta por arreglo directo de las partes y/o a través de la conciliación como mecanismo alternativo de solución de conflictos. Si lo anterior no fuere posible, la controversia se someterá a la justicia ordinaria.'
  ]);
  
  appendRichParagraph(body, [
    'Se aclara que los presentes Términos y Condiciones se encuentran sujetos a los Términos y Condiciones de Uso de la Plataforma Rappi, disponibles en: https://legal.rappi.com/colombia/terminos-y-condiciones-de-uso-de-plataforma-rappi-2/.'
  ]);
  
  doc.saveAndClose();
  Logger.log('📄 Template Concurso creado: ' + doc.getId());
  return doc.getId();
}

// ---- SEED REGISTRY ----
function _seedRegistry(cashbackDocId, concursoDocId) {
  const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
  let sheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(REGISTRY_SHEET_NAME);
    const headers = ['country_code', 'country_name', 'campaign_type', 'template_doc_id', 'version', 'status', 'currency_code', 'currency_symbol', 'legal_owner', 'last_updated', 'notes'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  const today = new Date().toISOString().split('T')[0];
  
  sheet.appendRow(['CO', 'Colombia', 'Cashback', cashbackDocId, '1.0', 'active', 'COP', '$', 'juan.gallego@rappi.com', today, 'Template base generado automáticamente']);
  sheet.appendRow(['CO', 'Colombia', 'Concurso Mayor Comprador', concursoDocId, '1.0', 'active', 'COP', '$', 'juan.gallego@rappi.com', today, 'Template base generado automáticamente']);
}

// ---- SEED FIELDS ----
function seedColombiaFields() {
  const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
  let sheet = ss.getSheetByName(FIELDS_SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(FIELDS_SHEET_NAME);
    const headers = ['field_id', 'country_code', 'campaign_type', 'placeholder', 'label_es', 'field_type', 'icon', 'required', 'section', 'validation_rule', 'options', 'default_value', 'tooltip', 'depends_on', 'order', 'group'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  
  // Si ya tiene datos, no duplicar
  if (sheet.getLastRow() > 1) {
    Logger.log('⚠️ Template_Fields ya tiene datos. Saltando seed.');
    return;
  }
  
  const fields = [
    ['userEmail', 'ALL', 'ALL', '', 'Email corporativo', 'email', 'fa-envelope', 'TRUE', '1', '', '', '', '', '', '1', 'Info Básica'],
    ['dynamicType', 'ALL', 'ALL', '', 'Tipo de campaña', 'select', 'fa-layer-group', 'TRUE', '1', '', 'Cashback|Concurso Mayor Comprador', '', '', '', '2', 'Info Básica'],
    ['campaignName', 'ALL', 'ALL', '{{NOMBRE_CAMPANA}}', 'Nombre campaña', 'text', 'fa-tag', 'FALSE', '1', '', '', '', 'Se genera automático si vacío', '', '3', 'Info Básica'],
    ['shopName', 'ALL', 'ALL', '{{TIENDA_BASE}}', 'Nombre tienda', 'text', 'fa-shop', 'TRUE', '2', '', '', '', '', '', '1', 'Aliado'],
    ['territory', 'ALL', 'ALL', '{{TEXTO_TERRITORIO}}', 'Territorio', 'text', 'fa-map-location-dot', 'TRUE', '2', '', '', '', '', '', '2', 'Aliado'],
    ['startDate', 'ALL', 'ALL', '{{FECHA_INICIO}}', 'Fecha inicio', 'date', 'fa-calendar', 'TRUE', '2', '', '', '', '', '', '3', 'Vigencia'],
    ['endDate', 'ALL', 'ALL', '{{FECHA_FIN}}', 'Fecha fin', 'date', 'fa-calendar-check', 'TRUE', '2', '', '', '', '', '', '4', 'Vigencia'],
    ['cashbackPct', 'ALL', 'Cashback', '{{TEXTO_PORCENTAJE}}', '% Cashback', 'number', 'fa-percent', 'TRUE', '3', 'min:1,max:100', '', '', '', '', '1', 'Cashback'],
    ['cap', 'ALL', 'Cashback', '{{TOPE_NUM}}', 'Tope máximo', 'number', 'fa-ban', 'TRUE', '3', '', '', '', '', '', '2', 'Cashback'],
    ['budget', 'ALL', 'Cashback', '{{PRESUPUESTO_NUM}}', 'Presupuesto', 'number', 'fa-sack-dollar', 'TRUE', '3', '', '', '', '', '', '3', 'Cashback'],
    ['redemptionPlace', 'ALL', 'Cashback', '{{TEXTO_LUGAR_REDENCION}}', 'Lugar redención', 'select', 'fa-location-crosshairs', 'FALSE', '3', '', 'Brand Credits|Restaurantes|Créditos Generales', '', '', '', '4', 'Cashback'],
    ['organizerLegalName', 'ALL', 'Concurso Mayor Comprador', '{{ORGANIZADOR}}', 'Razón social', 'text', 'fa-briefcase', 'TRUE', '3', '', '', '', '', '', '1', 'Organizador'],
    ['organizerPhone', 'ALL', 'Concurso Mayor Comprador', '{{TELEFONO_CONTACTO}}', 'Teléfono', 'text', 'fa-phone', 'TRUE', '3', '', '', '', '', '', '2', 'Organizador'],
    ['organizerEmail', 'ALL', 'Concurso Mayor Comprador', '{{EMAIL_CONTACTO}}', 'Email PQR', 'email', 'fa-envelope', 'TRUE', '3', '', '', '', '', '', '3', 'Organizador'],
    ['numberOfWinners', 'ALL', 'Concurso Mayor Comprador', '{{NUM_GANADORES}}', '# Ganadores', 'number', 'fa-users', 'TRUE', '3', '', '', '1', '', '', '4', 'Mecánica'],
    ['winnerCriteria', 'ALL', 'Concurso Mayor Comprador', '{{CRITERIO_GANADOR}}', 'Criterio ganador', 'select', 'fa-filter', 'TRUE', '3', '', 'Mayor Venta ($)|Más Órdenes (#)', '', '', '', '5', 'Mecánica'],
    ['announcementDate', 'ALL', 'Concurso Mayor Comprador', '{{FECHA_ANUNCIO}}', 'Fecha anuncio', 'date', 'fa-bullhorn', 'TRUE', '3', '', '', '', '', '', '6', 'Mecánica'],
    ['paymentMethods', 'ALL', 'ALL', '{{TEXTO_METODO_PAGO}}', 'Métodos de pago', 'select', 'fa-credit-card', 'FALSE', '4', '', 'Todos excepto Efectivo|Todos|Únicamente RappiCard', 'Todos excepto Efectivo', '', '', '1', 'Restricciones'],
    ['userSegment', 'ALL', 'ALL', '{{TEXTO_SEGMENTO}}', 'Segmento usuarios', 'select', 'fa-users', 'FALSE', '4', '', 'Todos los usuarios|Pro y Pro Black|Nuevos Usuarios|Reactivos', 'Todos los usuarios', '', '', '2', 'Restricciones'],
    ['maxOrders', 'ALL', 'ALL', '{{LIMITE_ORDENES}}', 'Máx órdenes', 'text', 'fa-list-ol', 'FALSE', '4', '', '', '1', '', '', '3', 'Restricciones'],
    ['specialConditions', 'ALL', 'ALL', '{{CONDICIONES_ESPECIALES}}', 'Condiciones especiales', 'textarea', 'fa-exclamation-circle', 'FALSE', '4', '', '', '', '', '', '5', 'Restricciones']
  ];
  
  fields.forEach(f => sheet.appendRow(f));
  Logger.log('✅ ' + fields.length + ' campos creados en Template_Fields');
}

// =================================================================
// 6. ADMIN PANEL — BACKEND CRUD
// =================================================================

function _isAdmin() {
  try {
    const email = Session.getActiveUser().getEmail();
    return ADMIN_EMAILS_LIST.includes(email);
  } catch (e) { return true; } // permisivo en dev
}

function adminGetTemplates() {
  try {
    const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    let sheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'ok', templates: [] });
    return JSON.stringify({ status: 'ok', templates: _getSheetAsObjects(sheet) });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminSaveTemplate(templateDataJson, editIndex) {
  try {
    const data = JSON.parse(templateDataJson);
    const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    let sheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(REGISTRY_SHEET_NAME);
      const headers = ['country_code', 'country_name', 'campaign_type', 'template_doc_id', 'version', 'status', 'currency_code', 'currency_symbol', 'legal_owner', 'last_updated', 'notes'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    
    // Validar Doc ID
    try { DriveApp.getFileById(data.template_doc_id); } 
    catch (docErr) { return JSON.stringify({ status: 'error', message: 'No se pudo acceder al Google Doc. Verifica el ID y permisos.' }); }
    
    const row = [data.country_code, data.country_name, data.campaign_type, data.template_doc_id, data.version || '1.0', data.status || 'active', data.currency_code || 'COP', data.currency_symbol || '$', data.legal_owner || '', data.last_updated || new Date().toISOString().split('T')[0], data.notes || ''];
    
    if (editIndex >= 0) {
      sheet.getRange(editIndex + 2, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    return JSON.stringify({ status: 'ok' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminToggleTemplate(index, newStatus) {
  try {
    const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    const sheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Hoja no encontrada' });
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol = headers.indexOf('status') + 1;
    const updatedCol = headers.indexOf('last_updated') + 1;
    sheet.getRange(index + 2, statusCol).setValue(newStatus);
    if (updatedCol > 0) sheet.getRange(index + 2, updatedCol).setValue(new Date().toISOString().split('T')[0]);
    return JSON.stringify({ status: 'ok' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminDeleteTemplate(index) {
  try {
    const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    const sheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Hoja no encontrada' });
    sheet.deleteRow(index + 2);
    return JSON.stringify({ status: 'ok' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminGetFields() {
  try {
    const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    let sheet = ss.getSheetByName(FIELDS_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'ok', fields: [] });
    return JSON.stringify({ status: 'ok', fields: _getSheetAsObjects(sheet) });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminSaveField(fieldDataJson, editIndex) {
  try {
    const data = JSON.parse(fieldDataJson);
    const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    let sheet = ss.getSheetByName(FIELDS_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(FIELDS_SHEET_NAME);
      const headers = ['field_id', 'country_code', 'campaign_type', 'placeholder', 'label_es', 'field_type', 'icon', 'required', 'section', 'validation_rule', 'options', 'default_value', 'tooltip', 'depends_on', 'order', 'group'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    const row = [data.field_id, data.country_code || 'ALL', data.campaign_type || 'ALL', data.placeholder || '', data.label_es || '', data.field_type || 'text', data.icon || '', data.required || 'FALSE', data.section || '3', data.validation_rule || '', data.options || '', data.default_value || '', data.tooltip || '', data.depends_on || '', data.order || '1', data.group || ''];
    if (editIndex >= 0) { sheet.getRange(editIndex + 2, 1, 1, row.length).setValues([row]); }
    else { sheet.appendRow(row); }
    return JSON.stringify({ status: 'ok' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminDeleteField(index) {
  try {
    const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    const sheet = ss.getSheetByName(FIELDS_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Hoja no encontrada' });
    sheet.deleteRow(index + 2);
    return JSON.stringify({ status: 'ok' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminGetLogs() {
  try {
    const ss = SpreadsheetApp.openById(AUDIT_SHEET_ID);
    const sheet = ss.getSheetByName('Respuestas_Audit_V2');
    if (!sheet) return JSON.stringify({ status: 'ok', logs: [] });
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ status: 'ok', logs: [] });
    const logs = [];
    const startRow = Math.max(1, data.length - 50);
    for (let i = data.length - 1; i >= startRow; i--) {
      const row = data[i];
      logs.push({
        timestamp: row[0] ? new Date(row[0]).toLocaleString('es-CO') : '-',
        email: row[1] || '-', docUrl: row[2] || '', type: row[3] || '-',
        country: row[4] || 'CO', shop: row[5] || '-'
      });
    }
    return JSON.stringify({ status: 'ok', logs: logs });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}
// =================================================================
// 🔐 RAPPIMIND V2.3 — AUTH + CARPETAS + APROBACIÓN
// =================================================================
// INSTRUCCIONES:
//   Pegar al FINAL de Código.gs (después del bloque V2.2)
//   Ejecutar setupAdminSystem() UNA sola vez
// =================================================================

// -----------------------------------------------------------------
// CONSTANTES AUTH
// -----------------------------------------------------------------
// Fix: asegurar que _sheetToObjects existe
if (typeof _sheetToObjects === 'undefined') {
  // no-op, ya definida abajo
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
const TEAM_SHEET_NAME     = 'Admin_Team';
const APPROVAL_LOG_SHEET  = 'Approval_Log';
const TEMPLATES_ROOT_NAME = 'RappiMind_Templates';

// Roles: owner > admin > editor > viewer
const ROLE_HIERARCHY = { owner: 4, admin: 3, editor: 2, viewer: 1 };

// =================================================================
// 1. AUTENTICACIÓN — Obtener usuario actual y su rol
// =================================================================
function adminGetCurrentUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      return JSON.stringify({
        status: 'error',
        message: 'No se pudo obtener el email. ¿Estás logueado con cuenta Rappi?'
      });
    }

    const team = _getTeamMembers();
    const member = team.find(m => m.email.toLowerCase() === email.toLowerCase());

    if (!member) {
      return JSON.stringify({
        status: 'unauthorized',
        email: email,
        message: 'No tienes acceso al Panel de Administración. Contacta a un Owner para solicitar acceso.'
      });
    }

    return JSON.stringify({
      status: 'ok',
      email: email,
      role: member.role,
      name: member.name,
      permissions: _getPermissions(member.role)
    });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function _getPermissions(role) {
  const perms = {
    canViewAdmin:       ROLE_HIERARCHY[role] >= 1,  // viewer+
    canEditFields:      ROLE_HIERARCHY[role] >= 2,  // editor+
    canCreateTemplates: ROLE_HIERARCHY[role] >= 2,  // editor+
    canApproveTemplates:ROLE_HIERARCHY[role] >= 3,  // admin+
    canActivateTemplates:ROLE_HIERARCHY[role] >= 3, // admin+
    canDeleteTemplates: ROLE_HIERARCHY[role] >= 3,  // admin+
    canManageTeam:      ROLE_HIERARCHY[role] >= 4,  // owner only
    canManageFolders:   ROLE_HIERARCHY[role] >= 3,  // admin+
  };
  return perms;
}

// =================================================================
// 2. EQUIPO — CRUD de miembros
// =================================================================
function adminGetTeam() {
  try {
    _requireRole('viewer');
    const team = _getTeamMembers();
    return JSON.stringify({ status: 'ok', team: team });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminAddTeamMember(jsonStr) {
  try {
    _requireRole('owner');
    const d = JSON.parse(jsonStr);

    if (!d.email || !d.role || !d.name) {
      return JSON.stringify({ status: 'error', message: 'Email, nombre y rol son requeridos' });
    }

    // Validar email @rappi.com
    if (!d.email.toLowerCase().endsWith('@rappi.com')) {
      return JSON.stringify({ status: 'error', message: 'Solo se permiten emails @rappi.com' });
    }

    // Validar rol válido
    if (!ROLE_HIERARCHY[d.role]) {
      return JSON.stringify({ status: 'error', message: 'Rol inválido: ' + d.role });
    }

    const sheet = _getOrCreateSheet(TEAM_SHEET_NAME,
      ['email', 'name', 'role', 'added_by', 'added_date', 'status', 'notes']);

    // Verificar duplicado
    const existing = _getTeamMembers();
    if (existing.find(m => m.email.toLowerCase() === d.email.toLowerCase())) {
      return JSON.stringify({ status: 'error', message: 'Este email ya está en el equipo' });
    }

    const adderEmail = Session.getActiveUser().getEmail();
    sheet.appendRow([
      d.email.toLowerCase(),
      d.name,
      d.role,
      adderEmail,
      new Date().toISOString().split('T')[0],
      'active',
      d.notes || ''
    ]);

    // Compartir carpeta de templates con el nuevo miembro
    _shareFolderWithMember(d.email, d.role);

    _logApprovalAction('team_add', `Agregó a ${d.name} (${d.email}) como ${d.role}`);

    return JSON.stringify({ status: 'ok', message: `${d.name} agregado como ${d.role}` });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminUpdateTeamMember(jsonStr) {
  try {
    _requireRole('owner');
    const d = JSON.parse(jsonStr);
    const sheet = _getSheet(TEAM_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Sheet de equipo no existe' });

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('email');
    const roleCol  = headers.indexOf('role');
    const statusCol = headers.indexOf('status');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][emailCol]).toLowerCase() === d.email.toLowerCase()) {
        if (d.role) sheet.getRange(i + 1, roleCol + 1).setValue(d.role);
        if (d.status) sheet.getRange(i + 1, statusCol + 1).setValue(d.status);
        _logApprovalAction('team_update', `Actualizó rol de ${d.email} a ${d.role || d.status}`);
        return JSON.stringify({ status: 'ok' });
      }
    }
    return JSON.stringify({ status: 'error', message: 'Miembro no encontrado' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminRemoveTeamMember(email) {
  try {
    _requireRole('owner');
    const callerEmail = Session.getActiveUser().getEmail();

    if (email.toLowerCase() === callerEmail.toLowerCase()) {
      return JSON.stringify({ status: 'error', message: 'No puedes eliminarte a ti mismo' });
    }

    const sheet = _getSheet(TEAM_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Sheet no existe' });

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === email.toLowerCase()) {
        sheet.deleteRow(i + 1);
        _logApprovalAction('team_remove', `Eliminó a ${email} del equipo`);
        return JSON.stringify({ status: 'ok' });
      }
    }
    return JSON.stringify({ status: 'error', message: 'No encontrado' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function _getTeamMembers() {
  const sheet = _getSheet(TEAM_SHEET_NAME);
  if (!sheet) return [];
  return _sheetToObjects(sheet).filter(m => m.status === 'active');
}

// =================================================================
// 3. CARPETAS ESTRUCTURADAS EN DRIVE
// =================================================================
function setupTemplateFolders() {
  Logger.log('📁 Creando estructura de carpetas...');

  // Crear carpeta raíz
  const root = _getOrCreateDriveFolder(null, TEMPLATES_ROOT_NAME);
  Logger.log('📁 Raíz: ' + root.getName() + ' (' + root.getId() + ')');

  const countries = [
    { code: 'CO', name: 'Colombia' },
    { code: 'MX', name: 'México' },
    { code: 'CR', name: 'Costa Rica' },
    { code: 'BR', name: 'Brasil' },
    { code: 'AR', name: 'Argentina' },
    { code: 'UY', name: 'Uruguay' },
    { code: 'CL', name: 'Chile' },
    { code: 'EC', name: 'Ecuador' }
  ];

  const types = ['Cashback', 'Concurso'];

  countries.forEach(country => {
    const countryFolder = _getOrCreateDriveFolder(root, `${country.code}_${country.name}`);
    types.forEach(type => {
      _getOrCreateDriveFolder(countryFolder, type);
    });
  });

  // Carpetas especiales
  _getOrCreateDriveFolder(root, '_Borradores');
  _getOrCreateDriveFolder(root, '_Archivo');

  Logger.log('✅ Estructura de carpetas creada');
  Logger.log('📋 Folder ID raíz: ' + root.getId());

  // Guardar ID de carpeta raíz en propiedades del script
  PropertiesService.getScriptProperties().setProperty('TEMPLATES_FOLDER_ID', root.getId());

  return root.getId();
}

function _getOrCreateDriveFolder(parent, name) {
  let iter;
  if (parent) {
    iter = parent.getFoldersByName(name);
  } else {
    iter = DriveApp.getFoldersByName(name);
  }

  if (iter.hasNext()) return iter.next();

  if (parent) {
    return parent.createFolder(name);
  } else {
    return DriveApp.createFolder(name);
  }
}

function _shareFolderWithMember(email, role) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty('TEMPLATES_FOLDER_ID');
    if (!folderId) return;

    const folder = DriveApp.getFolderById(folderId);

    if (role === 'viewer') {
      folder.addViewer(email);
    } else {
      folder.addEditor(email);
    }
    Logger.log('📁 Carpeta compartida con ' + email + ' como ' + role);
  } catch (e) {
    Logger.log('⚠️ No se pudo compartir carpeta con ' + email + ': ' + e.message);
  }
}

// Mover template a la carpeta correcta por país/tipo
function _moveTemplateToFolder(docId, countryCode, countryName, campaignType) {
  try {
    const rootId = PropertiesService.getScriptProperties().getProperty('TEMPLATES_FOLDER_ID');
    if (!rootId) return;

    const root = DriveApp.getFolderById(rootId);
    const countryFolderName = `${countryCode}_${countryName}`;
    const countryFolder = _getOrCreateDriveFolder(root, countryFolderName);

    // Determinar subcarpeta
    let subFolderName = campaignType;
    if (campaignType.includes('Concurso')) subFolderName = 'Concurso';
    const typeFolder = _getOrCreateDriveFolder(countryFolder, subFolderName);

    const file = DriveApp.getFileById(docId);
    typeFolder.addFile(file);

    // Remover de la ubicación anterior (raíz de Drive)
    const parents = file.getParents();
    while (parents.hasNext()) {
      const p = parents.next();
      if (p.getId() !== typeFolder.getId()) {
        p.removeFile(file);
      }
    }

    Logger.log('📁 Template movido a: ' + countryFolderName + '/' + subFolderName);
  } catch (e) {
    Logger.log('⚠️ No se pudo mover template: ' + e.message);
  }
}

function adminGetFolderStructure() {
  try {
    _requireRole('viewer');
    const rootId = PropertiesService.getScriptProperties().getProperty('TEMPLATES_FOLDER_ID');
    if (!rootId) {
      return JSON.stringify({ status: 'ok', folders: [], rootUrl: null });
    }

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

    return JSON.stringify({
      status: 'ok',
      rootId: rootId,
      rootUrl: root.getUrl(),
      folders: folders
    });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

// =================================================================
// 4. WORKFLOW DE APROBACIÓN
// =================================================================

/**
 * Estados de un template:
 *   draft           → Editor creó pero no envió
 *   pending_review  → Editor envió para revisión
 *   approved        → Admin aprobó (pero aún no activo)
 *   active          → Admin activó (en producción)
 *   rejected        → Admin rechazó (con comentarios)
 *   inactive        → Desactivado manualmente
 */

function adminSubmitForReview(index) {
  try {
    _requireRole('editor');
    const sheet = _getSheet(REGISTRY_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Registry no existe' });

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol    = headers.indexOf('status') + 1;
    const updCol       = headers.indexOf('last_updated') + 1;
    const submittedCol = headers.indexOf('submitted_by');
    const submittedDateCol = headers.indexOf('submitted_date');

    const currentStatus = sheet.getRange(index + 2, statusCol).getValue();
    if (currentStatus !== 'draft' && currentStatus !== 'rejected') {
      return JSON.stringify({ status: 'error', message: 'Solo templates en borrador o rechazados pueden enviarse a revisión' });
    }

    const email = Session.getActiveUser().getEmail();
    sheet.getRange(index + 2, statusCol).setValue('pending_review');
    if (updCol > 0) sheet.getRange(index + 2, updCol).setValue(new Date().toISOString().split('T')[0]);
    if (submittedCol >= 0) sheet.getRange(index + 2, submittedCol + 1).setValue(email);
    if (submittedDateCol >= 0) sheet.getRange(index + 2, submittedDateCol + 1).setValue(new Date().toISOString());

    _logApprovalAction('submit_review', `Template #${index} enviado a revisión por ${email}`);
    _notifyAdmins(index, 'pending_review', email);

    return JSON.stringify({ status: 'ok', message: 'Template enviado a revisión' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminApproveTemplate(index, activateNow) {
  try {
    _requireRole('admin');
    const sheet = _getSheet(REGISTRY_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Registry no existe' });

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol    = headers.indexOf('status') + 1;
    const updCol       = headers.indexOf('last_updated') + 1;
    const approvedByCol = headers.indexOf('approved_by');
    const approvedDateCol = headers.indexOf('approved_date');

    const currentStatus = sheet.getRange(index + 2, statusCol).getValue();
    if (currentStatus !== 'pending_review') {
      return JSON.stringify({ status: 'error', message: 'Solo se pueden aprobar templates en revisión' });
    }

    const email = Session.getActiveUser().getEmail();
    const newStatus = activateNow ? 'active' : 'approved';

    sheet.getRange(index + 2, statusCol).setValue(newStatus);
    if (updCol > 0) sheet.getRange(index + 2, updCol).setValue(new Date().toISOString().split('T')[0]);
    if (approvedByCol >= 0) sheet.getRange(index + 2, approvedByCol + 1).setValue(email);
    if (approvedDateCol >= 0) sheet.getRange(index + 2, approvedDateCol + 1).setValue(new Date().toISOString());

    _logApprovalAction('approve', `Template #${index} aprobado por ${email}` + (activateNow ? ' y activado' : ''));

    return JSON.stringify({ status: 'ok', message: activateNow ? 'Template aprobado y activado' : 'Template aprobado' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminRejectTemplate(index, reason) {
  try {
    _requireRole('admin');
    const sheet = _getSheet(REGISTRY_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Registry no existe' });

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const statusCol = headers.indexOf('status') + 1;
    const updCol    = headers.indexOf('last_updated') + 1;
    const notesCol  = headers.indexOf('notes') + 1;
    const rejectedByCol = headers.indexOf('rejected_by');

    const email = Session.getActiveUser().getEmail();
    sheet.getRange(index + 2, statusCol).setValue('rejected');
    if (updCol > 0) sheet.getRange(index + 2, updCol).setValue(new Date().toISOString().split('T')[0]);
    if (notesCol > 0) sheet.getRange(index + 2, notesCol).setValue('RECHAZADO: ' + (reason || 'Sin motivo'));
    if (rejectedByCol >= 0) sheet.getRange(index + 2, rejectedByCol + 1).setValue(email);

    _logApprovalAction('reject', `Template #${index} rechazado por ${email}: ${reason}`);

    return JSON.stringify({ status: 'ok', message: 'Template rechazado' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

function adminGetApprovalLog() {
  try {
    _requireRole('viewer');
    const sheet = _getSheet(APPROVAL_LOG_SHEET);
    if (!sheet) return JSON.stringify({ status: 'ok', logs: [] });
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ status: 'ok', logs: [] });

    const logs = [];
    const start = Math.max(1, data.length - 100);
    for (let i = data.length - 1; i >= start; i--) {
      logs.push({
        timestamp: data[i][0] ? new Date(data[i][0]).toLocaleString('es-CO') : '-',
        actor: data[i][1] || '-',
        action: data[i][2] || '-',
        details: data[i][3] || ''
      });
    }
    return JSON.stringify({ status: 'ok', logs: logs });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}

// =================================================================
// 5. HELPERS AUTH
// =================================================================
function _requireRole(minRole) {
  const email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('No autenticado');

  const team = _getTeamMembers();
  const member = team.find(m => m.email.toLowerCase() === email.toLowerCase());

  if (!member) throw new Error('Sin acceso: ' + email);
  if (ROLE_HIERARCHY[member.role] < ROLE_HIERARCHY[minRole]) {
    throw new Error(`Permiso insuficiente. Requiere: ${minRole}, tiene: ${member.role}`);
  }
  return member;
}

function _getSheet(name) {
  const SHEET_ID = '1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheetByName(name);
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

function _logApprovalAction(action, details) {
  try {
    const sheet = _getOrCreateSheet(APPROVAL_LOG_SHEET,
      ['timestamp', 'actor', 'action', 'details']);
    const email = Session.getActiveUser().getEmail();
    sheet.appendRow([new Date(), email, action, details]);
  } catch (e) {
    Logger.log('Log error: ' + e.message);
  }
}

function _notifyAdmins(templateIndex, status, submitterEmail) {
  try {
    const team = _getTeamMembers();
    const admins = team.filter(m => ROLE_HIERARCHY[m.role] >= ROLE_HIERARCHY['admin']);

    const registry = _getTemplateRegistry();
    const tpl = registry[templateIndex];
    if (!tpl) return;

    const subject = `🧠 RappiMind: Template pendiente de aprobación`;
    const htmlBody = `
      <div style="font-family:sans-serif;max-width:600px;margin:0 auto;">
        <h2 style="color:#FF5C00;">🧠 RappiMind — Solicitud de Aprobación</h2>
        <p><strong>${submitterEmail}</strong> envió un template para revisión:</p>
        <table style="border-collapse:collapse;width:100%;margin:16px 0;">
          <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">País</td><td style="padding:8px;border:1px solid #ddd;">${tpl.country_name || tpl.country_code}</td></tr>
          <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">Tipo</td><td style="padding:8px;border:1px solid #ddd;">${tpl.campaign_type}</td></tr>
          <tr><td style="padding:8px;border:1px solid #ddd;font-weight:bold;">Template</td><td style="padding:8px;border:1px solid #ddd;"><a href="https://docs.google.com/document/d/${tpl.template_doc_id}/edit">Abrir Google Doc</a></td></tr>
        </table>
        <p>Ingresa al <strong>Panel de Administración de RappiMind</strong> para aprobar o rechazar.</p>
      </div>
    `;

    admins.forEach(admin => {
      if (admin.email.toLowerCase() !== submitterEmail.toLowerCase()) {
        MailApp.sendEmail({
          to: admin.email,
          subject: subject,
          htmlBody: htmlBody
        });
      }
    });
  } catch (e) {
    Logger.log('Notify error: ' + e.message);
  }
}

// =================================================================
// 6. SETUP ADMIN SYSTEM — EJECUTAR UNA VEZ
// =================================================================
function setupAdminSystem() {
  Logger.log('🔐 ═══════════════════════════════════════');
  Logger.log('🔐 RAPPIMIND ADMIN SYSTEM — SETUP');
  Logger.log('🔐 ═══════════════════════════════════════');

  // ──── PASO 1: Crear sheet de equipo ────
  Logger.log('👥 Creando sheet de equipo...');
  const teamSheet = _getOrCreateSheet(TEAM_SHEET_NAME,
    ['email', 'name', 'role', 'added_by', 'added_date', 'status', 'notes']);

  // Seed: Juan como owner
  const callerEmail = Session.getActiveUser().getEmail();
  const existingTeam = _sheetToObjects(teamSheet);
  if (existingTeam.length === 0) {
    teamSheet.appendRow([callerEmail, 'Juan (Owner)', 'owner', 'system', new Date().toISOString().split('T')[0], 'active', 'Setup inicial']);
    Logger.log('✅ ' + callerEmail + ' agregado como owner');
  }

  // ──── PASO 2: Agregar columnas de aprobación al Registry ────
  Logger.log('📋 Actualizando Registry con columnas de aprobación...');
  _upgradeRegistryColumns();

  // ──── PASO 3: Crear approval log ────
  Logger.log('📝 Creando log de aprobaciones...');
  _getOrCreateSheet(APPROVAL_LOG_SHEET,
    ['timestamp', 'actor', 'action', 'details']);

  // ──── PASO 4: Crear carpetas ────
  Logger.log('📁 Creando estructura de carpetas...');
  const rootId = setupTemplateFolders();

  // ──── PASO 5: Mover templates existentes a carpetas ────
  Logger.log('📦 Moviendo templates existentes...');
  _moveExistingTemplatesToFolders();

  Logger.log('');
  Logger.log('🎉 ═══════════════════════════════════════');
  Logger.log('🎉 ADMIN SYSTEM CONFIGURADO');
  Logger.log('🎉 ═══════════════════════════════════════');
  Logger.log('');
  Logger.log('👤 Owner: ' + callerEmail);
  Logger.log('📁 Carpeta raíz: ' + rootId);
  Logger.log('');
  Logger.log('👉 SIGUIENTE: Agrega miembros al equipo desde el Admin Panel');
  Logger.log('   o manualmente en la hoja Admin_Team');
}

function _upgradeRegistryColumns() {
  const sheet = _getSheet(REGISTRY_SHEET_NAME);
  if (!sheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newCols = ['submitted_by', 'submitted_date', 'approved_by', 'approved_date', 'rejected_by'];

  newCols.forEach(col => {
    if (headers.indexOf(col) === -1) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(col)
        .setBackground('#1F2937').setFontColor('#FFFFFF').setFontWeight('bold');
      Logger.log('  + Columna: ' + col);
    }
  });
}

function _moveExistingTemplatesToFolders() {
  try {
    const registry = _getTemplateRegistry();
    registry.forEach(tpl => {
      if (tpl.template_doc_id) {
        _moveTemplateToFolder(
          tpl.template_doc_id,
          tpl.country_code,
          tpl.country_name || tpl.country_code,
          tpl.campaign_type
        );
      }
    });
  } catch (e) {
    Logger.log('⚠️ Error moviendo templates: ' + e.message);
  }
}

// =================================================================
// 7. OVERRIDE adminSaveTemplate — Versión con aprobación y carpetas
// =================================================================
// NOTA: Esta función REEMPLAZA la de V2.2
function adminSaveTemplate(jsonStr, editIndex) {
  try {
    const callerRole = _requireRole('editor');
    const d = JSON.parse(jsonStr);

    const SHEET_ID = '1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI';
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(REGISTRY_SHEET_NAME);
    if (!sheet) return JSON.stringify({ status: 'error', message: 'Registry no existe. Ejecuta setupTemplateEngine primero.' });

    // Validar Doc ID
    try { DriveApp.getFileById(d.template_doc_id); }
    catch (e) { return JSON.stringify({ status: 'error', message: 'No se pudo acceder al Google Doc: ' + d.template_doc_id }); }

    const callerEmail = Session.getActiveUser().getEmail();

    // Determinar estado inicial según rol
    let initialStatus;
    if (ROLE_HIERARCHY[callerRole.role] >= ROLE_HIERARCHY['admin']) {
      // Admin+ puede elegir estado
      initialStatus = d.status || 'active';
    } else {
      // Editor siempre crea como draft
      initialStatus = 'draft';
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = new Array(headers.length).fill('');

    const setValue = (colName, value) => {
      const idx = headers.indexOf(colName);
      if (idx >= 0) row[idx] = value || '';
    };

    setValue('country_code', d.country_code);
    setValue('country_name', d.country_name);
    setValue('campaign_type', d.campaign_type);
    setValue('template_doc_id', d.template_doc_id);
    setValue('version', d.version || '1.0');
    setValue('status', initialStatus);
    setValue('currency_code', d.currency_code || 'COP');
    setValue('currency_symbol', d.currency_symbol || '$');
    setValue('legal_owner', d.legal_owner || callerEmail);
    setValue('last_updated', new Date().toISOString().split('T')[0]);
    setValue('notes', d.notes || '');
    setValue('submitted_by', callerEmail);

    if (editIndex >= 0) {
      sheet.getRange(editIndex + 2, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    // Mover a carpeta estructurada
    _moveTemplateToFolder(d.template_doc_id, d.country_code, d.country_name, d.campaign_type);

    const action = editIndex >= 0 ? 'template_edit' : 'template_create';
    _logApprovalAction(action, `${d.country_code}/${d.campaign_type} por ${callerEmail} [${initialStatus}]`);

    const msg = initialStatus === 'draft'
      ? 'Template guardado como borrador. Envíalo a revisión cuando esté listo.'
      : 'Template guardado como ' + initialStatus;

    return JSON.stringify({ status: 'ok', message: msg, initialStatus: initialStatus });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.message });
  }
}
// ============================================================
// RAPPIMIND — TEMPLATE WIZARD BACKEND
// Archivo: TemplateWizard_Backend.gs
// 
// INSTRUCCIONES DE INTEGRACIÓN:
// Copia TODO este contenido y pégalo al FINAL de tu Código.gs
// No modifica ninguna función existente.
// ============================================================

// ============================================================
// CONFIGURACIÓN
// ============================================================

const TW_CONFIG = {
  SHEET_REGISTRY: 'Template_Registry',
  SHEET_FIELDS: 'Template_Fields',
  SHEET_AUDIT: 'Audit_Log',
  SHEET_TEAM: 'Admin_Team',
  DRIVE_ROOT_FOLDER_NAME: 'RappiMind_Templates',
  ADMIN_EMAILS: ['juan.gallego@rappi.com', 'david.gaviria@rappi.com'],
  
  // Mapeo de países para carpetas
  COUNTRY_FOLDERS: {
    'CO': 'CO_Colombia',
    'MX': 'MX_México',
    'PE': 'PE_Perú',
    'AR': 'AR_Argentina',
    'CL': 'CL_Chile',
    'EC': 'EC_Ecuador',
    'UY': 'UY_Uruguay',
    'CR': 'CR_Costa_Rica',
    'BR': 'BR_Brasil'
  },

  // Campos conocidos para sugerencia automática de placeholders
  KNOWN_PLACEHOLDERS: {
    // Fechas
    'FECHA_INICIO': { label: 'Fecha de inicio', type: 'date', pattern: /\d{1,2} de \w+ de \d{4}/ },
    'FECHA_FIN': { label: 'Fecha de fin', type: 'date', pattern: /\d{1,2} de \w+ de \d{4}/ },
    'FECHA_PUBLICACION': { label: 'Fecha de publicación', type: 'date' },
    // Campaña
    'NOMBRE_CAMPANA': { label: 'Nombre de campaña', type: 'text' },
    'TERRITORIO': { label: 'Territorio', type: 'text' },
    'ORGANIZADOR': { label: 'Organizador legal', type: 'text' },
    'NIT_ORGANIZADOR': { label: 'NIT del organizador', type: 'text' },
    // Cashback
    'PORCENTAJE_CASHBACK': { label: '% de cashback', type: 'number' },
    'TOPE_NUM': { label: 'Tope máximo (número)', type: 'number' },
    'TOPE_TEXTO': { label: 'Tope máximo (escrito)', type: 'text' },
    'MONEDA': { label: 'Moneda', type: 'text' },
    'MONTO_MINIMO': { label: 'Monto mínimo de compra', type: 'number' },
    // Concurso
    'PREMIO_DESCRIPCION': { label: 'Descripción del premio', type: 'text' },
    'PREMIO_VALOR': { label: 'Valor del premio', type: 'number' },
    'NUMERO_GANADORES': { label: 'Número de ganadores', type: 'number' },
    'CRITERIO_SELECCION': { label: 'Criterio de selección', type: 'text' },
    'TOP_N': { label: 'Top N (ej: 10)', type: 'number' },
    // Legal
    'JURISDICCION': { label: 'Jurisdicción legal', type: 'text' },
    'LEY_APLICABLE': { label: 'Ley aplicable', type: 'text' },
    'ENTIDAD_VIGILANCIA': { label: 'Entidad de vigilancia', type: 'text' },
    // Otros
    'CATEGORIAS_PARTICIPANTES': { label: 'Categorías de participantes', type: 'text' },
    'VIGENCIA_DIAS': { label: 'Vigencia en días', type: 'number' },
    'URL_BASES': { label: 'URL de bases legales', type: 'url' },
  }
};

// ============================================================
// FUNCIONES EXPUESTAS A google.script.run
// El frontend las llama directamente: google.script.run.analyzeTextForPlaceholders(data)
// No se necesita doPost() ni ningún router.
// ============================================================

// ============================================================
// 1. ANALIZAR TEXTO CON GEMINI → SUGERIR PLACEHOLDERS
// ============================================================

function analyzeTextForPlaceholders(payload) {
  try {
    const { text, countryCode, campaignType, userEmail } = payload;
    
    if (!text || text.trim().length < 50) {
      return buildResponse(false, 'El texto es muy corto. Pega el T&C completo.');
    }

    // Construir prompt para Gemini
    const prompt = buildAnalysisPrompt(text, countryCode, campaignType);
    
    // Llamar a Gemini
    const geminiResult = callGeminiForAnalysis(prompt);
    
    // Registrar en audit log
    logAuditEvent('tw_analyze', userEmail, { countryCode, campaignType, textLength: text.length });
    
    return buildResponse(true, 'Análisis completado', geminiResult);
    
  } catch (e) {
    Logger.log('Error en analyzeTextForPlaceholders: ' + e.message);
    return buildResponse(false, 'Error al analizar el texto: ' + e.message);
  }
}

function buildAnalysisPrompt(text, countryCode, campaignType) {
  const knownKeys = Object.entries(TW_CONFIG.KNOWN_PLACEHOLDERS)
    .map(([key, val]) => `- {{${key}}}: ${val.label}`)
    .join('\n');

  return `Eres un asistente legal especializado en términos y condiciones de campañas de marketing en Latinoamérica.

Analiza el siguiente documento de T&C y detecta TODOS los valores que son variables (que cambiarían entre una campaña y otra).

País: ${countryCode || 'No especificado'}
Tipo de campaña: ${campaignType || 'No especificado'}

PLACEHOLDERS ESTÁNDAR YA DEFINIDOS EN EL SISTEMA (úsalos cuando correspondan):
${knownKeys}

INSTRUCCIONES:
1. Identifica cada valor variable en el texto
2. Para cada uno, sugiere el placeholder más apropiado del sistema (o crea uno nuevo en MAYUSCULAS_CON_GUIONES_BAJOS si no existe)
3. Clasifica cada detección con nivel de confianza: HIGH (valor claramente variable), MEDIUM (probablemente variable), LOW (podría ser fijo)
4. Si un valor parece siempre fijo (ej: "República de Colombia"), NO lo incluyas
5. Agrupa valores del mismo tipo (ej: varias fechas deben mapearse a sus respectivos placeholders)

RESPONDE ÚNICAMENTE con un JSON con esta estructura exacta (sin markdown, sin explicación adicional):
{
  "detections": [
    {
      "original_text": "texto exacto encontrado en el documento",
      "suggested_placeholder": "NOMBRE_DEL_PLACEHOLDER",
      "label": "Descripción amigable",
      "confidence": "HIGH|MEDIUM|LOW",
      "context": "fragmento de contexto donde aparece (máx 80 chars)",
      "occurrences": 2
    }
  ],
  "summary": {
    "total_variables": 10,
    "high_confidence": 7,
    "campaign_type_detected": "Cashback|Concurso|Unknown",
    "notes": "observación general sobre el documento"
  }
}

DOCUMENTO A ANALIZAR:
---
${text.substring(0, 15000)}
---`;
}

function callGeminiForAnalysis(prompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    // Fallback: análisis básico por regex sin Gemini
    return fallbackAnalysis(prompt);
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  
  const requestBody = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: 4096,
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Error al conectar con Gemini: ' + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  const rawText = result.candidates?.[0]?.content?.parts?.[0]?.text || '';
  
  // Limpiar posible markdown
  const cleanJson = rawText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
  
  try {
    return JSON.parse(cleanJson);
  } catch (e) {
    throw new Error('Gemini devolvió un formato inesperado. Intenta de nuevo.');
  }
}

function fallbackAnalysis(originalText) {
  // Análisis básico sin Gemini: busca patrones comunes
  const detections = [];
  
  // Patrones básicos para detectar sin IA
  const patterns = [
    { regex: /\d{1,2} de \w+ de \d{4}/g, placeholder: 'FECHA_INICIO', label: 'Fecha detectada', confidence: 'MEDIUM' },
    { regex: /\d+(\.\d+)?%/g, placeholder: 'PORCENTAJE_CASHBACK', label: 'Porcentaje detectado', confidence: 'MEDIUM' },
    { regex: /\$\s*[\d,\.]+/g, placeholder: 'TOPE_NUM', label: 'Monto detectado', confidence: 'MEDIUM' },
  ];

  patterns.forEach(p => {
    const matches = [...new Set(originalText.match(p.regex) || [])];
    matches.forEach(match => {
      detections.push({
        original_text: match,
        suggested_placeholder: p.placeholder,
        label: p.label,
        confidence: p.confidence,
        context: '',
        occurrences: 1
      });
    });
  });

  return {
    detections,
    summary: {
      total_variables: detections.length,
      high_confidence: 0,
      campaign_type_detected: 'Unknown',
      notes: 'Análisis básico (configura GEMINI_API_KEY para análisis inteligente)'
    }
  };
}

// ============================================================
// 2. OBTENER CONTENIDO DE GOOGLE DOC POR URL/ID
// ============================================================

function fetchGoogleDocContent(payload) {
  try {
    const { docUrl, userEmail } = payload;
    
    // Extraer ID del Doc desde la URL
    const docId = extractDocIdFromUrl(docUrl);
    if (!docId) {
      return buildResponse(false, 'URL de Google Doc inválida. Verifica que sea un enlace correcto.');
    }

    // Intentar abrir el documento
    let doc;
    try {
      doc = DocumentApp.openById(docId);
    } catch (e) {
      return buildResponse(false, 'No se puede acceder al documento. Verifica que esté compartido con la cuenta del script.');
    }

    const body = doc.getBody();
    const text = body.getText();
    
    if (text.trim().length < 50) {
      return buildResponse(false, 'El documento parece estar vacío o tiene muy poco contenido.');
    }

    logAuditEvent('tw_fetchDoc', userEmail, { docId, docTitle: doc.getName(), charCount: text.length });

    return buildResponse(true, 'Documento cargado exitosamente', {
      docId,
      docTitle: doc.getName(),
      text,
      charCount: text.length,
      wordCount: text.split(/\s+/).length
    });

  } catch (e) {
    Logger.log('Error en fetchGoogleDocContent: ' + e.message);
    return buildResponse(false, 'Error al cargar el documento: ' + e.message);
  }
}

function extractDocIdFromUrl(url) {
  if (!url) return null;
  // Formato: https://docs.google.com/document/d/ID/edit
  const match = url.match(/\/d\/([a-zA-Z0-9_-]{25,})/);
  return match ? match[1] : null;
}

// ============================================================
// 3. CREAR TEMPLATE DESDE EL WIZARD
// ============================================================

function createTemplateFromWizard(payload) {
  try {
    const { 
      originalText,  // Texto original del T&C
      mappings,      // Array de {original_text, placeholder, confirmed: true/false}
      metadata,      // {countryCode, campaignType, version, notes}
      userEmail
    } = payload;

    // Validaciones
    if (!originalText || !mappings || !metadata) {
      return buildResponse(false, 'Faltan datos requeridos para crear el template.');
    }

    const confirmedMappings = mappings.filter(m => m.confirmed);
    if (confirmedMappings.length === 0) {
      return buildResponse(false, 'Debes confirmar al menos un placeholder antes de crear el template.');
    }

    // Verificar rol del usuario
    const userRole = getUserRole(userEmail);
    if (!userRole || userRole === 'none') {
      return buildResponse(false, 'No tienes permisos para crear templates.');
    }

    // Aplicar placeholders al texto
    let templateText = originalText;
    const appliedMappings = [];
    
    confirmedMappings.forEach(mapping => {
      const escapedText = mapping.original_text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const regex = new RegExp(escapedText, 'g');
      const occurrencesBefore = (templateText.match(regex) || []).length;
      templateText = templateText.replace(regex, `{{${mapping.placeholder}}}`);
      appliedMappings.push({
        original: mapping.original_text,
        placeholder: `{{${mapping.placeholder}}}`,
        occurrences: occurrencesBefore
      });
    });

    // Crear Google Doc con el template
    const docResult = createTemplateGoogleDoc(templateText, metadata);
    
    // Determinar estado inicial según rol
    const initialStatus = (userRole === 'admin' || userRole === 'owner') ? 'active' : 'pending_review';
    
    // Registrar en Template_Registry
    registerTemplate({
      countryCode: metadata.countryCode,
      campaignType: metadata.campaignType,
      docId: docResult.docId,
      docUrl: docResult.docUrl,
      version: metadata.version || '1.0',
      status: initialStatus,
      notes: metadata.notes || '',
      createdBy: userEmail,
      placeholderCount: confirmedMappings.length
    });

    // Registrar campos en Template_Fields
    registerTemplateFields(confirmedMappings, metadata);

    // Notificar a admins si es pending_review
    if (initialStatus === 'pending_review') {
      notifyAdminsForReview({
        userEmail,
        countryCode: metadata.countryCode,
        campaignType: metadata.campaignType,
        docUrl: docResult.docUrl,
        version: metadata.version
      });
    }

    logAuditEvent('tw_create', userEmail, {
      countryCode: metadata.countryCode,
      campaignType: metadata.campaignType,
      docId: docResult.docId,
      status: initialStatus,
      placeholderCount: confirmedMappings.length
    });

    return buildResponse(true, 
      initialStatus === 'active' 
        ? 'Template creado y activado exitosamente.' 
        : 'Template creado. Pendiente de revisión por un Admin.',
      {
        docId: docResult.docId,
        docUrl: docResult.docUrl,
        docTitle: docResult.docTitle,
        status: initialStatus,
        appliedMappings,
        placeholderCount: confirmedMappings.length
      }
    );

  } catch (e) {
    Logger.log('Error en createTemplateFromWizard: ' + e.message);
    return buildResponse(false, 'Error al crear el template: ' + e.message);
  }
}

function createTemplateGoogleDoc(templateText, metadata) {
  const countryName = TW_CONFIG.COUNTRY_FOLDERS[metadata.countryCode] || metadata.countryCode;
  const docTitle = `[TEMPLATE] ${countryName} - ${metadata.campaignType} v${metadata.version || '1.0'}`;
  
  // Crear el documento
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();
  
  // Limpiar y agregar el contenido
  body.clear();
  
  // Agregar encabezado de metadata
  const header = body.appendParagraph(`RAPPIMIND TEMPLATE — ${metadata.countryCode} / ${metadata.campaignType} / v${metadata.version || '1.0'}`);
  header.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  header.setAttributes({
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#FF4500',
    [DocumentApp.Attribute.BOLD]: true
  });
  
  body.appendParagraph('---');
  
  // Agregar el texto con placeholders
  // Dividir por párrafos para preservar estructura
  const paragraphs = templateText.split('\n');
  paragraphs.forEach(para => {
    if (para.trim()) {
      body.appendParagraph(para);
    } else {
      body.appendParagraph('');
    }
  });
  
  doc.saveAndClose();
  
  const file = DriveApp.getFileById(doc.getId());
  
  // Mover a carpeta correcta
  const folder = getOrCreateTemplateFolder(metadata.countryCode, metadata.campaignType);
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  return {
    docId: doc.getId(),
    docUrl: `https://docs.google.com/document/d/${doc.getId()}/edit`,
    docTitle
  };
}

function getOrCreateTemplateFolder(countryCode, campaignType) {
  const rootFolderName = TW_CONFIG.DRIVE_ROOT_FOLDER_NAME;
  const countryFolderName = TW_CONFIG.COUNTRY_FOLDERS[countryCode] || countryCode;
  
  // Root folder
  let rootFolder;
  const rootSearch = DriveApp.getFoldersByName(rootFolderName);
  if (rootSearch.hasNext()) {
    rootFolder = rootSearch.next();
  } else {
    rootFolder = DriveApp.createFolder(rootFolderName);
  }
  
  // Country folder
  let countryFolder;
  const countrySearch = rootFolder.getFoldersByName(countryFolderName);
  if (countrySearch.hasNext()) {
    countryFolder = countrySearch.next();
  } else {
    countryFolder = rootFolder.createFolder(countryFolderName);
  }
  
  // Campaign type folder
  let typeFolder;
  const typeSearch = countryFolder.getFoldersByName(campaignType);
  if (typeSearch.hasNext()) {
    typeFolder = typeSearch.next();
  } else {
    typeFolder = countryFolder.createFolder(campaignType);
  }
  
  return typeFolder;
}

// ============================================================
// 4. GESTIÓN DE TEMPLATES (CRUD)
// ============================================================

function getTemplatesList(payload) {
  try {
    const { filters, userEmail } = payload;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TW_CONFIG.SHEET_REGISTRY);
    
    if (!sheet) {
      return buildResponse(true, 'No hay templates registrados aún.', { templates: [] });
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return buildResponse(true, 'No hay templates registrados aún.', { templates: [] });
    }

    const headers = data[0];
    const templates = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });

    // Aplicar filtros si existen
    let filtered = templates;
    if (filters?.countryCode) filtered = filtered.filter(t => t.country_code === filters.countryCode);
    if (filters?.campaignType) filtered = filtered.filter(t => t.campaign_type === filters.campaignType);
    if (filters?.status) filtered = filtered.filter(t => t.status === filters.status);

    // Usuarios no-admin solo ven sus propios drafts y todos los activos
    const userRole = getUserRole(userEmail);
    if (userRole === 'editor' || userRole === 'viewer') {
      filtered = filtered.filter(t => 
        t.status === 'active' || 
        (t.created_by === userEmail && t.status !== 'archived')
      );
    }

    return buildResponse(true, `${filtered.length} templates encontrados`, { templates: filtered });
    
  } catch (e) {
    return buildResponse(false, 'Error al obtener templates: ' + e.message);
  }
}

function updateTemplateStatus(payload) {
  try {
    const { templateId, newStatus, comment, userEmail } = payload;
    
    const userRole = getUserRole(userEmail);
    if (!['admin', 'owner'].includes(userRole)) {
      return buildResponse(false, 'Solo los administradores pueden cambiar el estado de los templates.');
    }

    const validTransitions = {
      'pending_review': ['active', 'rejected'],
      'active': ['archived'],
      'rejected': ['pending_review'],
      'archived': ['active']
    };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TW_CONFIG.SHEET_REGISTRY);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idCol = headers.indexOf('template_id');
    const statusCol = headers.indexOf('status');
    const reviewedByCol = headers.indexOf('reviewed_by');
    const reviewedAtCol = headers.indexOf('reviewed_at');
    const reviewCommentCol = headers.indexOf('review_comment');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === templateId) {
        const currentStatus = data[i][statusCol];
        
        if (!validTransitions[currentStatus]?.includes(newStatus)) {
          return buildResponse(false, `No se puede pasar de "${currentStatus}" a "${newStatus}".`);
        }
        
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        if (reviewedByCol >= 0) sheet.getRange(i + 1, reviewedByCol + 1).setValue(userEmail);
        if (reviewedAtCol >= 0) sheet.getRange(i + 1, reviewedAtCol + 1).setValue(new Date().toISOString());
        if (reviewCommentCol >= 0 && comment) sheet.getRange(i + 1, reviewCommentCol + 1).setValue(comment);
        
        logAuditEvent('tw_status_change', userEmail, { templateId, from: currentStatus, to: newStatus, comment });
        
        // Notificar al creador
        const creatorCol = headers.indexOf('created_by');
        if (creatorCol >= 0 && data[i][creatorCol]) {
          notifyCreatorStatusChange(data[i][creatorCol], templateId, newStatus, comment);
        }
        
        return buildResponse(true, `Template actualizado a "${newStatus}".`);
      }
    }
    
    return buildResponse(false, 'Template no encontrado.');
    
  } catch (e) {
    return buildResponse(false, 'Error al actualizar estado: ' + e.message);
  }
}

// ============================================================
// 5. REGISTRO EN SHEETS
// ============================================================

function registerTemplate(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TW_CONFIG.SHEET_REGISTRY);
  
  // Crear hoja si no existe
  if (!sheet) {
    sheet = ss.insertSheet(TW_CONFIG.SHEET_REGISTRY);
    const headers = [
      'template_id', 'country_code', 'campaign_type', 'doc_id', 'doc_url',
      'version', 'status', 'notes', 'placeholder_count',
      'created_by', 'created_at', 'reviewed_by', 'reviewed_at', 'review_comment'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#FF4500').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  
  const templateId = `TW_${data.countryCode}_${data.campaignType.replace(/\s/g,'')}_${Date.now()}`;
  
  sheet.appendRow([
    templateId,
    data.countryCode,
    data.campaignType,
    data.docId,
    data.docUrl,
    data.version,
    data.status,
    data.notes,
    data.placeholderCount,
    data.createdBy,
    new Date().toISOString(),
    '', '', ''
  ]);
  
  return templateId;
}

function registerTemplateFields(mappings, metadata) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TW_CONFIG.SHEET_FIELDS);
  
  if (!sheet) {
    sheet = ss.insertSheet(TW_CONFIG.SHEET_FIELDS);
    const headers = [
      'field_id', 'placeholder', 'label', 'country_code', 
      'campaign_type', 'field_type', 'required', 'created_at'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#FF4500').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  mappings.forEach(mapping => {
    const knownField = TW_CONFIG.KNOWN_PLACEHOLDERS[mapping.placeholder];
    sheet.appendRow([
      `FIELD_${metadata.countryCode}_${mapping.placeholder}`,
      mapping.placeholder,
      mapping.label || knownField?.label || mapping.placeholder,
      metadata.countryCode,
      metadata.campaignType,
      knownField?.type || 'text',
      true,
      new Date().toISOString()
    ]);
  });
}

// ============================================================
// 6. GESTIÓN DEL EQUIPO ADMIN
// ============================================================

function getAdminTeam(payload) {
  try {
    const { userEmail } = payload;
    const userRole = getUserRole(userEmail);
    
    if (!['admin', 'owner'].includes(userRole)) {
      return buildResponse(false, 'Sin permisos para ver el equipo.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TW_CONFIG.SHEET_TEAM);
    
    if (!sheet) {
      return buildResponse(true, 'Equipo no configurado.', { team: [] });
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return buildResponse(true, '', { team: [] });

    const headers = data[0];
    const team = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });

    return buildResponse(true, '', { team });
    
  } catch (e) {
    return buildResponse(false, 'Error: ' + e.message);
  }
}

function addTeamMember(payload) {
  try {
    const { email, role, name, userEmail } = payload;
    
    const callerRole = getUserRole(userEmail);
    if (callerRole !== 'owner') {
      return buildResponse(false, 'Solo el Owner puede agregar miembros.');
    }
    
    if (!email.endsWith('@rappi.com')) {
      return buildResponse(false, 'Solo se permiten emails @rappi.com.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(TW_CONFIG.SHEET_TEAM);
    
    if (!sheet) {
      sheet = ss.insertSheet(TW_CONFIG.SHEET_TEAM);
      sheet.getRange(1, 1, 1, 5).setValues([['email', 'name', 'role', 'added_by', 'added_at']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#FF4500').setFontColor('#FFFFFF');
    }

    sheet.appendRow([email, name || '', role, userEmail, new Date().toISOString()]);
    logAuditEvent('tw_add_member', userEmail, { newMember: email, role });
    
    return buildResponse(true, `${email} agregado como ${role}.`);
    
  } catch (e) {
    return buildResponse(false, 'Error: ' + e.message);
  }
}

function removeTeamMember(payload) {
  try {
    const { email, userEmail } = payload;
    
    const callerRole = getUserRole(userEmail);
    if (callerRole !== 'owner') {
      return buildResponse(false, 'Solo el Owner puede eliminar miembros.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TW_CONFIG.SHEET_TEAM);
    if (!sheet) return buildResponse(false, 'Hoja de equipo no encontrada.');

    const data = sheet.getDataRange().getValues();
    const emailCol = data[0].indexOf('email');
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][emailCol] === email) {
        sheet.deleteRow(i + 1);
        logAuditEvent('tw_remove_member', userEmail, { removedMember: email });
        return buildResponse(true, `${email} eliminado del equipo.`);
      }
    }
    
    return buildResponse(false, 'Miembro no encontrado.');
    
  } catch (e) {
    return buildResponse(false, 'Error: ' + e.message);
  }
}

// ============================================================
// 7. UTILIDADES
// ============================================================

function getUserRole(email) {
  if (!email) return 'none';
  
  // Owners hardcodeados (máxima seguridad)
  if (TW_CONFIG.ADMIN_EMAILS.includes(email)) return 'owner';
  
  // Buscar en la hoja de equipo
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TW_CONFIG.SHEET_TEAM);
    if (!sheet) return 'none';
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('email');
    const roleCol = headers.indexOf('role');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol] === email) {
        return data[i][roleCol] || 'viewer';
      }
    }
  } catch (e) {
    Logger.log('Error getUserRole: ' + e.message);
  }
  
  return 'none';
}

function logAuditEvent(action, userEmail, details) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(TW_CONFIG.SHEET_AUDIT);
    
    if (!sheet) {
      sheet = ss.insertSheet(TW_CONFIG.SHEET_AUDIT);
      sheet.getRange(1, 1, 1, 5).setValues([['timestamp', 'action', 'user_email', 'details', 'session_id']]);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1a1a2e').setFontColor('#FFFFFF');
    }
    
    sheet.appendRow([
      new Date().toISOString(),
      action,
      userEmail || 'anonymous',
      JSON.stringify(details || {}),
      Utilities.getUuid()
    ]);
  } catch (e) {
    Logger.log('Error logAuditEvent: ' + e.message);
  }
}

function notifyAdminsForReview(data) {
  try {
    const subject = `[RappiMind] Nuevo template pendiente de revisión - ${data.countryCode} / ${data.campaignType}`;
    const body = `
Hola,

${data.userEmail} ha creado un nuevo template de T&C que requiere tu revisión.

Detalles:
• País: ${data.countryCode}
• Tipo: ${data.campaignType}
• Versión: ${data.version || '1.0'}
• Creado por: ${data.userEmail}

Ver documento: ${data.docUrl}

Para aprobar o rechazar, ingresa al Panel de Administración de RappiMind.

—RappiMind by Legal
    `.trim();

    TW_CONFIG.ADMIN_EMAILS.forEach(adminEmail => {
      GmailApp.sendEmail(adminEmail, subject, body);
    });
  } catch (e) {
    Logger.log('Error notifyAdminsForReview: ' + e.message);
  }
}

function notifyCreatorStatusChange(creatorEmail, templateId, newStatus, comment) {
  try {
    const statusLabels = {
      'active': '✅ Aprobado y activado',
      'rejected': '❌ Rechazado',
      'archived': '📦 Archivado'
    };
    
    const subject = `[RappiMind] Tu template fue ${statusLabels[newStatus] || newStatus}`;
    const body = `
Hola,

Tu template (ID: ${templateId}) ha sido actualizado.

Estado: ${statusLabels[newStatus] || newStatus}
${comment ? `\nComentario del revisor:\n"${comment}"` : ''}

${newStatus === 'active' ? 'El template ya está disponible para generación de T&C.' : ''}
${newStatus === 'rejected' ? 'Puedes corregirlo y volver a enviarlo para revisión.' : ''}

—RappiMind by Legal
    `.trim();
    
    GmailApp.sendEmail(creatorEmail, subject, body);
  } catch (e) {
    Logger.log('Error notifyCreatorStatusChange: ' + e.message);
  }
}

function buildResponse(success, message, data) {
  // Retorna objeto plano compatible con google.script.run
  // (ContentService solo funciona con doGet/doPost, no con google.script.run)
  return { success, message, data: data || null };
}

// ============================================================
// INSTRUCCIONES DE INTEGRACIÓN
// ============================================================
//
// Tu sistema usa google.script.run (NO doGet/doPost).
// Por eso NO necesitas agregar nada a doPost ni doGet.
//
// SOLO HAY 2 PASOS:
//
// PASO 1: Pega este archivo al FINAL de tu Código.gs
//         (después de la última función existente)
//         Sin borrar nada, sin modificar nada.
//
// PASO 2: Configura la Gemini API Key (opcional pero recomendado):
//         Apps Script → Configuración del proyecto → Propiedades de script
//         Clave: GEMINI_API_KEY
//         Valor: tu API key (obtenla en https://aistudio.google.com/apikey)
//
//         Sin la key, el wizard usa análisis básico por regex.
//         Con la key, Gemini detecta automáticamente todos los campos.
//
// ¡Listo! El frontend usa google.script.run.analyzeTextForPlaceholders(data)
// etc., igual que el resto de RappiMind.
//
// ============================================================
