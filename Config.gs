// =================================================================
// CONFIGURACIÓN GLOBAL Y CONSTANTES
// =================================================================

// 1. Correos y Carpetas
const LEGAL_AUDIT_EMAIL = ['juan.gallego@rappi.com'];
const ADMIN_EMAILS_LIST = ['juan.gallego@rappi.com'];
const DRIVE_FOLDER_ID = ''; 
const TEMPLATES_ROOT_NAME = 'RappiMind_Templates';

// 2. Base de Datos (Google Sheets)
const AUDIT_SHEET_ID = '1Ki9FvHGkGSxnUpZCM2RwieTZwkpIlcBxPIYnvLixqZI';
const REGISTRY_SHEET_NAME = 'Template_Registry';
const FIELDS_SHEET_NAME = 'Template_Fields';
const TEAM_SHEET_NAME = 'Admin_Team';
const APPROVAL_LOG_SHEET = 'Approval_Log';
const CAMPAIGN_TYPES_SHEET = 'Campaign_Types';
const COUNTRY_SETTINGS_SHEET = 'Country_Settings';

// 3. Diccionarios de formato
const MESES_ES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const LISTA_MUNICIPIOS = ['Soacha', 'Chía', 'Cajicá', 'Palmira', 'Bello', 'Buga', 'Envigado', 'Itagüí', 'Sabaneta', 'Jamundí', 'Yumbo', 'Floridablanca', 'Girón', 'Piedecuesta', 'Rionegro', 'Dosquebradas'];

// 4. Roles y Permisos: owner > admin > editor > viewer
const ROLE_HIERARCHY = { owner: 4, admin: 3, editor: 2, viewer: 1 };

// 5. Configuración para el Template Wizard (V3)
const TW_CONFIG = {
  SHEET_REGISTRY: REGISTRY_SHEET_NAME,
  SHEET_FIELDS: FIELDS_SHEET_NAME,
  SHEET_AUDIT: 'Audit_Log',
  SHEET_TEAM: TEAM_SHEET_NAME,
  DRIVE_ROOT_FOLDER_NAME: TEMPLATES_ROOT_NAME,
  ADMIN_EMAILS: ADMIN_EMAILS_LIST,
  
  // Mapeo de países para carpetas
  COUNTRY_FOLDERS: {
    'CO': 'CO_Colombia', 'MX': 'MX_México', 'PE': 'PE_Perú',
    'AR': 'AR_Argentina', 'CL': 'CL_Chile', 'EC': 'EC_Ecuador',
    'UY': 'UY_Uruguay', 'CR': 'CR_Costa_Rica', 'BR': 'BR_Brasil'
  },

  // Campos conocidos para sugerencia automática de IA
  KNOWN_PLACEHOLDERS: {
    'FECHA_INICIO': { label: 'Fecha de inicio', type: 'date', pattern: /\d{1,2} de \w+ de \d{4}/ },
    'FECHA_FIN': { label: 'Fecha de fin', type: 'date', pattern: /\d{1,2} de \w+ de \d{4}/ },
    'FECHA_PUBLICACION': { label: 'Fecha de publicación', type: 'date' },
    'NOMBRE_CAMPANA': { label: 'Nombre de campaña', type: 'text' },
    'TERRITORIO': { label: 'Territorio', type: 'text' },
    'ORGANIZADOR': { label: 'Organizador legal', type: 'text' },
    'NIT_ORGANIZADOR': { label: 'NIT del organizador', type: 'text' },
    'PORCENTAJE_CASHBACK': { label: '% de cashback', type: 'number' },
    'TOPE_NUM': { label: 'Tope máximo (número)', type: 'number' },
    'TOPE_TEXTO': { label: 'Tope máximo (escrito)', type: 'text' },
    'MONEDA': { label: 'Moneda', type: 'text' },
    'MONTO_MINIMO': { label: 'Monto mínimo de compra', type: 'number' },
    'PREMIO_DESCRIPCION': { label: 'Descripción del premio', type: 'text' },
    'PREMIO_VALOR': { label: 'Valor del premio', type: 'number' },
    'NUMERO_GANADORES': { label: 'Número de ganadores', type: 'number' },
    'CRITERIO_SELECCION': { label: 'Criterio de selección', type: 'text' },
    'TOP_N': { label: 'Top N (ej: 10)', type: 'number' },
    'JURISDICCION': { label: 'Jurisdicción legal', type: 'text' },
    'LEY_APLICABLE': { label: 'Ley aplicable', type: 'text' },
    'ENTIDAD_VIGILANCIA': { label: 'Entidad de vigilancia', type: 'text' },
    'CATEGORIAS_PARTICIPANTES': { label: 'Categorías de participantes', type: 'text' },
    'VIGENCIA_DIAS': { label: 'Vigencia en días', type: 'number' },
    'URL_BASES': { label: 'URL de bases legales', type: 'url' },
  }
};
// =================================================================
// V3.3: CAMPOS DERIVADOS — Se calculan a partir del payload
// Usados por _generateSmartTemplate() cuando el template tiene
// placeholders compuestos que no son un campo directo del formulario.
// =================================================================
const DERIVED_FIELDS = {
  // --- Nombres derivados ---
  'NOMBRE_CAMPANA_UPPER': function(p) {
    return (p.campaignName || p.shopName || 'CAMPAÑA').toUpperCase();
  },
  'NOMBRE_CAMPANA_LOWER': function(p) {
    return (p.campaignName || p.shopName || 'campaña').toLowerCase();
  },

  // --- Números a letras ---
  'TOPE_LETRAS': function(p) {
    return p.cap ? numeroALetras(Number(p.cap)) : '';
  },
  'PRESUPUESTO_LETRAS': function(p) {
    return p.budget ? numeroALetras(Number(p.budget)) : '';
  },
  'NUM_GANADORES_LETRAS': function(p) {
    return p.numberOfWinners ? numeroALetras(Number(p.numberOfWinners)) : '';
  },

  // --- Umbrales (Cashback) ---
  'UMBRAL_NUM': function(p) {
    var pct = Number(p.cashbackPct || 0);
    var cap = Number(p.cap || 0);
    if (pct > 0) return Math.ceil(cap / (pct / 100)).toLocaleString('es-CO');
    return '';
  },
  'UMBRAL_LETRAS': function(p) {
    var pct = Number(p.cashbackPct || 0);
    var cap = Number(p.cap || 0);
    if (pct > 0) return numeroALetras(Math.ceil(cap / (pct / 100)));
    return '';
  },

  // --- Plurales/gramática ---
  'PLURAL_GANADORES': function(p) {
    return Number(p.numberOfWinners || 1) === 1 ? 'ganador' : 'ganadores';
  },
  'TEXTO_ORDENES': function(p) {
    return Number(p.maxOrders || 1) === 1 ? 'orden' : 'órdenes';
  },

  // --- Tienda (singular vs plural) ---
  'REF_TIENDA': function(p) {
    var s = String(p.shopName || '').toUpperCase();
    return (s.indexOf('TODAS') === 0 || s.indexOf('TODOS') === 0)
      ? 'las Tiendas Participantes' : 'la Tienda Participante';
  },
  'TIENDA_DISPLAY': function(p) {
    var s = String(p.shopName || '').toUpperCase();
    return (s.indexOf('TODAS') === 0 || s.indexOf('TODOS') === 0)
      ? 'Aliados Comerciales' : '"' + capitalize(p.shopName || '') + '"';
  },
  'DEFINICION_TIENDA': function(p) {
    var s = String(p.shopName || '').toUpperCase();
    return (s.indexOf('TODAS') === 0 || s.indexOf('TODOS') === 0)
      ? ' (en adelante las "Tiendas Participantes") '
      : ' (en adelante la "Tienda Participante") ';
  },
  'DECLARACION_TIENDA': function(p) {
    var s = String(p.shopName || '').toUpperCase();
    return (s.indexOf('TODAS') === 0 || s.indexOf('TODOS') === 0)
      ? 'son las Tiendas Participantes' : 'es la Tienda Participante';
  },
  'TITULO_TIENDA': function(p) {
    return 'IV. Tienda Participante: ';
  },

  // --- Texto de porcentaje ---
  'TEXTO_PORCENTAJE': function(p) {
    var pct = Number(p.cashbackPct || 0);
    if (pct > 0) return numeroALetras(Math.floor(pct)) + ' por ciento (' + pct + '%)';
    return '';
  },

  // --- Texto de carga de créditos ---
  'TEXTO_CARGA': function(p) {
    var tipo = p.loadType || 'Inmediatamente (Al finalizar la orden)';
    if (tipo.indexOf('Inmediatamente') >= 0 || tipo.indexOf('inmediata') >= 0) {
      return 'de manera inmediata una vez finalizada la orden';
    }
    if (tipo.indexOf('día siguiente') >= 0 || tipo.indexOf('Al día') >= 0) {
      return 'al día calendario siguiente de haber finalizado la orden';
    }
    // Fecha específica
    var fechaCarga = p.loadDate || '';
    if (fechaCarga) {
      var parts = fechaCarga.split('-');
      if (parts.length === 3) {
        var day = parseInt(parts[2]);
        var monthIdx = parseInt(parts[1]) - 1;
        if (monthIdx >= 0 && monthIdx < 12) {
          return 'el ' + day + ' de ' + MESES_ES[monthIdx] + ' de ' + parts[0];
        }
      }
    }
    return 'en la fecha indicada por Rappi';
  },

  // --- Texto de vigencia de créditos ---
  'TEXTO_VIGENCIA_CREDITOS': function(p) {
    var tipo = p.validityType || 'Por días calendario (Duración)';
    if (tipo.indexOf('días') >= 0 || tipo.indexOf('Duración') >= 0) {
      var dias = p.validityDays || '30';
      return 'tendrán una vigencia de ' + numeroALetras(Number(dias)) + ' (' + dias + ') días calendario contados a partir de su carga';
    }
    // Por fechas específicas
    var iniRed = '';
    var finRed = '';
    if (p.redemptionStart) {
      var ps = p.redemptionStart.split('-');
      if (ps.length === 3) iniRed = parseInt(ps[2]) + ' de ' + MESES_ES[parseInt(ps[1])-1] + ' de ' + ps[0];
    }
    if (p.redemptionEnd) {
      var pe = p.redemptionEnd.split('-');
      if (pe.length === 3) finRed = parseInt(pe[2]) + ' de ' + MESES_ES[parseInt(pe[1])-1] + ' de ' + pe[0];
    }
    iniRed = iniRed || 'el momento en que sean cargados';
    finRed = finRed || '[FECHA FIN PENDIENTE]';
    return 'podrán ser utilizados entre ' + iniRed + ' y el ' + finRed;
  },

  // --- Texto de segmento ---
  'TEXTO_SEGMENTO': function(p) {
    var seg = p.userSegment || 'Todos los usuarios';
    if (seg.indexOf('Pro') >= 0) {
      return 'Pueden participar los Usuarios/Consumidores que tengan activa la suscripción RappiPro y/o RappiPro Black.';
    }
    if (seg.indexOf('Nuevos') >= 0) {
      return 'Campaña válida únicamente para Nuevos Usuarios/Consumidores.';
    }
    return 'Pueden participar todos los Usuarios/Consumidores de la Plataforma Rappi que sean mayores de edad.';
  },

  // --- Texto de método de pago ---
  'TEXTO_METODO_PAGO': function(p) {
    var met = p.paymentMethods || 'Todos excepto Efectivo';
    if (met.indexOf('excepto Efectivo') >= 0 || met.indexOf('excepto efectivo') >= 0) {
      return 'Campaña válida para órdenes pagadas con todos los medios de pago habilitados en la Plataforma Rappi, excepto efectivo. No se obtendrá el Beneficio respecto de órdenes pagadas en efectivo o parcial/totalmente con Créditos.';
    }
    if (met.indexOf('Todos') >= 0) {
      return 'Campaña válida para órdenes pagadas con todos los medios de pago habilitados en la Plataforma Rappi.';
    }
    return 'Campaña válida únicamente para órdenes pagadas con ' + met.replace('Únicamente ', '') + '.';
  },

  // --- Texto de lugar de redención (Cashback) ---
  'TEXTO_LUGAR_REDENCION': function(p) {
    var lugar = p.redemptionPlace || 'Únicamente en la Tienda Participante (Brand Credits)';
    var refTienda = DERIVED_FIELDS['REF_TIENDA'](p);
    var suffix = ' únicamente dentro del Territorio.';
    if (lugar.indexOf('Restaurantes') >= 0) {
      return 'Se aclara que los Créditos otorgados podrán ser redimidos en cualquier tienda de la sección "Restaurantes"' + suffix;
    }
    if (lugar.indexOf('Plataforma') >= 0 || lugar.indexOf('Generales') >= 0) {
      return 'Se aclara que los Créditos otorgados podrán ser redimidos en cualquier sección de la Plataforma Rappi (excepto Cajero ATM y RappiFavor)' + suffix;
    }
    return 'Se aclara que los Créditos otorgados podrán ser redimidos únicamente en ' + refTienda + ' donde se originó el Beneficio' + suffix;
  },

  // --- Texto de territorio (multi-país V3.4) ---
  'TEXTO_TERRITORIO': function(p) {
    var raw = p.territory || 'Nacional';
    var territorios = raw.split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s.length > 0; });

    // Resolver nombre legal del país desde Country_Settings
    var paisLegal = 'la República de Colombia'; // fallback
    try {
      var csSheet = _getSheet(COUNTRY_SETTINGS_SHEET);
      if (csSheet) {
        var csData = _sheetToObjects(csSheet);
        var cc = p.countryCode || 'CO';
        var csRow = csData.find(function(c) { return c.country_code === cc; });
        if (csRow && csRow.legal_country) {
          paisLegal = csRow.legal_country;
        }
      }
    } catch(e) { /* fallback a Colombia */ }

    if (territorios.indexOf('Nacional') >= 0 || territorios.indexOf('nacional') >= 0) {
      return 'el territorio nacional de ' + paisLegal;
    }

    // Formatear lista
    var listaStr = territorios.length === 1 ? territorios[0]
      : territorios.slice(0, -1).join(', ') + ' y ' + territorios[territorios.length - 1];

    var hayMunicipios = false;
    var hayCiudades = false;
    territorios.forEach(function(lugar) {
      if (typeof LISTA_MUNICIPIOS !== 'undefined' && LISTA_MUNICIPIOS.indexOf(lugar) >= 0) {
        hayMunicipios = true;
      } else {
        hayCiudades = true;
      }
    });

    var prefix = '';
    if (hayCiudades && hayMunicipios) {
      prefix = 'las ciudades y municipios de';
    } else if (hayMunicipios && !hayCiudades) {
      prefix = (territorios.length > 1) ? 'los municipios de' : 'el municipio de';
    } else {
      prefix = (territorios.length > 1) ? 'las ciudades de' : 'la ciudad de';
    }
    return prefix + ' ' + listaStr + ', dentro de ' + paisLegal;
  },

  // --- Fechas formateadas ---
  'FECHA_INICIO': function(p) {
    return _derivedFormatDate(p.startDate);
  },
  'FECHA_FIN': function(p) {
    return _derivedFormatDate(p.endDate);
  },
  'HORA_INICIO': function(p) {
    return _derivedFormatTime(p.startTime || '00:00');
  },
  'HORA_FIN': function(p) {
    return _derivedFormatTime(p.endTime || '23:59');
  },
  'FECHA_ANUNCIO': function(p) {
    return _derivedFormatDate(p.announcementDate);
  },

  // --- Presupuesto formateado ---
  'PRESUPUESTO_NUM': function(p) {
    var n = Number(p.budget || 0);
    return n > 0 ? n.toLocaleString('es-CO') : '';
  },
  'TOPE_NUM': function(p) {
    var n = Number(p.cap || 0);
    return n > 0 ? n.toLocaleString('es-CO') : '';
  },

  // --- Concurso: campos derivados ---
  'CRITERIO_GANADOR': function(p) {
    var crit = p.winnerCriteria || '';
    if (crit.indexOf('$') >= 0 || crit.indexOf('Venta') >= 0 || crit.indexOf('valor') >= 0) {
      return 'mayor valor acumulado (dinero) en compras';
    }
    return 'mayor cantidad de órdenes finalizadas';
  },
  'NUM_GANADORES': function(p) {
    return String(p.numberOfWinners || 1);
  },
  'LISTA_PREMIOS': function(p) {
    var tipo = p.prizeType || 'credits';
    if (tipo === 'credits') {
      var monto = Number(p.creditsAmount || 50000);
      return numeroALetras(monto) + ' (' + monto.toLocaleString('es-CO') + ') Créditos de la App por cada ganador.';
    }
    return p.physicalPrizeDescription || p.prizes || 'Premio sorpresa.';
  },
  'RESPONSABLE_ENTREGA': function(p) {
    return (p.prizeDeliveryBy === 'organizer') ? 'el Organizador' : 'Rappi';
  },
  'MINIMO_COMPRA_TEXTO': function(p) {
    var min = p.minParticipation || '';
    if (min && min !== '0') return '$' + min + ' M/Cte';
    return 'No aplica un valor mínimo.';
  },
  'MONTO_CREDITOS_LETRAS': function(p) {
    return p.creditsAmount ? numeroALetras(Number(p.creditsAmount)) : '';
  },
  'MONTO_CREDITOS_NUM': function(p) {
    return p.creditsAmount ? Number(p.creditsAmount).toLocaleString('es-CO') : '';
  },
  'DIAS_CARGA_LETRAS': function(p) {
    return p.creditLoadDays ? numeroALetras(Number(p.creditLoadDays)) : '';
  },
  'DIAS_CARGA_NUM': function(p) {
    return String(p.creditLoadDays || 5);
  },
  'VIGENCIA_CREDITOS_LETRAS': function(p) {
    return p.creditsValidityDays ? numeroALetras(Number(p.creditsValidityDays)) : '';
  },
  'VIGENCIA_CREDITOS_NUM': function(p) {
    return String(p.creditsValidityDays || 30);
  },
  'LUGAR_REDENCION_PREMIO': function(p) {
    var lugar = p.creditsRedemptionPlace || 'Únicamente en la Tienda Participante';
    if (lugar.indexOf('Restaurantes') >= 0) return 'en cualquier tienda de la sección "Restaurantes"';
    if (lugar.indexOf('Plataforma') >= 0) return 'en cualquier sección de la Plataforma Rappi';
    return 'únicamente en la Tienda Participante "' + capitalize(p.shopName || '') + '"';
  },
  'DESEMPATE_1': function(p) {
    var crit = p.winnerCriteria || '';
    if (crit.indexOf('$') >= 0 || crit.indexOf('Venta') >= 0) {
      return 'Mayor cantidad de órdenes finalizadas.';
    }
    return 'Mayor valor acumulado en compras.';
  },

  // --- Campos de formato base ---
  'TIENDA_BASE': function(p) {
    return capitalize(String(p.shopName || '').trim());
  },
  'NOMBRE_CAMPANA': function(p) {
    return p.campaignName || '';
  },
  'LIMITE_ORDENES': function(p) {
    return p.maxOrders || '1';
  },
  'CONDICIONES_ESPECIALES': function(p) {
    return p.specialConditions || '';
  }
};

// Helper para DERIVED_FIELDS: formatear fecha YYYY-MM-DD → "15 de marzo de 2026"
function _derivedFormatDate(dateStr) {
  if (!dateStr) return '';
  var parts = String(dateStr).split('-');
  if (parts.length < 3) return dateStr;
  var day = parseInt(parts[2]);
  var monthIdx = parseInt(parts[1]) - 1;
  if (monthIdx >= 0 && monthIdx < 12) {
    return day + ' de ' + MESES_ES[monthIdx] + ' de ' + parts[0];
  }
  return dateStr;
}

// Helper para DERIVED_FIELDS: formatear HH:MM → "12:00 p.m."
function _derivedFormatTime(timeStr) {
  if (!timeStr) return '';
  var parts = String(timeStr).split(':');
  var hours = parseInt(parts[0] || 0);
  var minutes = parseInt(parts[1] || 0);
  var ampm = hours >= 12 ? 'p.m.' : 'a.m.';
  hours = hours % 12;
  hours = hours ? hours : 12;
  var minutesStr = minutes < 10 ? '0' + minutes : String(minutes);
  return hours + ':' + minutesStr + ' ' + ampm;
}
// =================================================================
// V3.3: LEGAL DEFAULTS — Se resuelven automáticamente de Country_Settings
// El KAM nunca ve estos campos como preguntas.
// =================================================================
const LEGAL_DEFAULTS_MAP = {
  'JURISDICCION':              { column: 'jurisdiction_text' },
  'LEY_APLICABLE':             { column: 'applicable_law' },
  'ENTIDAD_VIGILANCIA':        { column: 'legal_entity' },
  'MONEDA_TEXTO':              { column: 'currency_name' },
  'PAIS_LEGAL':                { column: 'legal_country' },
  'URL_BASES':                 { column: 'legal_url' },
  // V3.4: Nuevos campos para modelo global de T&C
  'ENTIDAD_LEGAL':             { column: 'entidad_legal_nombre' },
  'ID_FISCAL':                 { column: 'id_fiscal' },
  'DOMICILIO_LEGAL':           { column: 'domicilio_legal' },
  'URL_PRIVACIDAD':            { column: 'url_privacidad' },
  'URL_TC_CREDITOS':           { column: 'url_tc_creditos' },
  'URL_TC_PLATAFORMA':         { column: 'url_tc_plataforma' },
  'AUTORIDAD_DATOS':           { column: 'autoridad_datos' },
  'NOMBRE_POLITICA_PRIVACIDAD': { column: 'nombre_politica_privacidad' },
  'EDAD_MINIMA':               { column: 'edad_minima' }
};
const BASE_FIELD_MAP = {
  'NOMBRE_CAMPANA':  { canonical: 'campaignName', format_as: '' },
  'FECHA_INICIO':    { canonical: 'startDate',    format_as: 'date_legal' },
  'FECHA_FIN':       { canonical: 'endDate',      format_as: 'date_legal' },
  'HORA_INICIO':     { canonical: 'startTime',    format_as: '' },
  'HORA_FIN':        { canonical: 'endTime',      format_as: '' },
  'TERRITORIO':      { canonical: 'territory',    format_as: '' },
  'TIENDA_BASE':     { canonical: 'shopName',     format_as: '' },
};
