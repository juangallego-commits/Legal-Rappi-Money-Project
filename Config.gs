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
const BASE_FIELD_MAP = {
  'NOMBRE_CAMPANA':  { canonical: 'campaignName', format_as: '' },
  'FECHA_INICIO':    { canonical: 'startDate',    format_as: 'date_legal' },
  'FECHA_FIN':       { canonical: 'endDate',      format_as: 'date_legal' },
  'HORA_INICIO':     { canonical: 'startTime',    format_as: '' },
  'HORA_FIN':        { canonical: 'endTime',      format_as: '' },
  'TERRITORIO':      { canonical: 'territory',    format_as: '' },
  'TIENDA_BASE':     { canonical: 'shopName',     format_as: '' },
};
